#!/usr/bin/env python3
"""Inject PowerPoint entrance animations into a pptxgenjs-built deck.

usage: add_animations.py deck.pptx out.pptx '<shape-name>[,<shape-name>...]' [--slide N] [--effect zoom|fly] [--pulse]

Each comma-separated group of shape names becomes ONE animation step (the shapes in the group
animate together). The first group starts on click; later groups start "after previous" with a
short delay. Optional --pulse adds an emphasis pulse after each group's entrance.
"""
import sys, re, zipfile, shutil, argparse
from defusedxml import minidom as dminidom
from xml.dom import minidom

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

def find_spids(slide_xml, names):
    """Return list of (name, spid) for shapes whose <p:cNvPr name=...> matches a name."""
    out = []
    for n in names:
        m = re.search(r'<p:cNvPr id="(\d+)" name="%s"' % re.escape(n), slide_xml)
        if not m:
            # pptxgenjs escapes & as &amp; in names
            m = re.search(r'<p:cNvPr id="(\d+)" name="%s"' % re.escape(n.replace('&', '&amp;')), slide_xml)
        if not m:
            raise SystemExit(f'shape named {n!r} not found in slide')
        out.append((n, m.group(1)))
    return out

class Ids:
    def __init__(self, start=3):
        self.n = start
    def next(self):
        self.n += 1
        return self.n

def zoom_entrance(ids, spid, node_type, delay_ms):
    c = ids.next()
    return f'''<p:par><p:cTn id="{c}" presetID="53" presetClass="entr" presetSubtype="16" fill="hold" grpId="0" nodeType="{node_type}"><p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{ids.next()}" dur="600" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav></p:tavLst></p:anim><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{ids.next()}" dur="600" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav></p:tavLst></p:anim><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{ids.next()}" dur="600"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>'''

def pulse(ids, spid, node_type, delay_ms):
    c = ids.next()
    return f'''<p:par><p:cTn id="{c}" presetID="26" presetClass="emph" presetSubtype="0" fill="hold" grpId="1" nodeType="{node_type}"><p:stCondLst><p:cond delay="{delay_ms}"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr><p:cTn id="{ids.next()}" dur="400" autoRev="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr><p:by x="108000" y="108000"/></p:animScale></p:childTnLst></p:cTn></p:par>'''

def build_timing(groups, do_pulse):
    ids = Ids(2)  # tmRoot=1, mainSeq=2
    ENT, PUL, GAP = 600, 400, 250
    pars = []   # sibling <p:par> under the single click par, each with its own delay
    bld = []
    t = 0
    for gi, spids in enumerate(groups):
        ent = []
        for si, spid in enumerate(spids):
            nt = ('clickEffect' if gi == 0 else 'afterEffect') if si == 0 else 'withEffect'
            ent.append(zoom_entrance(ids, spid, nt, 0))
            bld.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        pars.append(f'<p:par><p:cTn id="{ids.next()}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst><p:childTnLst>{"".join(ent)}</p:childTnLst></p:cTn></p:par>')
        t += ENT
        if do_pulse:
            pul = []
            for si, spid in enumerate(spids):
                pul.append(pulse(ids, spid, 'afterEffect' if si == 0 else 'withEffect', 0))
                bld.append(f'<p:bldP spid="{spid}" grpId="1" animBg="1"/>')
            pars.append(f'<p:par><p:cTn id="{ids.next()}" fill="hold"><p:stCondLst><p:cond delay="{t}"/></p:stCondLst><p:childTnLst>{"".join(pul)}</p:childTnLst></p:cTn></p:par>')
            t += PUL * 2  # autoRev doubles the duration
        t += GAP
    click_par = (f'<p:par><p:cTn id="{ids.next()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
                 f'<p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn></p:par>')
    timing = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
              '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
              + click_par +
              '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
              '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
              '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>' + ''.join(bld) + '</p:bldLst></p:timing>')
    return timing

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('dst'); ap.add_argument('groups', help='semicolon-separated groups of comma-separated shape names')
    ap.add_argument('--slide', type=int, default=1)
    ap.add_argument('--pulse', action='store_true')
    a = ap.parse_args()
    groups_names = [[n.strip() for n in g.split(',') if n.strip()] for g in a.groups.split(';') if g.strip()]
    part = f'ppt/slides/slide{a.slide}.xml'
    zin = zipfile.ZipFile(a.src)
    xml = zin.read(part).decode('utf-8')
    groups = [[spid for _, spid in find_spids(xml, names)] for names in groups_names]
    timing = build_timing(groups, a.pulse)
    # sanity: well-formed
    dminidom.parseString('<r xmlns:p="%s">%s</r>' % (P, timing))
    if '<p:timing>' in xml:
        xml = re.sub(r'<p:timing>.*?</p:timing>', '', xml, flags=re.S)
    # p:timing goes after p:clrMapOvr / p:transition, before p:extLst, i.e. right before </p:sld>
    if '<p:extLst' in xml.split('</p:cSld>')[-1]:
        xml = xml.replace('<p:extLst', timing + '<p:extLst', 1) if xml.rfind('<p:extLst') > xml.rfind('</p:cSld>') else xml.replace('</p:sld>', timing + '</p:sld>')
    else:
        xml = xml.replace('</p:sld>', timing + '</p:sld>')
    zout = zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == part:
            data = xml.encode('utf-8')
        zout.writestr(item, data)
    zout.close()
    print(f'animated groups {groups} -> {a.dst}')

if __name__ == '__main__':
    main()
