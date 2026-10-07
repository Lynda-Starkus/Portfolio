# Immersia — slide de présentation (commission d'ingénierie pédagogique)

Fichiers :

- `Immersia_slide.pptx` — la slide (16:9, une seule diapositive). Les deux grands chiffres apparaissent en animation : un clic fait surgir « jusqu'à 30 h » (zoom + pulsation), puis « 61 % » enchaîne automatiquement.
- `Immersia_slide.pdf` / `Immersia_slide_preview.png` — rendu statique de contrôle (LibreOffice ; les animations ne s'y voient pas).
- `build/` — générateur `pptxgenjs` (`build.js`), helper d'icônes, script d'injection des animations, logos.

## Sources des deux chiffres

1. **« jusqu'à 30 h »** — Storks, Yu, Ma & Chai, *NLP Reproducibility For All: Understanding Experiences of Beginners*, ACL 2023 (93 étudiants d'un cours d'introduction au NLP, Univ. du Michigan). Le temps d'installation de l'environnement d'un projet déclaré par les étudiants va « de moins d'une heure à près de 30 heures » ; la médiane est systématiquement supérieure aux 2 h des chercheurs experts ; plainte la plus fréquente : la spécification des versions de Python et des packages. https://aclanthology.org/2023.acl-long.568/
2. **« 61 % »** — Velez et al., *Student Adoption and Perceptions of a Web Integrated Development Environment: An Experience Report*, ACM SIGCSE 2020 (enquête, 140 étudiants, UC Davis) : 61 % des répondants citent « No Installation Required » parmi les fonctionnalités les plus utiles (figure 7). https://doi.org/10.1145/3328778.3366949

## Hypothèses à ajuster avant présentation

- Les **dates des jalons** (oct. 2026 → juin 2027) et les **KPIs cibles** sont des propositions : les échanges Teams/Outlook n'étaient pas accessibles depuis l'environnement de génération. Le titre de la timeline porte la mention « dates indicatives » ; la retirer une fois les dates validées.
- Le logo PST&B provient d'un export SVG public (variante monochrome, passée en noir) ; remplacer par le fichier officiel de la charte si nécessaire. Le logo OPCO Atlas est le vectoriel officiel (bleu nuit 2D2D5A / violet 7850DC).

## Regénérer

```bash
cd presentations/immersia/build
npm install                     # une fois (pptxgenjs, react-icons, sharp)
node build.js                   # -> ../Immersia_slide_base.pptx (sans animations)
python3 add_animations.py ../Immersia_slide_base.pptx ../Immersia_slide.pptx \
  'stat1-number,stat1-label;stat2-number,stat2-label' --pulse
```

Tout le texte et la géométrie sont dans `build.js` (palette dans `THEME`, copie dans les blocs commentés).
