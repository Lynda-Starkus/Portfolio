// Shared helpers for the Immersia slide builders.
// Usage: const { iconData } = require("./helpers.js")  (react-icons rendu en PNG via sharp)
// Dépendances : react, react-dom, react-icons, sharp (voir package.json)
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');

const SETS = {
  fa: () => require('react-icons/fa'),
  fa6: () => require('react-icons/fa6'),
  md: () => require('react-icons/md'),
  hi: () => require('react-icons/hi2'),
  tb: () => require('react-icons/tb'),
  lu: () => require('react-icons/lu'),
  bs: () => require('react-icons/bs'),
  pi: () => require('react-icons/pi'),
};

/** Render a react-icon to a PNG data URI usable in pptxgenjs addImage({ data }).
 *  icon('fa', 'FaUserGraduate', 'FFFFFF') -> "image/png;base64,...."
 *  size: raster size in px (>= 256). color: 6-hex without '#'. */
async function iconData(set, name, color = '1E2761', size = 512) {
  const mod = SETS[set]();
  const Cmp = mod[name];
  if (!Cmp) throw new Error(`icon ${set}/${name} not found`);
  let svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Cmp, { color: '#' + color, size: String(size) }));
  if (!svg.includes('xmlns=')) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  const buf = await sharp(Buffer.from(svg), { density: 300 }).resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

/** Convenience: returns an addImage options fragment { data } */
async function icon(set, name, color, size) {
  return { data: await iconData(set, name, color, size) };
}

module.exports = { iconData, icon, SETS };
