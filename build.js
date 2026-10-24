// Builds index.html: the IMX editor with the style engine, the PPTX exporter,
// the libraries and the blank PowerPoint template inlined, so it works as one file.
// Usage: node build.js
const fs = require('fs');
const path = require('path');
const read = (...p) => fs.readFileSync(path.join(__dirname, ...p), 'utf8');

const editor = read('src', 'editor.html');
const libs = [
  read('vendor', 'jszip.min.js'),
  read('vendor', 'mp4-muxer.js'),
  // gifenc ships as CommonJS: give it an exports object and publish it as window.gifenc
  '(function(){var exports={};\n' + read('vendor', 'gifenc.js') + '\nwindow.gifenc=exports;})();',
  read('src', 'styles.js'),
  read('src', 'pptx.js'),
].join('\n;\n').replace(/<\/script/gi, '<\\/script');
const template = fs.readFileSync(path.join(__dirname, 'src', 'template', 'blank.pptx')).toString('base64');

for (const slot of ['/*__LIBS__*/', '__TEMPLATE__'])
  if (!editor.includes(slot)) throw new Error('editor.html is missing ' + slot);
const out = editor.replace('/*__LIBS__*/', () => libs).replace('__TEMPLATE__', () => template);
fs.writeFileSync(path.join(__dirname, 'index.html'), out);
console.log(`index.html written (${(out.length / 1024).toFixed(0)} KB)`);
