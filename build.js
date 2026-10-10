// Builds index.html (the IMX editor) by embedding the presentation player into the editor.
// Usage: node build.js
const fs = require('fs');
const path = require('path');
const src = f => fs.readFileSync(path.join(__dirname, 'src', f), 'utf8');

const player = src('player.html');
const editor = src('editor.html');
if (!player.includes('__IMX_DATA__')) throw new Error('player.html is missing the __IMX_DATA__ slot');
if (!editor.includes('__PLAYER__')) throw new Error('editor.html is missing the __PLAYER__ slot');

// "</" is escaped so the embedded player cannot close the editor's <script> tag early.
const literal = JSON.stringify(player).replace(/<\//g, '<\\/').replace(/<!--/g, '<\\!--');
const out = editor.replace('__PLAYER__', () => literal);

fs.writeFileSync(path.join(__dirname, 'index.html'), out);
console.log(`index.html written (${(out.length / 1024).toFixed(0)} KB)`);
