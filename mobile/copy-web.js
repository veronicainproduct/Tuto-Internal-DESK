// Copies the web app into www/ so the Android app ships the exact same code.
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..', 'app', 'public');
const out = path.join(__dirname, 'www');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(path.join(out, 'icons'), { recursive: true });

for (const f of fs.readdirSync(root)) if (f.endsWith('.html') || f.endsWith('.js') || f.endsWith('.webmanifest')) fs.copyFileSync(path.join(root, f), path.join(out, f));
for (const f of fs.readdirSync(path.join(root, 'icons'))) {
  fs.copyFileSync(path.join(root, 'icons', f), path.join(out, 'icons', f));
}

console.log('Copied web app to mobile/www');
