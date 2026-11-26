// electron-builder afterPack hook. The NSIS installer copies files with a 260-char
// path limit: install or temp folder (50 chars + user name) + relative path.
// A longer path makes the copy fail, which the installer reports as
// "Event Automator cannot be closed". Fail the build instead.
const fs = require('fs');
const path = require('path');

const MAX = 170; // fits user names up to ~39 chars; the anthropic SDK alone needs 161

function walk(dir, root, out) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, root, out);
    else if (path.relative(root, p).length > MAX) out.push(path.relative(root, p));
  }
  return out;
}

module.exports = async function (context) {
  const long = walk(context.appOutDir, context.appOutDir, []);
  if (long.length) {
    throw new Error(`${long.length} packed path(s) over ${MAX} chars break the installer's copy step:\n  `
      + long.slice(0, 10).map((p) => `${p.length}  ${p}`).join('\n  '));
  }
};
