// What counts as a One-Pager, and how to file one without destroying another.
//
// Extracted from main.js because the interesting part cannot be reached there:
// the upload begins with a native file dialog, which no automated driver can
// click, so anything downstream of it was only ever verifiable by inspection.
// The two rules that actually matter - which formats are usable, and how a name
// clash is resolved - are pure functions of their inputs and belong somewhere a
// test can call them.
//
// The clash rule earns that scrutiny: the One-Pager folder is shared, so the
// file already sitting at a name may be a colleague's, and getting this wrong
// overwrites work that is not yours to overwrite.

// .md and .txt are read straight off disk. .docx and .pdf are converted to
// Markdown first by doc_convert.py, which writes the .md beside the original -
// so every format is usable, and everything downstream still only ever sees
// Markdown. The original stays where it is, because it is the file the studio
// actually signed off and someone will want to open it.
const READABLE = ['md', 'txt'];
const CONVERTIBLE = ['docx', 'pdf'];
const STORABLE = READABLE.concat(CONVERTIBLE);

function extOf(name) {
  const i = String(name || '').lastIndexOf('.');
  return i <= 0 ? '' : String(name).slice(i + 1).toLowerCase();
}

function kindOf(name) {
  const ext = extOf(name);
  const readable = READABLE.includes(ext);
  const convertible = CONVERTIBLE.includes(ext);
  // `usable` is what the UI should offer a button for. `readable` is only
  // whether it can be read without converting first.
  return { ext, readable, convertible, usable: readable || convertible,
           storable: readable || convertible };
}

// Where a converted .docx or .pdf lands: the same name with a .md extension.
function markdownSiblingOf(filePath) {
  const i = filePath.lastIndexOf('.');
  return (i <= 0 ? filePath : filePath.slice(0, i)) + '.md';
}

// The filename to use when "Keep both" is chosen: the first "name (n).ext" that
// is free. `exists` is injected so this is testable without a filesystem, and so
// the caller decides what "already there" means.
//
// Counting from 2 rather than 1 matches what Windows Explorer does, which is
// what a designer looking at the folder will expect to see.
function nextFreeName(baseName, ext, exists) {
  const dot = ext ? '.' + ext : '';
  if (!exists(baseName + dot)) return baseName + dot;
  let n = 2;
  // Bounded so a pathological folder cannot spin forever; 999 copies of one
  // One-Pager is a different problem than this function can solve.
  while (n < 1000 && exists(`${baseName} (${n})${dot}`)) n++;
  return `${baseName} (${n})${dot}`;
}

module.exports = { READABLE, CONVERTIBLE, STORABLE, extOf, kindOf,
                   markdownSiblingOf, nextFreeName };
