// ONE-PAGER FILING TEST
// The upload begins with a native file dialog, which no automated driver can
// click, so the rules downstream of it are the part worth pinning here.
//
// The clash rule especially: the One-Pager folder is shared, so a file already
// sitting at a name may be a colleague's, and overwriting it silently is the one
// outcome this feature must never produce.
//
//   node tests/onepager_test.js
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { READABLE, CONVERTIBLE, STORABLE, extOf, kindOf,
        markdownSiblingOf, nextFreeName } = require('../lib/onepager');

let failures = 0;
function ok(cond, msg, detail) {
  if (cond) console.log('OK  ' + msg);
  else { console.error('FAIL ' + msg + (detail ? '\n     ' + detail : '')); failures++; }
}

// ---- which formats are usable ---------------------------------------------
ok(kindOf('Merge2610 One-Pager.md').readable, 'markdown can drive a generation');
ok(kindOf('brief.txt').readable, 'so can plain text');
// Word and PDF are converted to Markdown by doc_convert.py on first use, so
// they are usable - just not readable without that step. The distinction is what
// lets the UI show "Convert & use" instead of a dead row.
const pdf = kindOf('Sim2609 One-Pager.pdf');
ok(!pdf.readable && pdf.convertible && pdf.usable,
   'a PDF is not read directly, but is usable via conversion');
const docx = kindOf('draft.docx');
ok(!docx.readable && docx.convertible && docx.usable,
   'and the same for Word - the real shipped One-Pagers are PDFs, so this matters');
ok(kindOf('brief.md').usable && !kindOf('brief.md').convertible,
   'Markdown needs no conversion step');
ok(!kindOf('theme.xlsx').storable, 'an unrelated format is refused outright');
ok(!kindOf('README').storable, 'and so is a file with no extension');
ok(extOf('a.name.with.dots.md') === 'md', 'the extension is the LAST dot, not the first',
   extOf('a.name.with.dots.md'));
ok(extOf('Merge2610 One-Pager.MD') === 'md', 'and case does not matter');
ok(READABLE.every(e => STORABLE.includes(e)) && CONVERTIBLE.every(e => STORABLE.includes(e)),
   'everything readable or convertible is also storable');
ok(!READABLE.some(e => CONVERTIBLE.includes(e)),
   'and no format is both - a file is either read or converted, never ambiguously either');

// A converted file lands beside its original, which is how the picker knows to
// hide the generated .md rather than offering it as a second document.
ok(markdownSiblingOf('C:/x/Merge2608 One-Pager.pdf') === 'C:/x/Merge2608 One-Pager.md',
   'a conversion lands beside its original with a .md extension',
   markdownSiblingOf('C:/x/Merge2608 One-Pager.pdf'));
ok(markdownSiblingOf('C:/x/brief.v2.docx') === 'C:/x/brief.v2.md',
   'and only the final extension is replaced');

// ---- a name clash never overwrites ----------------------------------------
const present = new Set();
const exists = (n) => present.has(n);

ok(nextFreeName('Merge2610 One-Pager', 'md', exists) === 'Merge2610 One-Pager.md',
   'a free name is used as-is');

present.add('Merge2610 One-Pager.md');
ok(nextFreeName('Merge2610 One-Pager', 'md', exists) === 'Merge2610 One-Pager (2).md',
   'a taken name becomes "(2)" - counting from 2 like Explorer, so the folder reads as expected',
   nextFreeName('Merge2610 One-Pager', 'md', exists));

present.add('Merge2610 One-Pager (2).md');
present.add('Merge2610 One-Pager (3).md');
ok(nextFreeName('Merge2610 One-Pager', 'md', exists) === 'Merge2610 One-Pager (4).md',
   'and it skips every copy already filed rather than reusing one');

// The original must still be there afterwards - that is the whole point.
ok(present.has('Merge2610 One-Pager.md'),
   'the file that was already there is never the one written to');

ok(nextFreeName('brief', '', exists) === 'brief',
   'a name with no extension still resolves');

// ---- and it holds against a real filesystem, not only a fake --------------
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'onepager-'));
const realExists = (n) => fs.existsSync(path.join(tmp, n));
fs.writeFileSync(path.join(tmp, 'Shared Brief.md'), 'a colleague wrote this');
const chosen = nextFreeName('Shared Brief', 'md', realExists);
fs.writeFileSync(path.join(tmp, chosen), 'mine');
ok(chosen === 'Shared Brief (2).md', 'against a real folder it picks "(2)"', chosen);
ok(fs.readFileSync(path.join(tmp, 'Shared Brief.md'), 'utf8') === 'a colleague wrote this',
   'and the colleague\'s file is untouched');
fs.rmSync(tmp, { recursive: true, force: true });

console.log('');
if (failures) { console.error(`${failures} ONE-PAGER CHECK(S) FAILED`); process.exit(1); }
console.log('ONE-PAGER FILING OK');
