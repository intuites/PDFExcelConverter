// Exercise upload -> parse -> preview -> export against an isolated DOM stub.
// Optional: node test_web_flow.cjs path/to/sample.pdf (local server required).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class Element {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.textContent = '';
    this.hidden = false;
    this.attrs = {};
    this.listeners = {};
    const classes = new Set();
    this.classList = { add: x => classes.add(x), remove: x => classes.delete(x) };
  }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.children = nodes; }
  addEventListener(name, callback) { this.listeners[name] = callback; }
  setAttribute(name, value) { this.attrs[name] = value; }
  querySelectorAll(tag) { return this.children.filter(node => node.tag === tag); }
  scrollIntoView() { this.scrolled = true; }
}

async function main() {
  let statement = {
    info: { 'Account Number': '000123', 'Statement Period': 'January 2025' },
    summary: { 'Beginning Balance': 100, 'Ending Balance': 90 },
    transactions: [['01/01', 'Purchase <untrusted> description', -10, 90]],
    mismatches: [], warnings: [],
  };
  if (process.argv[2]) {
    const response = await fetch('http://127.0.0.1:8000/convert?name=sample.pdf&format=json', {
      method: 'POST', body: fs.readFileSync(process.argv[2]),
    });
    assert.equal(response.status, 200);
    statement = await response.json();
    assert.equal(statement.transactions.length, Number(process.argv[3] || 62));
  }

  const nodes = {};
  let answer;
  const context = vm.createContext({
    document: {
      getElementById: id => nodes[id] ||= new Element(id),
      createElement: tag => new Element(tag),
    },
    Intl, AbortController, setTimeout, URL,
    matchMedia: () => ({ matches: true }),
    fetch: () => new Promise(resolve => { answer = resolve; }),
  });
  vm.runInContext(fs.readFileSync('web/app.js', 'utf8'), context);
  const upload = name => vm.runInContext(`pick({ name: '${name}', size: 1024 })`, context);
  const success = () => answer({ ok: true, json: async () => statement });

  const first = upload('statement.pdf');
  assert.equal(nodes.preview.hidden, false);
  assert.equal(nodes.preview.attrs['aria-busy'], 'true');
  assert.equal(nodes.actions.children.length, 0, 'No export before parse succeeds');
  success();
  await first;
  assert.equal(nodes.preview.attrs['aria-busy'], 'false');
  assert.equal(nodes.rows.children[2].children.length, statement.transactions.length);
  assert.equal(nodes.rows.children[1].children[0].children.length, 4);
  assert.equal(nodes.summary.children[0].textContent, 'Statement Period');
  assert.equal(nodes.actions.children.length, 2);
  assert.equal(nodes.actions.children[0].textContent, 'Export Excel');
  assert.equal(nodes.actions.children[1].textContent, 'Export CSV');
  assert.equal(nodes.preview.scrolled, true);

  const failure = upload('bad.pdf');
  answer({ ok: true, json: async () => ({ transactions: [] }) });
  await failure;
  assert.equal(nodes.preview.hidden, true);
  assert.equal(nodes.actions.children.length, 0);
  assert.match(nodes.status.textContent, /No transactions could be mapped/);

  const old = upload('old.pdf');
  const oldAnswer = answer;
  const latest = upload('latest.pdf');
  success();
  await latest;
  oldAnswer({ ok: true, json: async () => ({ transactions: [] }) });
  await old;
  assert.equal(nodes.preview.hidden, false, 'Stale response must not replace latest preview');
  assert.equal(nodes['file-name'].textContent, 'latest.pdf');
  console.log(`PASS: automatic parsing, ${statement.transactions.length}-row preview, export gating, error state, and selection race`);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
