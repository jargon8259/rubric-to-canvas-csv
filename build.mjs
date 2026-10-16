// Zero-dependency build. Needs Node 22.13 or newer (uses node:module stripTypeScriptTypes).
// Usage: node build.mjs        Optional env: REPO_URL, VERSION
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { stripTypeScriptTypes } from 'node:module';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = dirname(fileURLToPath(import.meta.url));
const read = (p) => readFileSync(join(root, p), 'utf8');

// Dependency order. All modules are inlined into ONE scope, so top-level names must be unique across files.
// Modules use `import ... from './x.ts'` and `export` on declarations only.
const ORDER = ['text', 'csv', 'zip', 'xml', 'docx', 'html', 'scoring', 'state', 'rba', 'plaintext', 'validate', 'changes', 'export', 'app', 'main'];

function inline(name) {
  let code = stripTypeScriptTypes(read(`src/${name}.ts`), { mode: 'strip' });
  code = code.replace(/^import\s[^;]*?from\s+['"][^'"]+['"];?[ \t]*$/gms, '');
  code = code.replace(/^export\s+(?=(?:async\s+)?(?:function|const|let|class)\b)/gm, '');
  if (/^\s*(import|export)\s/m.test(code)) throw new Error(`Unhandled import/export left in src/${name}.ts`);
  return `// ---- ${name} ----\n${code}`;
}

const script = `'use strict';\n(function () {\n${ORDER.map(inline).join('\n')}\n})();`;
if (script.includes('</script')) throw new Error('Script contains "</script"');

const sampleDocx = join(root, 'samples/sample-rba.docx');
const config = {
  repoUrl: process.env.REPO_URL || undefined,
  version: process.env.VERSION || JSON.parse(read('package.json')).version,
  sampleDocxBase64: existsSync(sampleDocx) ? readFileSync(sampleDocx).toString('base64') : undefined,
  sampleCsv: existsSync(join(root, 'samples/sample-rubric-table.csv')) ? read('samples/sample-rubric-table.csv') : undefined,
};

const html = read('src/index.template.html')
  .replace('/*__STYLE__*/', () => read('src/styles.css'))
  .replace('/*__CONFIG__*/{}', () => JSON.stringify(config).replace(/</g, '\\u003c'))
  .replace('/*__SCRIPT__*/', () => script);

mkdirSync(join(root, 'dist'), { recursive: true });
writeFileSync(join(root, 'dist/rubric-to-canvas.html'), html);
writeFileSync(join(root, 'dist/index.html'), html); // same file, so GitHub Pages serves it at the site root
console.log(`Built dist/rubric-to-canvas.html (${(html.length / 1024).toFixed(1)} KB)`);
