const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

// Le navigateur garde les fichiers du widget en cache : chaque déploiement
// change le numéro de version (?v=) de toutes les ressources à la fois, pour
// qu'un ancien fichier ne soit jamais chargé avec un nouveau.
test('index.html : les fichiers du widget portent tous le même numéro de version', () => {
    const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
    const local = [...html.matchAll(/(?:src|href)="(?!inc\/|https?:)([^"?]+\.(?:js|css))(?:\?v=([^"]*))?"/g)];
    assert.deepEqual(local.map(m => m[1]).sort(),
        ['docx-tools.js', 'main.js', 'merge-utils.js', 'pdf-tools.js', 'relations-tools.js', 'render-tools.js', 'style.css', 'template-store.js', 'ui.js']);
    const versions = new Set(local.map(m => m[2]));
    assert.equal(versions.size, 1, `versions différentes : ${[...versions].join(', ')}`);
    assert.match([...versions][0] || '', /^\d{8}-\d+$/);
});
