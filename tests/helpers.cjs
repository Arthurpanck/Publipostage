const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

function loadApp(extra = {}) {
    const app = vm.createContext({
        Blob, Uint8Array, ArrayBuffer, TextEncoder, TextDecoder, DOMParser, XMLSerializer,
        console: { log() {}, warn() {}, error() {} }, ...extra,
    });
    app.window = app;
    for (const file of ['inc/pizzip.js', 'inc/docxtemplater.js', 'merge-utils.js', 'relations-tools.js', 'docx-tools.js']) {
        vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), app, { filename: file });
    }
    return app;
}
const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = text => `<w:p>${run(text)}</w:p>`;
function template(app, body, header = '', footer = '') {
    const zip = new app.PizZip();
    const types = [['document', 'document.main'], ...(header ? [['header1', 'header']] : []), ...(footer ? [['footer1', 'footer']] : [])];
    zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${types.map(([name, type]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')}</Types>`);
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/document.xml', `<w:document xmlns:w="${NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
    if (header) zip.file('word/header1.xml', `<w:hdr xmlns:w="${NS}">${header}</w:hdr>`);
    if (footer) zip.file('word/footer1.xml', `<w:ftr xmlns:w="${NS}">${footer}</w:ftr>`);
    return zip.generate({ type: 'uint8array' });
}
async function render(app, buffer, data) {
    const blob = app.generateDocxBlob(data, buffer);
    return new app.PizZip(await blob.arrayBuffer());
}
function xmlText(xml) {
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    return Array.from(doc.getElementsByTagName('w:t')).map(t => t.textContent).join('');
}
module.exports = { loadApp, template, render, xmlText, para, run, NS };
