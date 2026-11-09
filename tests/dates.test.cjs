const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, para, render, xmlText } = require('./helpers.cjs');

const app = loadApp();
const fmt = (value, type, opts = {}, info = { locale: 'fr-FR' }) => app.formatDate(value, { type, opts }, info);

test('date : format de la colonne Grist, qu\'elle arrive en Date (ligne du widget) ou en secondes (fetchTable)', () => {
    assert.equal(fmt(new Date('1990-03-14T00:00:00Z'), 'Date', { dateFormat: 'DD/MM/YYYY' }), '14/03/1990');
    assert.equal(fmt(1790899200, 'Date', { dateFormat: 'DD/MM/YYYY' }), '02/10/2026');
    assert.equal(fmt(1790899200, 'Date', { dateFormat: 'dddd D MMMM YYYY' }), 'vendredi 2 octobre 2026');
    assert.equal(fmt(1790899200, 'Date', { dateFormat: 'D MMM YY [à Lyon]' }), '2 oct. 26 à Lyon');
    assert.equal(fmt(1790899200, 'Date', {}), '2026-10-02', 'format par défaut de Grist');
});

test('date et heure : fuseau de la colonne, heure d\'été, format de l\'heure', () => {
    const paris = 'DateTime:Europe/Paris', o = { dateFormat: 'DD/MM/YYYY', timeFormat: 'HH:mm' };
    assert.equal(fmt(new Date('2026-11-09T09:30:00Z'), paris, o), '09/11/2026 10:30');
    assert.equal(fmt(new Date('2026-07-09T09:30:00Z'), paris, { ...o, timeFormat: 'HH:mm Z' }), '09/07/2026 11:30 +02:00');
    assert.equal(fmt(new Date('2026-11-09T21:30:00Z'), 'DateTime:America/New_York', { dateFormat: 'MMMM Do, YYYY', timeFormat: 'h:mma' }, { locale: 'en-US' }),
        'November 9th, 2026 4:30pm');
    assert.equal(fmt(new Date('2026-11-09T09:30:00Z'), 'DateTime', {}, { locale: 'fr-FR', timezone: 'Europe/Paris' }), '2026-11-09 10:30am',
        'sans fuseau dans le type : fuseau du document');
    assert.equal(fmt(new Date('2026-11-09T09:30:00Z'), 'DateTime:Inconnu/Fuseau', o), '09/11/2026 09:30');
});

test('date vide ou texte invalide : rien ou le texte tel quel', () => {
    assert.equal(fmt(null, 'Date'), '');
    assert.equal(fmt('pas une date', 'Date'), 'pas une date');
});

test('dates dans le corps, le pied de page et une table liée', async () => {
    const info = {
        locale: 'fr-FR', timezone: 'Europe/Paris', byKey: new Map([['visites', { jour: { type: 'Date', opts: { dateFormat: 'DD/MM/YYYY' } } }]]),
        root: {
            naissance: { type: 'Date', opts: { dateFormat: 'D MMMM YYYY' } },
            maj: { type: 'DateTime:Europe/Paris', opts: { dateFormat: 'DD/MM/YYYY', timeFormat: 'HH:mm' } },
        },
    };
    const buffer = template(app, para('Né le {Naissance}') + para('{#Visites}Visite du {Visites.Jour} {/Visites}'), '', para('Mise à jour : {MAJ}'));
    const data = { Naissance: new Date('1990-03-14T00:00:00Z'), MAJ: new Date('2026-11-09T09:30:00Z'), Visites: [{ Jour: 1790899200 }] };
    const blob = app.generateDocxBlob(data, buffer, new Set(), info);
    const zip = new app.PizZip(await blob.arrayBuffer());
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Né le 14 mars 1990Visite du 02/10/2026 ');
    assert.equal(xmlText(zip.file('word/footer1.xml').asText()), 'Mise à jour : 09/11/2026 10:30');
});
