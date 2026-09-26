/**
 * MAISON RETROUVAILLES – bestillingssystem (Google Apps Script)
 * ---------------------------------------------------------------
 * Denne ene filen er hele «serveren»: den tar imot bestillinger, lagrer dem
 * i regnearket, viser køen til kjøkkenet og sender push-varsel til gjestene.
 *
 * Du trenger bare å endre de to linjene under INNSTILLINGER.
 */

// ============================ INNSTILLINGER ============================
var KJOKKEN_PIN = String(typeof HEMMELIG_PIN !== 'undefined' ? HEMMELIG_PIN : 'ikke-satt-' + Date.now()); // settes i oppstartsskriptet i Apps Script                    // PIN-kode for kjøkkensiden (bytt gjerne)
var KONTAKT     = 'https://freddzy-jald.github.io/bar/'; // Kontaktadresse Apple krever for push (må være en ekte adresse). Kan stå som den er.
// ======================================================================

var VAPID_PUBLIC  = 'BLwD9USQl8x5TDdiGerqHhAxmkcOF2fWi49p4Tt5bMKp5AiVW0GKe6q489xFW3eE4gNUy6fs8y-oRVy8ePNf8nA';
var VAPID_PRIVATE = String(typeof HEMMELIG_VAPID !== 'undefined' ? HEMMELIG_VAPID : ''); // settes i oppstartsskriptet i Apps Script
var TIDSSONE = 'Europe/Oslo';
var ARK_ID = '122vw5iLLoBgGXOTIV8GpXqpuoFy5ksa8Byo28CS4y94'; // brukes bare når skriptet ikke ligger inne i et regneark

function bok() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
  if (!ss && ARK_ID.indexOf('__') !== 0) ss = SpreadsheetApp.openById(ARK_ID);
  if (!ss) throw new Error('Fant ikke regnearket');
  return ss;
}
var ARK_BESTILLINGER = 'Bestillinger', ARK_GJESTER = 'Gjester', ARK_VARSLER = 'Varsler';
var KOL_B = ['id', 'nr', 'tid', 'navn', 'drink', 'drinknavn', 'antall', 'notat', 'status', 'endret', 'varslet', 'token', 'klokkeslett', 'egen drink'];
var MAKS_APNE_PER_GJEST = 4;

/* ---------------------------------------------------------------------
   INNGANGER
   --------------------------------------------------------------------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.a) return jsonSvar(kjorApi(p.a, parseData(p.d)));
  return visSide(p.side === 'kjokken' ? 'kjokken' : 'gjest');
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  return jsonSvar(kjorApi(body.a, body.d || {}));
}

/* Brukes av reservesidene (HtmlService) via google.script.run – trenger ingen nettadresse */
function apiKall(a, d) { return kjorApi(a, d || {}); }

function parseData(s) {
  if (!s) return {};
  try { return JSON.parse(s); } catch (err) { return {}; }
}

function jsonSvar(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Reserve: sidene kan også vises direkte fra Apps Script (uten push). */
function visSide(hvilken) {
  var url = ScriptApp.getService().getUrl();
  var html = (hvilken === 'kjokken' ? HTML_KJOKKEN : HTML_GJEST).split('__API_URL__').join(url);
  return HtmlService.createHtmlOutput(html)
    .setTitle(hvilken === 'kjokken' ? 'Kjøkkenet – ' + BAR.navn : BAR.navn)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------------------------------------------------------------------
   API
   --------------------------------------------------------------------- */
var OFFENTLIG = { ping: 1, meny: 1, registrer: 1, meg: 1, bestill: 1, mine: 1, avbestill: 1, abonner: 1, varsel: 1, testmeg: 1 };
var KJOKKEN = { ko: 1, status: 1, utsolgt: 1, tomt: 1, innstillinger: 1, testvarsel: 1, sjekk: 1 };

function kjorApi(a, d) {
  d = d || {};
  try {
    if (OFFENTLIG[a]) return API[a](d);
    if (KJOKKEN[a]) {
      if (String(d.pin || '') !== String(KJOKKEN_PIN)) return { ok: false, feil: 'feil_pin', melding: 'Feil PIN-kode.' };
      return API[a](d);
    }
    return { ok: false, feil: 'ukjent', melding: 'Ukjent handling.' };
  } catch (err) {
    console.error(a, err && err.stack || err);
    return { ok: false, feil: 'server', melding: String(err && err.message || err) };
  }
}

var API = {};

API.ping = function () { return { ok: true, tid: Date.now() }; };

API.meny = function () {
  var inn = hentInnstillinger();
  return { ok: true, apen: inn.apen, melding: inn.melding, utsolgt: inn.utsolgt, tomt: inn.tomt, vapid: VAPID_PUBLIC, tid: Date.now() };
};

/* Gjesten lagrer navn (og valgfritt telefonnummer) på sin token */
API.registrer = function (d) {
  var token = rensToken(d.token), navn = rensTekst(d.navn, 30);
  if (!token) return { ok: false, feil: 'token', melding: 'Mangler enhets-ID.' };
  if (!navn) return { ok: false, feil: 'navn', melding: 'Skriv inn navnet ditt.' };
  var tlf = rensTelefon(d.tlf);
  medLas(function () {
    var ark = hentArk(ARK_GJESTER), rader = ark.getDataRange().getValues(), funnet = false;
    for (var i = 1; i < rader.length; i++) {
      if (rader[i][0] === token) {
        ark.getRange(i + 1, 2, 1, 3).setValues([[navn, d.tlf === undefined ? rader[i][2] : tlf, Date.now()]]);
        funnet = true; break;
      }
    }
    if (!funnet) ark.appendRow([token, navn, tlf, Date.now()]);
    CacheService.getScriptCache().remove('gjester');
  });
  return { ok: true };
};

API.meg = function (d) {
  var token = rensToken(d.token);
  var g = hentGjester()[token] || null;
  return { ok: true, navn: g ? g.navn : '', tlf: g ? g.tlf : '', harVarsel: !!hentVarselTokens()[token] };
};

API.bestill = function (d) {
  var inn = hentInnstillinger();
  if (!inn.apen) return { ok: false, feil: 'stengt', melding: inn.melding || 'Baren tar ikke imot bestillinger akkurat nå.' };
  var token = rensToken(d.token), navn = rensTekst(d.navn, 30), erEgen = String(d.drink || '') === 'egen';
  var drink = erEgen ? { id: 'egen', navn: 'Votre Création', ekte: 'Egen drink' } : BAR.finnDrink(String(d.drink || ''));
  var antall = Math.max(1, Math.min(3, parseInt(d.antall, 10) || 1));
  var notat = rensTekst(d.notat, 100), id = String(d.id || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 40);
  if (!token || !id) return { ok: false, feil: 'token', melding: 'Noe gikk galt – last siden på nytt.' };
  if (!navn) return { ok: false, feil: 'navn', melding: 'Skriv inn navnet ditt først.' };
  if (!drink) return { ok: false, feil: 'drink', melding: 'Ukjent drink.' };
  var egen = null;
  if (erEgen) {
    var e = d.egen || {};
    egen = { ing: (Array.isArray(e.ing) ? e.ing : []).map(String).slice(0, 20), stil: String(e.stil || ''), navn: rensTekst(e.navn, 30) };
    var feilEgen = BAR.egenSjekk(egen, inn.tomt);
    if (feilEgen) return { ok: false, feil: 'egen', melding: feilEgen };
    egen.ing = BAR.egenSortert(egen.ing);
    antall = Math.min(antall, 2);
  } else if (inn.utsolgt.indexOf(drink.id) >= 0) return { ok: false, feil: 'utsolgt', melding: drink.navn + ' er dessverre utsolgt for kvelden.' };

  var resultat = null;
  medLas(function () {
    var ordre = lesOrdreFraArk();
    for (var i = 0; i < ordre.length; i++) if (ordre[i].id === id) { resultat = { ok: true, ordre: offentligOrdre(ordre[i], ordre), duplikat: true }; return; }
    var apne = ordre.filter(function (o) { return o.token === token && (o.status === 'ny' || o.status === 'lages'); }).length;
    if (apne >= MAKS_APNE_PER_GJEST) { resultat = { ok: false, feil: 'for_mange', melding: 'Du har allerede ' + apne + ' drinker på vei. Vent til en er klar før du bestiller mer.' }; return; }
    var nr = ordre.length ? Math.max.apply(null, ordre.map(function (o) { return o.nr || 0; })) + 1 : 1;
    var na = Date.now();
    var o = { id: id, nr: nr, tid: na, navn: navn, drink: drink.id, drinknavn: erEgen ? 'Egen: ' + (egen.navn || BAR.egenTekst(egen.ing)) : drink.ekte, antall: antall, notat: notat, status: 'ny', endret: na, varslet: '', token: token, egen: egen };
    hentArk(ARK_BESTILLINGER).appendRow(ordreTilRad(o));
    ordre.push(o);
    lagreOrdreICache(ordre);
    resultat = { ok: true, ordre: offentligOrdre(o, ordre) };
  });
  return resultat;
};

API.mine = function (d) {
  var token = rensToken(d.token);
  var alle = hentOrdre();
  var mine = alle.filter(function (o) { return o.token === token && o.status !== 'kansellert'; })
    .sort(function (a, b) { return b.tid - a.tid; }).slice(0, 20)
    .map(function (o) { return offentligOrdre(o, alle); });
  var inn = hentInnstillinger();
  return { ok: true, ordre: mine, apen: inn.apen, melding: inn.melding, utsolgt: inn.utsolgt, tomt: inn.tomt, tid: Date.now() };
};

API.avbestill = function (d) {
  var token = rensToken(d.token), id = String(d.id || '');
  var svar = { ok: false, feil: 'finnes_ikke', melding: 'Fant ikke bestillingen.' };
  medLas(function () {
    var ordre = lesOrdreFraArk();
    for (var i = 0; i < ordre.length; i++) {
      var o = ordre[i];
      if (o.id === id && o.token === token) {
        if (o.status !== 'ny') { svar = { ok: false, feil: 'startet', melding: 'Kjøkkenet har allerede begynt på denne.' }; return; }
        o.status = 'kansellert'; o.endret = Date.now();
        skrivOrdreRad(i, o); lagreOrdreICache(ordre);
        svar = { ok: true }; return;
      }
    }
  });
  return svar;
};

API.abonner = function (d) {
  var token = rensToken(d.token), sub = d.sub;
  if (!token || !sub || !sub.endpoint || !/^https:\/\//.test(sub.endpoint) || !sub.keys) return { ok: false, feil: 'sub', melding: 'Ugyldig varselabonnement.' };
  var json = JSON.stringify({ endpoint: sub.endpoint, keys: sub.keys });
  medLas(function () {
    var ark = hentArk(ARK_VARSLER), rader = ark.getDataRange().getValues();
    for (var i = 1; i < rader.length; i++) if (rader[i][0] === token && rader[i][1] === sub.endpoint && rader[i][2] === json) return; // uendret
    for (var j = rader.length - 1; j >= 1; j--) if (rader[j][1] === sub.endpoint) ark.deleteRow(j + 1);
    ark.appendRow([token, sub.endpoint, json, Date.now()]);
    CacheService.getScriptCache().remove('vtokens');
  });
  return { ok: true };
};

/* Service workeren spør hva varselet skal si */
API.varsel = function (d) {
  var token = rensToken(d.token);
  var c = CacheService.getScriptCache().get('sistevarsel_' + token);
  if (c) return JSON.parse(c);
  return { ok: true, tittel: 'Votre verre est prêt', tekst: 'Drinken din venter på deg ved baren.' };
};

API.testmeg = function (d) {
  var token = rensToken(d.token);
  var cache = CacheService.getScriptCache();
  if (cache.get('testmeg_' + token)) return { ok: false, feil: 'vent', melding: 'Vent litt før du tester igjen.' };
  cache.put('testmeg_' + token, '1', 15);
  var r = sendVarselTilToken(token, 'Bienvenue ✨', 'Varslene virker! Slik får du beskjed når drinken er klar.');
  return { ok: r.sendt > 0, resultat: r, melding: r.sendt > 0 ? 'Testvarsel sendt.' : 'Fant ikke noe varsel på denne telefonen.' };
};

/* ------------------------- Kjøkkenet ------------------------- */
API.ko = function () {
  var alle = hentOrdre(), gjester = hentGjester(), vt = hentVarselTokens();
  var aktive = [], ferdige = [], antallPerDrink = {}, servert = 0;
  alle.forEach(function (o) {
    if (o.status === 'kansellert') return;
    var g = gjester[o.token] || {};
    var ut = { id: o.id, nr: o.nr, tid: o.tid, navn: o.navn, drink: o.drink, antall: o.antall, notat: o.notat, status: o.status, endret: o.endret, varslet: o.varslet, tlf: g.tlf || '', harVarsel: !!vt[o.token], egen: o.egen || null };
    if (o.status === 'ny' || o.status === 'lages') aktive.push(ut); else ferdige.push(ut);
    antallPerDrink[o.drink] = (antallPerDrink[o.drink] || 0) + o.antall;
    if (o.status === 'ferdig') servert += o.antall;
  });
  aktive.sort(function (a, b) { return (a.tid - b.tid) || (a.nr - b.nr); });
  ferdige.sort(function (a, b) { return b.endret - a.endret; });
  var inn = hentInnstillinger();
  var abonnenter = Object.keys(vt).map(function (t) { return { token: t, navn: (gjester[t] || {}).navn || 'Ukjent' }; });
  return { ok: true, aktive: aktive, ferdige: ferdige.slice(0, 40), antallPerDrink: antallPerDrink, servert: servert, apen: inn.apen, melding: inn.melding, utsolgt: inn.utsolgt, tomt: inn.tomt, abonnenter: abonnenter, tid: Date.now() };
};

API.status = function (d) {
  var id = String(d.id || ''), ny = String(d.status || '');
  if (['ny', 'lages', 'ferdig', 'kansellert'].indexOf(ny) < 0) return { ok: false, feil: 'status', melding: 'Ugyldig status.' };
  var funnet = null, uendret = false, avbestilt = false;
  medLas(function () {
    var ordre = lesOrdreFraArk();
    for (var i = 0; i < ordre.length; i++) {
      if (ordre[i].id === id) {
        funnet = ordre[i];
        if (ordre[i].status === ny) { uendret = true; return; }          // dobbelttrykk / nytt forsøk: ingen ny push
        if (ordre[i].status === 'kansellert') { avbestilt = true; return; }
        ordre[i].status = ny; ordre[i].endret = Date.now();
        if (ny !== 'ferdig') ordre[i].varslet = '';
        skrivOrdreRad(i, ordre[i]); lagreOrdreICache(ordre);
        return;
      }
    }
  });
  if (!funnet) return { ok: false, feil: 'finnes_ikke', melding: 'Fant ikke bestillingen.' };
  if (avbestilt) return { ok: false, feil: 'avbestilt', melding: 'Gjesten har avbestilt denne.' };
  if (uendret) return { ok: true, uendret: true, varsel: ny === 'ferdig' ? { sendt: funnet.varslet === 'push' ? 1 : 0, feil: 0, fjernet: 0, detaljer: [] } : null };
  var varsel = null;
  if (ny === 'ferdig') {
    var tittel = 'Votre verre est prêt ✨';
    var tekst;
    if (funnet.drink === 'egen') {
      var en = funnet.egen && funnet.egen.navn ? '«' + funnet.egen.navn + '»' : 'Din egen drink';
      tekst = (funnet.antall > 1 ? funnet.antall + ' × ' : '') + en + ' (Votre Création) venter på deg ved baren, ' + funnet.navn + '.';
    } else {
      var drink = BAR.finnDrink(funnet.drink) || { navn: funnet.drinknavn, ekte: funnet.drinknavn };
      tekst = (funnet.antall > 1 ? funnet.antall + ' × ' : '') + drink.navn + ' (' + drink.ekte + ') venter på deg ved baren, ' + funnet.navn + '.';
    }
    varsel = sendVarselTilToken(funnet.token, tittel, tekst);
    var merke = varsel.sendt > 0 ? 'push' : (varsel.feil > 0 ? 'feilet' : 'ingen');
    medLas(function () {
      var ordre = lesOrdreFraArk();
      for (var i = 0; i < ordre.length; i++) if (ordre[i].id === id) { ordre[i].varslet = merke; skrivOrdreRad(i, ordre[i]); lagreOrdreICache(ordre); return; }
    });
  }
  return { ok: true, varsel: varsel };
};

API.utsolgt = function (d) {
  var inn = hentInnstillinger(), id = String(d.drink || '');
  if (!BAR.finnDrink(id)) return { ok: false, feil: 'drink' };
  var liste = inn.utsolgt.filter(function (x) { return x !== id; });
  if (d.utsolgt) liste.push(id);
  PropertiesService.getScriptProperties().setProperty('utsolgt', JSON.stringify(liste));
  CacheService.getScriptCache().remove('innst');
  return { ok: true, utsolgt: liste };
};

/* Ingredienser som er tomme (skjules i «Lag din egen») */
API.tomt = function (d) {
  var inn = hentInnstillinger(), id = String(d.ing || '');
  if (!BAR.egenKategori(id)) return { ok: false, feil: 'ing' };
  var liste = inn.tomt.filter(function (x) { return x !== id; });
  if (d.tomt) liste.push(id);
  PropertiesService.getScriptProperties().setProperty('tomt', JSON.stringify(liste));
  CacheService.getScriptCache().remove('innst');
  return { ok: true, tomt: liste };
};

API.innstillinger = function (d) {
  var props = PropertiesService.getScriptProperties();
  if (d.apen !== undefined) props.setProperty('apen', d.apen ? '1' : '0');
  if (d.melding !== undefined) props.setProperty('melding', rensTekst(d.melding, 140));
  CacheService.getScriptCache().remove('innst');
  return { ok: true, innstillinger: hentInnstillinger() };
};

API.testvarsel = function (d) {
  var token = rensToken(d.token);
  var r = sendVarselTilToken(token, 'Test fra kjøkkenet', 'Hvis du ser dette, virker varslene. Santé!');
  return { ok: r.sendt > 0, resultat: r };
};

API.sjekk = function () { return { ok: true, rapport: systemsjekk() }; };

/* ---------------------------------------------------------------------
   LAGRING (regneark + hurtigbuffer)
   --------------------------------------------------------------------- */
function hentArk(navn) {
  var ss = bok();
  var ark = ss.getSheetByName(navn);
  if (!ark) {
    ark = ss.insertSheet(navn);
    var hode = navn === ARK_BESTILLINGER ? KOL_B : navn === ARK_GJESTER ? ['token', 'navn', 'telefon', 'oppdatert'] : ['token', 'endpoint', 'abonnement', 'tid'];
    ark.getRange(1, 1, 1, hode.length).setValues([hode]).setFontWeight('bold');
    ark.setFrozenRows(1);
  }
  return ark;
}

function ordreTilRad(o) {
  var klokke = Utilities.formatDate(new Date(o.tid), TIDSSONE, 'HH:mm');
  return [o.id, o.nr, o.tid, o.navn, o.drink, o.drinknavn, o.antall, o.notat, o.status, o.endret, o.varslet || '', o.token, klokke, o.egen ? JSON.stringify(o.egen) : ''];
}

function radTilOrdre(r) {
  var egen = null; if (r[13]) { try { egen = JSON.parse(r[13]); } catch (err) { } }
  return { id: String(r[0]), nr: Number(r[1]) || 0, tid: Number(r[2]) || 0, navn: String(r[3]), drink: String(r[4]), drinknavn: String(r[5]), antall: Number(r[6]) || 1, notat: String(r[7] || ''), status: String(r[8]), endret: Number(r[9]) || 0, varslet: String(r[10] || ''), token: String(r[11]), egen: egen };
}

function lesOrdreFraArk() {
  var ark = hentArk(ARK_BESTILLINGER);
  var n = ark.getLastRow();
  if (n < 2) return [];
  return ark.getRange(2, 1, n - 1, KOL_B.length).getValues().filter(function (r) { return r[0] !== ''; }).map(radTilOrdre);
}

function skrivOrdreRad(indeks, o) {
  hentArk(ARK_BESTILLINGER).getRange(indeks + 2, 1, 1, KOL_B.length).setValues([ordreTilRad(o)]);
}

function lagreOrdreICache(ordre) {
  try { CacheService.getScriptCache().put('ordre', JSON.stringify(ordre), 1800); } catch (err) { CacheService.getScriptCache().remove('ordre'); }
}

function hentOrdre() {
  var c = CacheService.getScriptCache().get('ordre');
  if (c) { try { return JSON.parse(c); } catch (err) { } }
  // Hurtigbufferen er tom: les regnearket. Fyll bufferen bare hvis vi får låsen (så vi aldri overskriver nyere data).
  var las = LockService.getScriptLock();
  if (las.tryLock(3000)) {
    try { var ordre = lesOrdreFraArk(); lagreOrdreICache(ordre); return ordre; } finally { las.releaseLock(); }
  }
  return lesOrdreFraArk();
}

function hentGjester() {
  var cache = CacheService.getScriptCache(), c = cache.get('gjester');
  if (c) return JSON.parse(c);
  var rader = hentArk(ARK_GJESTER).getDataRange().getValues(), m = {};
  for (var i = 1; i < rader.length; i++) if (rader[i][0]) m[rader[i][0]] = { navn: String(rader[i][1]), tlf: String(rader[i][2] || '') };
  try { cache.put('gjester', JSON.stringify(m), 300); } catch (err) { }
  return m;
}

function hentVarselTokens() {
  var cache = CacheService.getScriptCache(), c = cache.get('vtokens');
  if (c) return JSON.parse(c);
  var rader = hentArk(ARK_VARSLER).getDataRange().getValues(), m = {};
  for (var i = 1; i < rader.length; i++) if (rader[i][0]) m[rader[i][0]] = 1;
  try { cache.put('vtokens', JSON.stringify(m), 300); } catch (err) { }
  return m;
}

function hentInnstillinger() {
  var cache = CacheService.getScriptCache(), c = cache.get('innst');
  if (c) return JSON.parse(c);
  var p = PropertiesService.getScriptProperties().getProperties();
  var inn = { apen: p.apen !== '0', melding: p.melding || '', utsolgt: [], tomt: [] };
  try { inn.utsolgt = JSON.parse(p.utsolgt || '[]'); } catch (err) { }
  try { inn.tomt = JSON.parse(p.tomt || '[]'); } catch (err) { }
  cache.put('innst', JSON.stringify(inn), 60);
  return inn;
}

function medLas(fn) {
  var las = LockService.getScriptLock();
  las.waitLock(25000);
  try { fn(); SpreadsheetApp.flush(); } finally { las.releaseLock(); }
}

/* Det gjesten får se om en bestilling (inkl. plass i køen) */
function offentligOrdre(o, alle) {
  var foran = 0;
  if (o.status === 'ny' || o.status === 'lages') {
    alle.forEach(function (x) { if ((x.status === 'ny' || x.status === 'lages') && (x.tid < o.tid || (x.tid === o.tid && x.nr < o.nr))) foran++; });
  }
  return { id: o.id, nr: o.nr, tid: o.tid, drink: o.drink, antall: o.antall, notat: o.notat, status: o.status, endret: o.endret, foran: foran, egen: o.egen || null };
}

function rensTekst(s, maks) { return String(s === undefined || s === null ? '' : s).replace(/[\u0000-\u001f<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, maks); }
function rensToken(s) { return String(s || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 40); }
function rensTelefon(s) { var t = String(s || '').replace(/[^0-9+]/g, ''); return t.length >= 8 ? t.slice(0, 16) : ''; }

/* ---------------------------------------------------------------------
   PUSH-VARSLER (Web Push med VAPID, uten innhold – service workeren henter teksten)
   --------------------------------------------------------------------- */
function sendVarselTilToken(token, tittel, tekst) {
  var res = { sendt: 0, feil: 0, fjernet: 0, detaljer: [] };
  if (!token) return res;
  CacheService.getScriptCache().put('sistevarsel_' + token, JSON.stringify({ ok: true, tittel: tittel, tekst: tekst }), 600);
  var ark = hentArk(ARK_VARSLER), rader = ark.getDataRange().getValues();
  var slett = [];
  for (var i = 1; i < rader.length; i++) {
    if (rader[i][0] !== token) continue;
    var sub; try { sub = JSON.parse(rader[i][2]); } catch (err) { continue; }
    var kode = 0, svartekst = '';
    try {
      var r = sendPush(sub.endpoint);
      kode = r.getResponseCode(); svartekst = String(r.getContentText() || '').slice(0, 200);
    } catch (err) { kode = -1; svartekst = String(err); }
    if (kode >= 200 && kode < 300) res.sendt++;
    else {
      res.feil++;
      if (kode === 404 || kode === 410) slett.push(sub.endpoint);
    }
    res.detaljer.push({ kode: kode, svar: svartekst, tjeneste: sub.endpoint.split('/')[2] });
  }
  if (slett.length) {
    medLas(function () {
      var r2 = ark.getDataRange().getValues();
      for (var j = r2.length - 1; j >= 1; j--) if (slett.indexOf(r2[j][1]) >= 0) ark.deleteRow(j + 1);
    });
    res.fjernet = slett.length;
    CacheService.getScriptCache().remove('vtokens');
  }
  if (res.feil) console.warn('Push-feil', JSON.stringify(res.detaljer));
  return res;
}

function sendPush(endpoint) {
  var aud = endpoint.split('/').slice(0, 3).join('/');
  var jwt = vapidJwt(aud);
  return UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/octet-stream',
    payload: '',
    headers: { 'TTL': '600', 'Urgency': 'high', 'Authorization': 'vapid t=' + jwt + ', k=' + VAPID_PUBLIC },
    muteHttpExceptions: true
  });
}

function vapidJwt(aud) {
  var cache = CacheService.getScriptCache(), key = 'jwt_' + Utilities.base64EncodeWebSafe(aud).replace(/=+$/, '');
  var c = cache.get(key);
  if (c) return c;
  var jwt = lagJwt(aud, Math.floor(Date.now() / 1000) + 12 * 3600);
  cache.put(key, jwt, 11 * 3600);
  return jwt;
}

function lagJwt(aud, exp) {
  var b64 = function (s) { return Utilities.base64EncodeWebSafe(s, Utilities.Charset.UTF_8).replace(/=+$/, ''); };
  var innhold = b64(JSON.stringify({ typ: 'JWT', alg: 'ES256' })) + '.' + b64(JSON.stringify({ aud: aud, exp: exp, sub: KONTAKT }));
  var sig = EC.sign(innhold, EC.privatNokkel());
  return innhold + '.' + Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

/* ECDSA P-256 (ES256) i ren JavaScript med BigInt. Bruker deterministisk k (RFC 6979). */
var EC = (function () {
  var B = function (x) { return BigInt(x); };
  var ZERO, ONE, TWO, THREE, FOUR, EIGHT, P, N, GX, GY;
  function init() {
    if (P) return;
    ZERO = B(0); ONE = B(1); TWO = B(2); THREE = B(3); FOUR = B(4); EIGHT = B(8);
    P = B('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
    N = B('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
    GX = B('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296');
    GY = B('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5');
  }
  function mod(a, m) { var r = a % m; return r < ZERO ? r + m : r; }
  function powmod(b, e, m) { var r = ONE; b = mod(b, m); while (e > ZERO) { if (e & ONE) r = (r * b) % m; b = (b * b) % m; e >>= ONE; } return r; }
  function inv(a, m) { return powmod(a, m - TWO, m); }
  // Jacobiske koordinater [X, Y, Z]
  function dbl(p) {
    if (p[2] === ZERO || p[1] === ZERO) return [ZERO, ONE, ZERO];
    var X = p[0], Y = p[1], Z = p[2];
    var YY = (Y * Y) % P, S = (FOUR * X * YY) % P, ZZ = (Z * Z) % P;
    var M = (THREE * (X - ZZ) * (X + ZZ)) % P;
    var X3 = mod(M * M - TWO * S, P);
    var Y3 = mod(M * (S - X3) - EIGHT * YY * YY, P);
    var Z3 = mod(TWO * Y * Z, P);
    return [X3, Y3, Z3];
  }
  function add(p, q) {
    if (p[2] === ZERO) return q;
    if (q[2] === ZERO) return p;
    var Z1Z1 = (p[2] * p[2]) % P, Z2Z2 = (q[2] * q[2]) % P;
    var U1 = (p[0] * Z2Z2) % P, U2 = (q[0] * Z1Z1) % P;
    var S1 = (p[1] * q[2] * Z2Z2) % P, S2 = (q[1] * p[2] * Z1Z1) % P;
    if (U1 === U2) { if (S1 !== S2) return [ZERO, ONE, ZERO]; return dbl(p); }
    var H = mod(U2 - U1, P), R = mod(S2 - S1, P), HH = (H * H) % P, HHH = (H * HH) % P, V = (U1 * HH) % P;
    var X3 = mod(R * R - HHH - TWO * V, P);
    var Y3 = mod(R * (V - X3) - S1 * HHH, P);
    var Z3 = (p[2] * q[2] * H) % P;
    return [X3, Y3, Z3];
  }
  function mul(k, pt) {
    var R = [ZERO, ONE, ZERO], Q = [pt[0], pt[1], ONE];
    while (k > ZERO) { if (k & ONE) R = add(R, Q); Q = dbl(Q); k >>= ONE; }
    return R;
  }
  function affin(p) { var zi = inv(p[2], P), zi2 = (zi * zi) % P; return [(p[0] * zi2) % P, (p[1] * zi2 * zi) % P]; }
  function bytesTilInt(b) { var h = '0x'; for (var i = 0; i < b.length; i++) h += ((b[i] & 255) + 256).toString(16).slice(1); return b.length ? B(h) : ZERO; }
  function intTilBytes(x) { var h = x.toString(16); while (h.length < 64) h = '0' + h; var out = []; for (var i = 0; i < 64; i += 2) out.push(parseInt(h.substr(i, 2), 16)); return out; }
  function signert(arr) { return arr.map(function (v) { v = v & 255; return v > 127 ? v - 256 : v; }); }
  function hmac(k, data) { return Utilities.computeHmacSha256Signature(signert(data), signert(k)).map(function (v) { return v & 255; }); }
  function sha256(str) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.US_ASCII).map(function (v) { return v & 255; }); }
  function fyll(n, v) { var a = []; for (var i = 0; i < n; i++) a.push(v); return a; }
  function rfc6979(x, h) {
    var xb = intTilBytes(x), z = bytesTilInt(h); if (z >= N) z -= N; var hb = intTilBytes(z);
    var V = fyll(32, 1), K = fyll(32, 0);
    K = hmac(K, V.concat([0], xb, hb)); V = hmac(K, V);
    K = hmac(K, V.concat([1], xb, hb)); V = hmac(K, V);
    for (var t = 0; t < 100; t++) {
      V = hmac(K, V);
      var k = bytesTilInt(V);
      if (k >= ONE && k < N) return k;
      K = hmac(K, V.concat([0])); V = hmac(K, V);
    }
    throw new Error('rfc6979');
  }
  function sign(melding, d) {
    init();
    var h = sha256(melding), e = bytesTilInt(h), k = rfc6979(d, h);
    var R = affin(mul(k, [GX, GY])), r = R[0] % N;
    var s = (inv(k, N) * ((e + r * d) % N)) % N;
    if (r === ZERO || s === ZERO) throw new Error('ugyldig signatur');
    return signert(intTilBytes(r).concat(intTilBytes(s)));
  }
  function verifiser(melding, sig, pub) {
    init();
    var b = sig.map(function (v) { return v & 255; });
    var r = bytesTilInt(b.slice(0, 32)), s = bytesTilInt(b.slice(32, 64));
    if (r <= ZERO || r >= N || s <= ZERO || s >= N) return false;
    var e = bytesTilInt(sha256(melding)), w = inv(s, N);
    var u1 = (e * w) % N, u2 = (r * w) % N;
    var X = add(mul(u1, [GX, GY]), mul(u2, pub));
    if (X[2] === ZERO) return false;
    return affin(X)[0] % N === r;
  }
  function privatNokkel() { init(); return bytesTilInt(Utilities.base64DecodeWebSafe(pad(VAPID_PRIVATE))); }
  function offentligPunkt() { var b = Utilities.base64DecodeWebSafe(pad(VAPID_PUBLIC)).map(function (v) { return v & 255; }); return [bytesTilInt(b.slice(1, 33)), bytesTilInt(b.slice(33, 65))]; }
  function avledOffentlig() { init(); var p = affin(mul(privatNokkel(), [GX, GY])); return [p[0], p[1]]; }
  function pad(s) { while (s.length % 4) s += '='; return s; }
  return { sign: sign, verifiser: verifiser, privatNokkel: privatNokkel, offentligPunkt: offentligPunkt, avledOffentlig: avledOffentlig };
})();

/* ---------------------------------------------------------------------
   OPPSETT OG SJEKK – kjør «oppsett» én gang fra redigeringsvinduet
   --------------------------------------------------------------------- */
function oppsett() {
  var ab = hentArk(ARK_BESTILLINGER); hentArk(ARK_GJESTER); hentArk(ARK_VARSLER);
  ab.getRange(1, 1, 1, KOL_B.length).setValues([KOL_B]).setFontWeight('bold'); // oppdaterer overskriftene hvis arket er fra en eldre versjon
  var ss = bok();
  var ark1 = ss.getSheetByName('Ark1') || ss.getSheetByName('Sheet1');
  if (ark1 && ark1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(ark1);
  var r = systemsjekk();
  console.log(r.map(function (x) { return (x.ok ? '✔ ' : '✘ ') + x.navn + (x.info ? ' – ' + x.info : ''); }).join('\n'));
  return r;
}

function systemsjekk() {
  var ut = [];
  var sjekk = function (navn, fn) { try { var info = fn(); ut.push({ navn: navn, ok: true, info: info || '' }); } catch (err) { ut.push({ navn: navn, ok: false, info: String(err && err.message || err) }); } };
  sjekk('Regneark', function () { hentArk(ARK_BESTILLINGER); return lesOrdreFraArk().length + ' bestillinger lagret'; });
  sjekk('Hurtigbuffer', function () { CacheService.getScriptCache().put('test', '1', 10); if (CacheService.getScriptCache().get('test') !== '1') throw new Error('virker ikke'); });
  sjekk('BigInt (kreves for push)', function () { if (typeof BigInt !== 'function') throw new Error('Mangler – slå på «Chrome V8-kjøretid» under Prosjektinnstillinger'); return 'OK'; });
  sjekk('Push-nøkkel', function () {
    var a = EC.avledOffentlig(), b = EC.offentligPunkt();
    if (a[0] !== b[0] || a[1] !== b[1]) throw new Error('Nøkkelparet stemmer ikke');
    return 'OK';
  });
  sjekk('Push-signatur', function () {
    var t0 = Date.now(), jwt = lagJwt('https://web.push.apple.com', Math.floor(Date.now() / 1000) + 3600);
    var deler = jwt.split('.');
    var sig = Utilities.base64DecodeWebSafe(deler[2] + '=='.slice(0, (4 - deler[2].length % 4) % 4));
    if (!EC.verifiser(deler[0] + '.' + deler[1], sig, EC.offentligPunkt())) throw new Error('Signaturen verifiserer ikke');
    return 'OK (' + (Date.now() - t0) + ' ms)';
  });
  sjekk('Nettadresse', function () { return ScriptApp.getService().getUrl() || 'Ikke distribuert ennå – gjør «Distribuer» etterpå'; });
  return ut;
}

/* ---------------------------------------------------------------------
   Nullstill kvelden (kjør manuelt fra redigeringsvinduet hvis dere vil tømme testbestillinger)
   --------------------------------------------------------------------- */
function nullstillBestillinger() {
  var ark = hentArk(ARK_BESTILLINGER);
  if (ark.getLastRow() > 1) ark.deleteRows(2, ark.getLastRow() - 1);
  CacheService.getScriptCache().removeAll(['ordre', 'innst']);
  console.log('Bestillingene er tømt. Gjester og varsler er beholdt.');
}

/* ======================= DATA OG SIDER (ikke endre) ======================= */
/* Maison Retrouvailles – felles data for meny, gjesteside, kjøkken, kalkulator og backend.
   Én kilde til sannhet: endrer du noe her, bygges alt på nytt fra dette. */
var BAR = {
  navn: "Maison Retrouvailles",
  kortnavn: "Retrouvailles",
  undertittel: "Bar & Salon",
  dato: "Samedi 26 septembre 2026",
  datoNorsk: "lørdag 26. september 2026",
  grunnlagt: "Fondée le 26 septembre 2026",

  intro: "Mesdames et messieurs. Det har gått noen år. Noen av oss har fått barn, boliglån og sterke meninger om grillkull. I kveld legger vi alt det igjen i garderoben. Huset serverer tretten drikker, hver med sitt eget temperament. Skann koden, velg med hjertet og len deg tilbake – vi gir beskjed når glasset ditt er klart.",
  fotnote: "Vann er gratis, ubegrenset og på det varmeste anbefalt mellom slagene.",

  seksjoner: [
    { id: "rafraichissements", navn: "Les Rafraîchissements", undertekst: "Høye, kalde og uforskammet forfriskende" },
    { id: "classiques", navn: "Les Classiques", undertekst: "Korte, stolte og ettertenksomme" },
    { id: "gourmandises", navn: "Les Gourmandises", undertekst: "Kremet, mørkt og etter midnatt" }
  ],

  /* Ingredienskatalog. enhet = hvordan oppskriften måler. kjop = hvordan det handles inn. */
  ingredienser: {
    gin:            { navn: "Gin", kort: "gin",                      kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 70, tekst: "flaske à 70 cl" } },
    vodka:          { navn: "Vodka", kort: "vodka",                    kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 70, tekst: "flaske à 70 cl" } },
    lys_rom:        { navn: "Lys rom", kort: "lys rom",                  kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 70, tekst: "flaske à 70 cl" } },
    mork_rom:       { navn: "Mørk rom", kort: "mørk rom",                 kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 70, tekst: "flaske à 70 cl" } },
    whisky:         { navn: "Whisky (gjerne bourbon)", kort: "whisky",  kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 50, tekst: "flaske à 50 cl" } },
    tequila:        { navn: "Tequila", kort: "tequila",                  kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 50, tekst: "flaske à 50 cl" } },
    amaretto:       { navn: "Amaretto", kort: "amaretto",                 kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 50, tekst: "flaske à 50 cl" } },
    kahlua:         { navn: "Kaffelikør (Kahlúa)", kort: "kaffelikør",      kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 50, tekst: "flaske à 50 cl" } },
    triple_sec:     { navn: "Triple sec / Cointreau", kort: "triple sec",   kat: "Brennevin og likør", enhet: "cl", kjop: { type: "flaske", str: 50, tekst: "flaske à 50 cl" } },
    bitter:         { navn: "Angostura bitter", kort: "Angostura",         kat: "Brennevin og likør", enhet: "dråper", kjop: { type: "flaske", str: 250, tekst: "flaske à 20 cl" } },

    tonic:          { navn: "Tonic", kort: "tonic",                    kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 100, tekst: "flaske à 1 l" } },
    ingefaerol:     { navn: "Ingefærøl (ginger beer)", kort: "ingefærøl",  kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 50,  tekst: "flaske à 0,5 l" } },
    cola:           { navn: "Cola", kort: "cola",                     kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 150, tekst: "flaske à 1,5 l" } },
    soda:           { navn: "Soda / Farris", kort: "soda",            kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 150, tekst: "flaske à 1,5 l" } },
    grapefruitbrus: { navn: "Grapefruitbrus", kort: "grapefruitbrus",           kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 150, tekst: "flaske à 1,5 l" } },
    appelsinjuice:  { navn: "Appelsinjuice", kort: "appelsinjuice",            kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 100, tekst: "kartong à 1 l" } },
    grenadine:      { navn: "Grenadine", kort: "grenadine",                kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 70,  tekst: "flaske à 70 cl" } },
    flote:          { navn: "Kremfløte", kort: "kremfløte",                kat: "Mixere", enhet: "cl", kjop: { type: "flaske", str: 50,  tekst: "kartong à 5 dl" } },
    espresso:       { navn: "Espresso", kort: "espresso",                 kat: "Mixere", enhet: "cl", kjop: { type: "kaffe", gPer: 2.7, tekst: "kaffe, ca. 8 g per shot à 3 cl" } },

    limejuice:      { navn: "Limejuice, fersk", kort: "fersk limejuice",         kat: "Frukt og friskt", enhet: "cl", kjop: { type: "frukt", frukt: "lime", perFrukt: 3 } },
    limebat:        { navn: "Limebåt", kort: "limebåt", kortN: "limebåter",                  kat: "Frukt og friskt", enhet: "stk", kjop: { type: "frukt", frukt: "lime", perFrukt: 6 } },
    sitronsaft:     { navn: "Sitronsaft, fersk", kort: "fersk sitronsaft",        kat: "Frukt og friskt", enhet: "cl", kjop: { type: "frukt", frukt: "sitron", perFrukt: 4 } },
    sitronskall:    { navn: "Sitronskall (pynt)", kort: "sitronskall", kortN: "sitronskall",       kat: "Frukt og friskt", enhet: "stk", kjop: { type: "frukt", frukt: "sitron", perFrukt: 6 } },
    appelsinskall:  { navn: "Appelsinskall (pynt)", kort: "appelsinskall", kortN: "appelsinskall",     kat: "Frukt og friskt", enhet: "stk", kjop: { type: "frukt", frukt: "appelsin", perFrukt: 8 } },
    appelsinskive:  { navn: "Appelsinskive (pynt)", kort: "appelsinskive", kortN: "appelsinskiver",     kat: "Frukt og friskt", enhet: "stk", kjop: { type: "frukt", frukt: "appelsin", perFrukt: 8 } },
    mynte:          { navn: "Mynteblader", kort: "mynteblader",              kat: "Frukt og friskt", enhet: "blader", kjop: { type: "frukt", frukt: "mynte", perFrukt: 50 } },
    egg:            { navn: "Eggehvite", kort: "eggehvite", kortN: "eggehviter",                kat: "Frukt og friskt", enhet: "stk", kjop: { type: "frukt", frukt: "egg", perFrukt: 1 } },

    sukker:         { navn: "Sukker", kort: "sukker",                   kat: "Søtt", enhet: "ts", kjop: { type: "sukker", gPer: 4 } },
    sukkersirup:    { navn: "Sukkersirup (1:1)", kort: "sukkersirup",        kat: "Søtt", enhet: "cl", kjop: { type: "sukker", gPer: 6 } },
    honningsirup:   { navn: "Honningsirup (1:1)", kort: "honningsirup",       kat: "Søtt", enhet: "cl", kjop: { type: "honning", gPer: 7 } },
    kaffebonner:    { navn: "Hele kaffebønner (pynt)", kort: "kaffebønne", kortN: "kaffebønner",  kat: "Søtt", enhet: "stk", kjop: { type: "ingen" } }
  },

  /* Hvordan frukt og annet handles inn */
  frukt: {
    lime:     { navn: "Lime",               enhet: "stk" },
    sitron:   { navn: "Sitron",             enhet: "stk" },
    appelsin: { navn: "Appelsin",           enhet: "stk" },
    mynte:    { navn: "Mynte",              enhet: "potte/bunt" },
    egg:      { navn: "Egg",                enhet: "stk" }
  },
  /* Tips / pourboire */
  vipps: "948 62 106",

  /* «Lag din egen» – husets spiskammer i logisk rekkefølge */
  egen: {
    maksSprit: 3,
    maksTotalt: 8,
    stiler: [
      { id: "lang", navn: "Lang og leskende", besk: "Høyt glass med is, toppet med brus" },
      { id: "kort", navn: "Kort og konsentrert", besk: "Lite glass, mer smak per slurk" },
      { id: "fri", navn: "Bartenderen bestemmer", besk: "Overrask meg" }
    ],
    kategorier: [
      { id: "sprit",  navn: "Brennevin",            ing: ["gin", "vodka", "lys_rom", "mork_rom", "whisky", "tequila"] },
      { id: "likor",  navn: "Likører",              ing: ["amaretto", "kahlua", "triple_sec"] },
      { id: "sitrus", navn: "Sitrus og juice",      ing: ["limejuice", "sitronsaft", "appelsinjuice"] },
      { id: "brus",   navn: "Brus",                 ing: ["tonic", "ingefaerol", "cola", "soda", "grapefruitbrus"] },
      { id: "sott",   navn: "Søtt, krem og kaffe",  ing: ["sukkersirup", "honningsirup", "grenadine", "flote", "espresso", "egg"] },
      { id: "aroma",  navn: "Aroma og pynt",        ing: ["mynte", "bitter", "limebat", "appelsinskall", "kaffebonner"] }
    ],
    navn: {
      gin: "Gin", vodka: "Vodka", lys_rom: "Lys rom", mork_rom: "Mørk rom", whisky: "Whisky", tequila: "Tequila",
      amaretto: "Amaretto", kahlua: "Kaffelikør", triple_sec: "Triple sec (appelsinlikør)",
      limejuice: "Fersk lime", sitronsaft: "Fersk sitron", appelsinjuice: "Appelsinjuice",
      tonic: "Tonic", ingefaerol: "Ingefærøl", cola: "Cola", soda: "Soda", grapefruitbrus: "Grapefruitbrus",
      sukkersirup: "Sukkersirup", honningsirup: "Honningsirup", grenadine: "Grenadine", flote: "Kremfløte", espresso: "Espresso", egg: "Eggehvite (gir skum)",
      mynte: "Fersk mynte", bitter: "Angostura bitter", limebat: "Limebåt", appelsinskall: "Appelsinskall", kaffebonner: "Kaffebønner"
    }
  },
  isKgPerDrink: 0.3,
  isPoseKg: 2,
  sikkerhetsmargin: 0.10,

  drinker: [
    {
      id: "gt", navn: "L'Heure Bleue", ekte: "Gin & Tonic", seksjon: "rafraichissements",
      tagline: "Einer og kinin i skumringens første minutt.",
      beskrivelse: "Det finnes et øyeblikk etter solnedgang da himmelen blir blå og ingen har bestemt seg for kvelden ennå. Vi har fanget det i et glass: einerbær, bitter kinin og et presset glimt av lime. Klar som krystall, tørr som en britisk kommentar.",
      iGlasset: "Gin · tonic · lime",
      menyIng: "5 cl gin · 15 cl tonic · presset limebåt",
      glass: "Highball eller ballongglass", pynt: "Limebåt",
      oppskrift: [ { i: "gin", m: 5 }, { i: "tonic", m: 15 }, { i: "limebat", m: 1, vis: "1 limebåt, presset" } ],
      metode: [ "Fyll glasset helt med is.", "Hell i gin, topp med tonic.", "Press limebåten over og slipp den i. Rør forsiktig én gang." ]
    },
    {
      id: "mule", navn: "Le Tsar en Cavale", ekte: "Moscow Mule", seksjon: "rafraichissements",
      tagline: "En keiserlig flukt gjennom snøstorm – med ingefær i lomma.",
      beskrivelse: "Vodka like kald som steppene han forlot, ingefærøl som brenner som en hemmelighet han aldri fortalte, og lime som den siste solstrålen før grensen. Serveres til dem som ikke har tenkt å bli tatt.",
      iGlasset: "Vodka · ingefærøl · lime",
      menyIng: "5 cl vodka · 10 cl ingefærøl · 1 cl fersk lime",
      glass: "Kobberkrus eller highball", pynt: "Limebåt",
      oppskrift: [ { i: "vodka", m: 5 }, { i: "ingefaerol", m: 10 }, { i: "limejuice", m: 1 }, { i: "limebat", m: 1, pynt: true } ],
      metode: [ "Fyll glasset med is.", "Hell i vodka og limejuice.", "Topp med ingefærøl, rør én gang. Pynt med limebåt." ]
    },
    {
      id: "cuba", navn: "La Jeunesse Retrouvée", ekte: "Cuba Libre", seksjon: "rafraichissements",
      tagline: "Drinken fra da vi var sytten – nå i dress.",
      beskrivelse: "Husker du den første? Lunken, i et plastbeger, bak en gymsal. Dette er den samme historien, fortalt på nytt av noen som har lært seg manerer: lys rom, iskald cola og lime presset rett over. Vi har blitt voksne. Nesten.",
      iGlasset: "Lys rom · cola · lime",
      menyIng: "5 cl lys rom · 12 cl cola · presset limebåt",
      glass: "Highball", pynt: "Limebåt",
      oppskrift: [ { i: "lys_rom", m: 5 }, { i: "cola", m: 12 }, { i: "limebat", m: 1, vis: "1 limebåt, presset" } ],
      metode: [ "Fyll glasset med is.", "Hell i rom, topp med cola.", "Press limebåten over og slipp den i." ]
    },
    {
      id: "sunrise", navn: "Impression, Soleil Levant", ekte: "Tequila Sunrise", seksjon: "rafraichissements",
      tagline: "Monet malte en soloppgang. Vi heller en.",
      beskrivelse: "Le Havre, 1872: en maler setter penselen mot et lerret av tåke og oransje lys. Her er vår gjengivelse – tequila og appelsin, med grenadine som synker som morgenrøde mot bunnen. Se på den et øyeblikk før du drikker. I noen sekunder er den et mesterverk.",
      iGlasset: "Tequila · appelsin · grenadine",
      menyIng: "5 cl tequila · 10 cl appelsinjuice · 1,5 cl grenadine",
      glass: "Highball", pynt: "Appelsinskive",
      oppskrift: [ { i: "tequila", m: 5 }, { i: "appelsinjuice", m: 10 }, { i: "grenadine", m: 1.5 }, { i: "appelsinskive", m: 1, pynt: true } ],
      metode: [ "Fyll glasset med is, hell i tequila og appelsinjuice og rør.", "Hell grenadinen sakte langs innsiden av glasset så den synker til bunns.", "Ikke rør etterpå! Pynt med appelsinskive." ]
    },
    {
      id: "paloma", navn: "La Colombe Amoureuse", ekte: "Paloma", seksjon: "rafraichissements",
      tagline: "Lett som en due, bitter som et gammelt kjærlighetsbrev.",
      beskrivelse: "Paloma betyr due, og denne flyr lett. Tequila, perlende grapefrukt og lime – en drink som flørter litt med alle rundt bordet og aldri helt bestemmer seg. Bitter nok til å være ærlig, søt nok til å bli tilgitt.",
      iGlasset: "Tequila · grapefruitbrus · lime",
      menyIng: "5 cl tequila · 10 cl grapefruitbrus · 1 cl fersk lime",
      glass: "Highball", pynt: "Limebåt",
      oppskrift: [ { i: "tequila", m: 5 }, { i: "grapefruitbrus", m: 10 }, { i: "limejuice", m: 1 }, { i: "limebat", m: 1, pynt: true } ],
      metode: [ "Fyll glasset med is.", "Hell i tequila og limejuice.", "Topp med grapefruitbrus, rør én gang. Pynt med limebåt." ]
    },
    {
      id: "mojito", navn: "Le Jardin Secret", ekte: "Mojito", seksjon: "rafraichissements",
      tagline: "En liten ferie du kan holde i hånden.",
      beskrivelse: "Bak en port ingen andre vet om, vokser mynten vill. Vi knuser bladene varsomt, som om de betrodde oss noe, og blander dem med lys rom, lime og sukker før boblene får det siste ordet. Friskt som et åpent vindu en sommermorgen.",
      iGlasset: "Lys rom · mynte · lime · sukker · soda",
      menyIng: "5 cl lys rom · 2,5 cl fersk lime · 2 ts sukker · mynte · soda",
      glass: "Highball", pynt: "Myntekvast",
      oppskrift: [ { i: "lys_rom", m: 5 }, { i: "limejuice", m: 2.5 }, { i: "sukker", m: 2 }, { i: "mynte", m: 10, vis: "ca. 10 mynteblader" }, { i: "soda", m: 5, vis: "Topp med soda (ca. 5 cl)" } ],
      metode: [ "Legg mynte, sukker og limejuice i glasset. Knus mynten forsiktig (ikke mos den).", "Fyll glasset med is (gjerne knust) og hell i rom.", "Topp med soda og rør fra bunnen. Pynt med en myntekvast." ]
    },
    {
      id: "dark", navn: "Le Naufrage Élégant", ekte: "Dark 'n' Stormy", seksjon: "rafraichissements",
      tagline: "Et uvær i glass. Gå under med stil.",
      beskrivelse: "Mørk rom synker som et vrak gjennom skummende ingefærøl, mens limen blinker som et fyrtårn i det fjerne. Himmelen er svart, sjøen er vill, og kapteinen har bestemt seg for å nyte det. Det har vi også.",
      iGlasset: "Mørk rom · ingefærøl · lime",
      menyIng: "5 cl mørk rom · 10 cl ingefærøl · presset limebåt",
      glass: "Highball", pynt: "Limebåt",
      oppskrift: [ { i: "mork_rom", m: 5 }, { i: "ingefaerol", m: 10 }, { i: "limebat", m: 1, vis: "1 limebåt, presset" } ],
      metode: [ "Fyll glasset med is og hell i ingefærølet først.", "Hell den mørke rommen sakte over baksiden av en skje, så den legger seg øverst som en storm.", "Press limebåten over og slipp den i." ]
    },
    {
      id: "oldf", navn: "L'Héritage", ekte: "Old Fashioned", seksjon: "classiques",
      tagline: "Rørt med tålmodigheten til gamle penger.",
      beskrivelse: "Fra et skinnkledd bibliotek der klokken tikker saktere enn ellers. Whisky, et snev av sukker og tre dråper bitter, rørt til den blir silke og servert over én stor isbit. En drink man ikke drikker, men arver. Nytes sittende, gjerne med en betydningsfull pause.",
      iGlasset: "Whisky · sukker · Angostura",
      menyIng: "6 cl whisky · 1 cl sukkersirup · 3 dråper Angostura · en skvett soda",
      glass: "Lavt tumblerglass, gjerne med én stor isbit", pynt: "Appelsinskall",
      oppskrift: [ { i: "whisky", m: 6 }, { i: "sukkersirup", m: 1 }, { i: "bitter", m: 3, vis: "3 dråper Angostura" }, { i: "soda", m: 1, vis: "En skvett soda" }, { i: "appelsinskall", m: 1, pynt: true } ],
      metode: [ "Rør whisky, sukkersirup og bitter med is i 20–30 sekunder.", "Sil over i glass med en stor isbit, og tilsett en skvett soda.", "Vri appelsinskallet over glasset og legg det oppi." ]
    },
    {
      id: "amaretto", navn: "Le Velours d'Amande", ekte: "Amaretto Sour", seksjon: "classiques",
      tagline: "Mandel, sitron og et skum som burde vært forbudt.",
      beskrivelse: "Stryk fingeren over fløyel, og du vet omtrent hvordan denne kjennes. Amaretto og fersk sitron hviler under en sky av pisket eggehvite, så myk at den nesten er uanstendig. Søt, sur og helt uten skam. Drikk den langsomt, og la barten være i fred.",
      iGlasset: "Amaretto · sitron · eggehvite",
      menyIng: "6 cl amaretto · 3 cl fersk sitron · 1,5 cl sukkersirup · eggehvite",
      glass: "Tumbler med is", pynt: "Sitronskall",
      oppskrift: [ { i: "amaretto", m: 6 }, { i: "sitronsaft", m: 3 }, { i: "sukkersirup", m: 1.5 }, { i: "egg", m: 1, vis: "Eggehviten fra 1 egg" }, { i: "sitronskall", m: 1, pynt: true } ],
      metode: [ "Rist alt UTEN is i 10 sekunder (det gir skummet).", "Tilsett is og rist hardt i 10 sekunder til.", "Sil over i glass med is. Pynt med sitronskall." ]
    },
    {
      id: "bees", navn: "Les Genoux de l'Abeille", ekte: "Bee's Knees", seksjon: "classiques",
      tagline: "Forbudstid, jazz og honning.",
      beskrivelse: "Oppfunnet i forbudstiden, da ginen ble brent i badekar og honningen måtte skjule forbrytelsen. Ginen er heldigvis lovlig nå, men honningen beholdt vi av ren nostalgi. Sitron, gull og saksofon – «the bee's knees» var 1920-tallets måte å si «det aller beste» på.",
      iGlasset: "Gin · honning · sitron",
      menyIng: "5 cl gin · 2 cl fersk sitron · 2 cl honningsirup",
      glass: "Coupe eller cocktailglass, gjerne kjølt", pynt: "Sitronskall",
      oppskrift: [ { i: "gin", m: 5 }, { i: "sitronsaft", m: 2 }, { i: "honningsirup", m: 2 }, { i: "sitronskall", m: 1, pynt: true } ],
      metode: [ "Rist alt hardt med is i 12 sekunder.", "Sil over i et kjølt glass uten is.", "Pynt med sitronskall." ]
    },
    {
      id: "whip", navn: "Le Fouet Doré", ekte: "Orange Whip", seksjon: "gourmandises",
      tagline: "Iskrem for voksne. Ingen trenger den. Alle vil ha den.",
      beskrivelse: "En kremet, oransje overdådighet med tre sorter brennevin, fordi to hadde vært for beskjedent. Vodka, lys rom og appelsinlikør virvles sammen med appelsinjuice og fløte til noe som smaker som en sommerferie på syttitallet. Kjent fra en viss film der noen bestiller tre på rad. Vi anbefaler det samme.",
      iGlasset: "Vodka · lys rom · appelsinlikør · appelsin · fløte",
      menyIng: "2 cl vodka · 2 cl lys rom · 2 cl triple sec · 4 cl appelsinjuice · 2 cl fløte",
      glass: "Tumbler eller highball med is", pynt: "Appelsinskive",
      oppskrift: [ { i: "vodka", m: 2 }, { i: "lys_rom", m: 2 }, { i: "triple_sec", m: 2 }, { i: "appelsinjuice", m: 4 }, { i: "flote", m: 2 }, { i: "appelsinskive", m: 1, pynt: true } ],
      metode: [ "Rist alt med is i 10 sekunder.", "Sil over i glass med fersk is.", "Pynt med appelsinskive." ]
    },
    {
      id: "wr", navn: "Le Duc en Peignoir", ekte: "White Russian", seksjon: "gourmandises",
      tagline: "For aristokraten som har gitt opp alt unntatt kosen.",
      beskrivelse: "Hertugen har ikke kledd på seg i dag, og han har ingen planer om å gjøre det. Vodka og kaffelikør under et teppe av kald fløte – en drink for morgenkåpe, tøfler og store tanker som aldri blir til noe. Ro, rikdom og null forpliktelser.",
      iGlasset: "Vodka · kaffelikør · fløte",
      menyIng: "5 cl vodka · 2,5 cl kaffelikør · 3 cl kremfløte",
      glass: "Lavt tumblerglass", pynt: "Ingen – fløten er pynten",
      oppskrift: [ { i: "vodka", m: 5 }, { i: "kahlua", m: 2.5 }, { i: "flote", m: 3 } ],
      metode: [ "Fyll glasset med is, hell i vodka og kaffelikør.", "Hell fløten sakte over baksiden av en skje så den legger seg øverst.", "Serveres urørt – gjesten rører selv." ]
    },
    {
      id: "em", navn: "L'Insomnie Parisienne", ekte: "Espresso Martini", seksjon: "gourmandises",
      tagline: "Søvn er for dem som ikke ble invitert.",
      beskrivelse: "Klokken har passert midnatt, men samtalen nekter å dø. Fersk espresso, vodka og kaffelikør ristes hardt til en fløyelsbrun crema og krones med tre bønner: én for helse, én for rikdom, én for lykke. Paris sover aldri, og det skal ikke du heller.",
      iGlasset: "Vodka · espresso · kaffelikør",
      menyIng: "5 cl vodka · 2,5 cl kaffelikør · 3 cl espresso · 1 cl sukkersirup",
      glass: "Coupe eller martiniglass, gjerne kjølt", pynt: "3 kaffebønner",
      oppskrift: [ { i: "vodka", m: 5 }, { i: "kahlua", m: 2.5 }, { i: "espresso", m: 3 }, { i: "sukkersirup", m: 1 }, { i: "kaffebonner", m: 3, pynt: true } ],
      metode: [ "Lag en fersk espresso (3 cl).", "Rist alt veldig hardt med mye is i 15 sekunder – det gir cremaen.", "Sil over i kjølt glass. Legg tre kaffebønner på skummet." ]
    }
  ]
};

/* ---------- «Lag din egen» ---------- */
BAR.egenKategori = function (id) { var k = BAR.egen.kategorier; for (var i = 0; i < k.length; i++) if (k[i].ing.indexOf(id) >= 0) return k[i].id; return null; };
BAR.egenAlkohol = function (liste) { return liste.filter(function (i) { var k = BAR.egenKategori(i); return k === "sprit" || k === "likor"; }); };

/* Sjekker et valg fra gjesten. Returnerer feilmelding eller null. */
BAR.egenSjekk = function (e, tomt) {
  if (!e || !e.ing || !e.ing.length) return "Velg minst én ingrediens.";
  var sett = {};
  for (var i = 0; i < e.ing.length; i++) {
    var id = e.ing[i];
    if (!BAR.egenKategori(id)) return "Ukjent ingrediens.";
    if (sett[id]) return "Samme ingrediens er valgt to ganger.";
    if (tomt && tomt.indexOf(id) >= 0) return BAR.egen.navn[id] + " er dessverre tomt.";
    sett[id] = 1;
  }
  if (e.ing.length > BAR.egen.maksTotalt) return "Maks " + BAR.egen.maksTotalt + " ingredienser.";
  if (BAR.egenAlkohol(e.ing).length > BAR.egen.maksSprit) return "Maks " + BAR.egen.maksSprit + " sorter brennevin eller likør.";
  var stilOk = false; BAR.egen.stiler.forEach(function (s) { if (s.id === e.stil) stilOk = true; });
  if (!stilOk) return "Velg hvordan du vil ha drinken.";
  return null;
};

/* Kort liste med ingrediensnavn i spiskammer-rekkefølge */
BAR.egenSortert = function (liste) {
  var ut = [];
  BAR.egen.kategorier.forEach(function (k) { k.ing.forEach(function (i) { if (liste.indexOf(i) >= 0) ut.push(i); }); });
  return ut;
};
BAR.egenTekst = function (liste) { return BAR.egenSortert(liste).map(function (i) { return BAR.egen.navn[i]; }).join(" · "); };

/* Forslag til bartenderen: mengder, metode og glass */
BAR.egenForslag = function (e) {
  var ing = BAR.egenSortert(e.ing || []);
  var har = function (i) { return ing.indexOf(i) >= 0; };
  var kat = function (k) { return ing.filter(function (i) { return BAR.egenKategori(i) === k; }); };
  var sprit = kat("sprit"), likor = kat("likor"), sitrus = kat("sitrus"), brus = kat("brus"), sott = kat("sott"), aroma = kat("aroma");
  var lang = e.stil === "lang" || (e.stil === "fri" && brus.length > 0);
  var halv = function (x) { return Math.max(0.5, Math.round(x * 2) / 2); };
  var linjer = [];
  var legg = function (i, m, tekst) { linjer.push({ i: i, m: m, tekst: tekst }); };
  // Alkohol: brennevin teller dobbelt så mye som likør
  var alkohol = sprit.concat(likor), total = lang ? 5 : 6;
  var vekt = 0; alkohol.forEach(function (i) { vekt += BAR.egenKategori(i) === "sprit" ? 2 : 1; });
  alkohol.forEach(function (i) { var m = halv(total * (BAR.egenKategori(i) === "sprit" ? 2 : 1) / vekt); legg(i, m, BAR.tall(m) + " cl " + BAR.egen.navn[i].toLowerCase().replace(/ \(.*\)$/, "")); });
  sitrus.forEach(function (i) {
    var m = i === "appelsinjuice" ? (lang ? (brus.length ? 6 : 10) : 3) : (lang ? 1.5 : 2);
    legg(i, m, BAR.tall(m) + " cl " + BAR.egen.navn[i].toLowerCase());
  });
  sott.forEach(function (i) {
    if (i === "egg") return legg(i, 1, "Eggehviten fra 1 egg");
    var m = i === "flote" ? 3 : i === "espresso" ? 3 : i === "grenadine" ? 1.5 : (lang ? 1 : 1.5);
    legg(i, m, BAR.tall(m) + " cl " + BAR.egen.navn[i].toLowerCase());
  });
  if (brus.length) {
    var bt = lang ? 12 : 3;
    brus.forEach(function (i) { var m = halv(bt / brus.length); legg(i, m, (lang ? "" : "En skvett ") + BAR.tall(m) + " cl " + BAR.egen.navn[i].toLowerCase()); });
  }
  aroma.forEach(function (i) {
    var t = { mynte: "Ca. 8 mynteblader", bitter: "2–3 dråper Angostura", limebat: "1 limebåt", appelsinskall: "1 appelsinskall", kaffebonner: "3 kaffebønner" }[i];
    legg(i, i === "mynte" ? 8 : i === "bitter" ? 3 : i === "kaffebonner" ? 3 : 1, t);
  });
  var metode = [];
  var ristes = har("egg") || har("flote") || har("espresso") || ((sitrus.length || sott.length) && !brus.length && !lang);
  if (har("mynte")) metode.push("Knus mynten forsiktig i bunnen av glasset (eller shakeren) først.");
  if (har("egg")) metode.push("Rist " + (brus.length ? "alt unntatt brus" : "alt") + " uten is i 10 sekunder – det gir skummet.");
  if (ristes) metode.push("Rist" + (brus.length ? " alt unntatt brus" : "") + " hardt med is i 10–12 sekunder og sil over i glasset" + (lang ? " med fersk is." : "."));
  else if (brus.length) metode.push("Fyll glasset med is og hell i alt unntatt brusen. Rør én gang.");
  else metode.push("Rør med is i 20 sekunder og sil over i glass med is.");
  if (brus.length) metode.push("Topp med " + brus.map(function (i) { return BAR.egen.navn[i].toLowerCase(); }).join(" og ") + ".");
  else if (lang && !har("appelsinjuice")) metode.push("Gjesten ville ha en lang drink, men valgte ingen brus – topp gjerne med litt soda.");
  var pynt = aroma.filter(function (i) { return i === "limebat" || i === "appelsinskall" || i === "kaffebonner"; });
  if (pynt.length) metode.push("Pynt med " + pynt.map(function (i) { return BAR.egen.navn[i].toLowerCase(); }).join(" og ") + ".");
  var glass = lang ? "Highball med is" : ((har("egg") || har("espresso")) && !brus.length ? "Coupe eller cocktailglass" : "Tumbler med is");
  return { linjer: linjer, metode: metode, glass: glass, alkoholfri: alkohol.length === 0, lang: lang };
};

/* ---------- Beregninger (brukes av kjøkkenets kalkulator og tester) ---------- */
BAR.finnDrink = function (id) { for (var k = 0; k < BAR.drinker.length; k++) if (BAR.drinker[k].id === id) return BAR.drinker[k]; return null; };

/* Formater tall norsk: 2.5 -> "2,5" */
BAR.tall = function (x, des) {
  if (des === undefined) des = 1;
  var f = Math.pow(10, des), r = Math.round(x * f) / f;
  return String(r).replace(".", ",");
};

/* Tekst for én ingredienslinje i en oppskrift, ganget med antall */
BAR.linjeTekst = function (linje, antall) {
  antall = antall || 1;
  var ing = BAR.ingredienser[linje.i];
  if (linje.vis && antall === 1) return linje.vis;
  var m = linje.m * antall;
  var navn = ing.kort || ing.navn;
  if (ing.enhet === "stk") return BAR.tall(m) + " " + (m > 1 && ing.kortN ? ing.kortN : navn);
  if (ing.enhet === "ts") return BAR.tall(m) + " ts " + navn;
  if (ing.enhet === "blader") return "ca. " + BAR.tall(m, 0) + " " + navn;
  if (ing.enhet === "dråper") return BAR.tall(m, 0) + " dråper " + navn;
  return BAR.tall(m) + " cl " + navn;
};

/* antall = { drinkId: n }. Returnerer forbruk per ingrediens og en handleliste. */
BAR.beregn = function (antall, opts) {
  opts = opts || {};
  var margin = opts.margin !== undefined ? opts.margin : BAR.sikkerhetsmargin;
  var str = opts.flaskestr || {};
  var forbruk = {}, totalDrinker = 0;
  BAR.drinker.forEach(function (d) {
    var n = Number(antall[d.id]) || 0;
    if (n <= 0) return;
    totalDrinker += n;
    d.oppskrift.forEach(function (l) {
      forbruk[l.i] = (forbruk[l.i] || 0) + l.m * n;
    });
  });

  var linjer = [], frukt = {}, sukkerG = 0, honningG = 0, kaffeG = 0;
  Object.keys(BAR.ingredienser).forEach(function (id) {
    var mengde = forbruk[id] || 0;
    if (!mengde) return;
    var ing = BAR.ingredienser[id], k = ing.kjop;
    if (k.type === "flaske") {
      var s = Number(str[id]) || k.str;
      var flasker = Math.ceil((mengde * (1 + margin)) / s - 1e-9);
      linjer.push({ id: id, navn: ing.navn, kat: ing.kat, mengde: mengde, enhet: ing.enhet, kjop: flasker, kjopTekst: k.tekst, str: s });
    } else if (k.type === "frukt") {
      frukt[k.frukt] = (frukt[k.frukt] || 0) + mengde / k.perFrukt;
    } else if (k.type === "sukker") {
      sukkerG += mengde * k.gPer;
    } else if (k.type === "honning") {
      honningG += mengde * k.gPer;
    } else if (k.type === "kaffe") {
      kaffeG += mengde * k.gPer;
    }
  });

  var andre = [];
  Object.keys(BAR.frukt).forEach(function (f) {
    if (!frukt[f]) return;
    var n = Math.ceil(frukt[f] * (1 + margin) - 1e-9);
    andre.push({ id: f, navn: BAR.frukt[f].navn, kjop: n, enhet: BAR.frukt[f].enhet });
  });
  if (sukkerG) andre.push({ id: "sukker_g", navn: "Sukker (til sukkersirup og mojito)", kjop: Math.ceil(sukkerG * (1 + margin) / 50) * 50, enhet: "g" });
  if (honningG) andre.push({ id: "honning_g", navn: "Honning (til honningsirup)", kjop: Math.ceil(honningG * (1 + margin) / 50) * 50, enhet: "g" });
  if (kaffeG) andre.push({ id: "kaffe_g", navn: "Espressokaffe (+ noen hele bønner til pynt)", kjop: Math.ceil(kaffeG * (1 + margin) / 10) * 10, enhet: "g" });
  var isKg = totalDrinker * BAR.isKgPerDrink * (1 + margin);
  if (totalDrinker) andre.push({ id: "is", navn: "Isbiter", kjop: Math.ceil(isKg / BAR.isPoseKg - 1e-9), enhet: "poser à " + BAR.isPoseKg + " kg", kg: isKg });

  return { totalDrinker: totalDrinker, forbruk: forbruk, flasker: linjer, andre: andre };
};

if (typeof module !== "undefined" && module.exports) module.exports = BAR;


var HTML_GJEST = "<p style=\"font-family:Georgia;padding:30px\">Gå til <a href=\"https://freddzy-jald.github.io/bar/\" target=\"_top\">freddzy-jald.github.io/bar</a></p>";
var HTML_KJOKKEN = HTML_GJEST;
