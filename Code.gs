
/**
 * KLETARSKA EVIDENCA – iPhone / Google Apps Script
 * Dostop je omejen na dovoljene Google račune.
 */

const CONFIG = {
  SPREADSHEET_ID: '1aZ4TiMAxB-OeiouAy1fbWTxDcLMWaJfv',

  SHEETS: {
    WINES: 'VINA',
    TANKS: 'CISTERNE',
    PROCEDURES: 'POSTOPKI',
    ANALYSES: 'ANALIZE',
    TRANSFERS: 'PRETOKI',
    BOTTLING: 'POLNITVE',
    SALES: 'PRODAJA',
    WAREHOUSE: 'SKLADISCE',
    GRAPE_PURCHASES: 'NAKUP_GROZDJA',
    LOOKUPS: 'SIFRANTI'
  }
};

const APP_VERSION = '2026-09-08-v22-diagnostika-kamere';

const ALLOWED_USERS = [
  'joze.malnaric@gmail.com',
  's.malnaric@gmail.com',
  'helena.malnaric@gmail.com',
  'jozemalnaric.semic@gmail.com'
];

function getCurrentUserEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function isAllowedUser_() {
  const email = getCurrentUserEmail_();
  return email !== '' && ALLOWED_USERS.includes(email);
}

function requireAllowedUser_() {
  const email = getCurrentUserEmail_();

  if (!email) {
    throw new Error(
      'Google ni posredoval e-poštnega naslova prijavljenega uporabnika. ' +
      'Preveri, da je Web App objavljen kot "Execute as: User accessing the web app".'
    );
  }

  if (!ALLOWED_USERS.includes(email)) {
    throw new Error('Za račun ' + email + ' dostop do aplikacije ni dovoljen.');
  }

  return email;
}

function testDostopa() {
  const email = getCurrentUserEmail_();
  Logger.log('APP_VERSION: ' + APP_VERSION);
  Logger.log('Prijavljen uporabnik: ' + email);
  Logger.log('Dovoljen: ' + ALLOWED_USERS.includes(email));
  Logger.log('Dovoljeni uporabniki: ' + ALLOWED_USERS.join(', '));
  return {
    version: APP_VERSION,
    email: email,
    allowed: ALLOWED_USERS.includes(email),
    allowedUsers: ALLOWED_USERS
  };
}

function doGet(e) {
  const email = getCurrentUserEmail_();

  if (!email) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:60px auto;padding:24px">' +
      '<h2>Kletarska evidenca - Vinska klet Lipec</h2>' +
      '<p>Google ni mogel določiti prijavljenega uporabnika.</p>' +
      '<p>Preveri, da si prijavljen v Google račun in da je aplikacija objavljena kot ' +
      '<b>Execute as: User accessing the web app</b>.</p>' +
      '</div>'
    ).setTitle('Kletarska evidenca - Vinska klet Lipec');
  }

  if (!ALLOWED_USERS.includes(email)) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:60px auto;padding:24px">' +
      '<h2>Dostop ni dovoljen</h2>' +
      '<p style="color:#777;font-size:12px">Verzija: ' + APP_VERSION + '</p>' +
      '<p>Račun <b>' + escapeHtmlServer_(email) + '</b> nima dovoljenja za uporabo aplikacije.</p>' +
      '<p>Prijavi se z dovoljenim Google računom.</p>' +
      '</div>'
    ).setTitle('Kletarska evidenca - Vinska klet Lipec');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.initialTank =
    (e && e.parameter && e.parameter.cisterna)
      ? String(e.parameter.cisterna)
      : '';

  template.initialStock =
    (e && e.parameter && e.parameter.zaloga)
      ? String(e.parameter.zaloga)
      : '';

  return template.evaluate()
    .setTitle('Kletarska evidenca - Vinska klet Lipec')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

function escapeHtmlServer_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getSs_() {
  requireAllowedUser_();

  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('V CONFIG.SPREADSHEET_ID manjka ID Google preglednice.');
  }

  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function sheet_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh) throw new Error('Manjka zavihek "' + name + '" v Google preglednici.');
  return sh;
}

function ensureSalesSheet_() {
  requireAllowedUser_();

  const ss = getSs_();
  let sh = ss.getSheetByName(CONFIG.SHEETS.SALES);

  const requiredHeaders = [
    'ID_prodaje',
    'Datum',
    'ID_vina',
    'ID_zaloge',
    'Vrsta_prodaje',
    'ID_cisterne',
    'Nacin_prodaje',
    'Serija_lot',
    'Stevilo_kosov',
    'Volumen_kos_L',
    'Kolicina_L',
    'Cena_enota_EUR',
    'Cena_skupaj_EUR',
    'Kupec',
    'Kupec_naslov',
    'Kupec_posta',
    'Kupec_davcna',
    'Nacin_placila',
    'Datum_placila',
    'Izvajalec',
    'Odstej_od_cisterne',
    'Opomba'
  ];

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.SALES);
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sh.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return sh;
  }

  // Starega zavihka PRODAJA ne brišemo. Če manjkata nova stolpca,
  // ju dodamo na konec in s tem ohranimo vse obstoječe zapise.
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existingHeaders = sh.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(String);

  const missing = requiredHeaders.filter(h => !existingHeaders.includes(h));

  if (missing.length) {
    sh.getRange(1, existingHeaders.length + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
  }

  sh.setFrozenRows(1);
  return sh;
}

function ensureWarehouseSheet_() {
  requireAllowedUser_();

  const ss = getSs_();
  let sh = ss.getSheetByName(CONFIG.SHEETS.WAREHOUSE);

  const requiredHeaders = [
    'ID_zaloge',
    'Datum',
    'ID_vina',
    'Naziv_vina',
    'Letnik',
    'Tip_embalaze',
    'Kolicina_L',
    'Stevilo_kosov',
    'Zacetno_stevilo_kosov',
    'Volumen_kos_L',
    'Serija_lot',
    'Status',
    'QR_oznaka',
    'Opomba'
  ];

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.WAREHOUSE);
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sh.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return sh;
  }

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existingHeaders = sh.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(String);

  const missing = requiredHeaders.filter(h => !existingHeaders.includes(h));

  if (missing.length) {
    sh.getRange(1, existingHeaders.length + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
  }

  sh.setFrozenRows(1);
  return sh;
}



function ensureGrapePurchaseSheet_() {
  requireAllowedUser_();

  const ss = getSs_();
  let sh = ss.getSheetByName(CONFIG.SHEETS.GRAPE_PURCHASES);

  const requiredHeaders = [
    'ID_nakupa',
    'Datum',
    'Dobavitelj',
    'Izvor_vinograd',
    'Sorta',
    'Letnik',
    'Kolicina_kg',
    'Cena_kg_EUR',
    'Cena_skupaj_EUR',
    'Oe_Brix',
    'ID_vina',
    'Opomba',
    'Izvajalec'
  ];

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.GRAPE_PURCHASES);
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sh.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return sh;
  }

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existingHeaders = sh.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(String);

  const missing = requiredHeaders.filter(h => !existingHeaders.includes(h));

  if (missing.length) {
    sh.getRange(1, existingHeaders.length + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
  }

  sh.setFrozenRows(1);
  return sh;
}

function grapePurchases_() {
  ensureGrapePurchaseSheet_();

  return valuesAsObjects_(CONFIG.SHEETS.GRAPE_PURCHASES)
    .sort((a, b) => {
      const da = String(a.Datum || '');
      const db = String(b.Datum || '');
      if (db !== da) return db.localeCompare(da);
      return String(a.Dobavitelj || '').localeCompare(
        String(b.Dobavitelj || ''),
        'sl'
      );
    });
}

function grapeOrigins_() {
  const values = grapePurchases_()
    .map(x => String(x.Izvor_vinograd || '').trim())
    .filter(Boolean);

  return [...new Set(values)].sort((a, b) => a.localeCompare(b, 'sl'));
}


function grapeSuppliers_() {
  const values = grapePurchases_()
    .map(x => String(x.Dobavitelj || '').trim())
    .filter(Boolean);

  return [...new Set(values)].sort((a, b) => a.localeCompare(b, 'sl'));
}

function saveGrapePurchase(form) {
  requireAllowedUser_();
  ensureGrapePurchaseSheet_();
  ensureWineSupplierColumn_();

  const kg = Number(normalizeNumber_(form.Kolicina_kg) || 0);
  const priceKg = Number(normalizeNumber_(form.Cena_kg_EUR) || 0);

  if (kg <= 0) {
    throw new Error('Vnesi količino kupljenega grozdja v kg.');
  }

  let total = normalizeNumber_(form.Cena_skupaj_EUR);

  if ((total === '' || Number(total) === 0) && priceKg > 0) {
    total = kg * priceKg;
  }

  const existingId = String(form.ID_nakupa || '').trim();

  const obj = {
    ID_nakupa: existingId || uuid_('GRZ'),
    Datum: form.Datum || today_(),
    Dobavitelj: form.Dobavitelj || '',
    Izvor_vinograd: form.Izvor_vinograd || '',
    Sorta: form.Sorta || '',
    Letnik: normalizeNumber_(form.Letnik),
    Kolicina_kg: kg,
    Cena_kg_EUR: priceKg || '',
    Cena_skupaj_EUR: normalizeNumber_(total),
    Oe_Brix: normalizeNumber_(form.Oe_Brix),
    ID_vina: form.ID_vina || '',
    Opomba: form.Opomba || '',
    Izvajalec: getCurrentUserEmail_()
  };

  if (existingId) {
    const updated = setCellByKey_(
      CONFIG.SHEETS.GRAPE_PURCHASES,
      'ID_nakupa',
      existingId,
      {
        Datum: obj.Datum,
        Dobavitelj: obj.Dobavitelj,
        Izvor_vinograd: obj.Izvor_vinograd,
        Sorta: obj.Sorta,
        Letnik: obj.Letnik,
        Kolicina_kg: obj.Kolicina_kg,
        Cena_kg_EUR: obj.Cena_kg_EUR,
        Cena_skupaj_EUR: obj.Cena_skupaj_EUR,
        Oe_Brix: obj.Oe_Brix,
        ID_vina: obj.ID_vina,
        Opomba: obj.Opomba,
        Izvajalec: obj.Izvajalec
      }
    );

    if (!updated) {
      throw new Error('Zapisa za popravek ni bilo mogoče najti.');
    }
  } else {
    appendObject_(CONFIG.SHEETS.GRAPE_PURCHASES, obj);
  }

  // Poveži nakup grozdja z vinom/moštom.
  // V VINA se shranita tako izvor/vinograd kot dobavitelj grozdja.
  if (obj.ID_vina) {
    const wine = activeWines_()
      .find(w => String(w.ID_vina) === String(obj.ID_vina));

    if (!wine) {
      throw new Error(
        (existingId ? 'Popravek' : 'Nakup') +
        ' je shranjen, vendar izbranega vina/mošta ni bilo mogoče najti.'
      );
    }

    const updates = {
      Izvor_vinograd: obj.Izvor_vinograd || '',
      Dobavitelj_grozdja: obj.Dobavitelj || ''
    };

    if (!String(wine.Naziv_vina || '').trim() && obj.Sorta) {
      updates.Naziv_vina = obj.Sorta;
    }

    setCellByKey_(
      CONFIG.SHEETS.WINES,
      'ID_vina',
      obj.ID_vina,
      updates
    );
  }

  return {
    ok: true,
    message:
      (existingId
        ? 'Nakup grozdja je popravljen.'
        : 'Nakup grozdja je shranjen.') +
      (obj.ID_vina
        ? ' Dobavitelj in izvor/vinograd sta povezana z izbranim vinom/moštom.'
        : ''),
    item: obj
  };
}


function ensureBottlingSheet_() {
  requireAllowedUser_();

  const ss = getSs_();
  let sh = ss.getSheetByName(CONFIG.SHEETS.BOTTLING);

  const requiredHeaders = [
    'ID_polnitve',
    'Datum',
    'ID_vina',
    'Nacin',
    'Volumen_embalaze_L',
    'Stevilo_kosov',
    'Skupaj_L',
    'Serija_lot',
    'Prosti_SO2_mgL',
    'Sladkor_gL',
    'Opomba'
  ];

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.BOTTLING);
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sh.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return sh;
  }

  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existingHeaders = sh.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(String);

  const missing = requiredHeaders.filter(h => !existingHeaders.includes(h));

  if (missing.length) {
    sh.getRange(1, existingHeaders.length + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
  }

  sh.setFrozenRows(1);
  return sh;
}


function salesItems_() {
  ensureSalesSheet_();

  return valuesAsObjects_(CONFIG.SHEETS.SALES)
    .sort((a, b) => {
      const da = String(a.Datum || '');
      const db = String(b.Datum || '');
      if (db !== da) return db.localeCompare(da);
      return String(b.ID_prodaje || '').localeCompare(String(a.ID_prodaje || ''));
    });
}

function warehouseItems_() {
  ensureWarehouseSheet_();
  return valuesAsObjects_(CONFIG.SHEETS.WAREHOUSE)
    .sort((a, b) => {
      const da = String(a.Datum || '');
      const db = String(b.Datum || '');
      if (db !== da) return db.localeCompare(da);
      return String(a.Naziv_vina || '').localeCompare(String(b.Naziv_vina || ''), 'sl');
    });
}

function isBottlingProcedure_(name) {
  const value = String(name || '').trim().toLowerCase();
  return (
    value === 'stekleničenje' ||
    value === 'steklenicenje' ||
    value === 'polnitev' ||
    value.includes('steklenič') ||
    value.includes('steklenic') ||
    value.includes('polnitev')
  );
}

function moveWineToWarehouse_(procedureObj) {
  ensureWarehouseSheet_();

  const wines = activeWines_();
  const wine = wines.find(w =>
    String(w.ID_vina) === String(procedureObj.ID_vina)
  );

  if (!wine) {
    throw new Error('Vino za prenos v skladišče ni bilo najdeno.');
  }

  const currentTank = String(wine.Trenutna_cisterna || '');
  const sourceTank = String(procedureObj.ID_cisterne || '');

  if (sourceTank && currentTank && sourceTank !== currentTank) {
    throw new Error(
      'Izbrano vino trenutno ni v cisterni ' + sourceTank + '.'
    );
  }

  const currentQty = Number(
    String(wine.Trenutna_kolicina_L || '0').replace(',', '.')
  ) || 0;

  let bottledQty = Number(procedureObj.Kolicina_vina_L || 0);

  if (bottledQty <= 0) bottledQty = currentQty;

  if (bottledQty <= 0) {
    throw new Error('Za stekleničenje ni vpisane količine vina.');
  }

  if (bottledQty > currentQty) {
    throw new Error(
      'Količina za stekleničenje (' + bottledQty +
      ' L) je večja od trenutne količine v cisterni (' +
      currentQty + ' L).'
    );
  }

  const stockId = uuid_('SKL');

  const stockObj = {
    ID_zaloge: stockId,
    Datum: procedureObj.Datum || today_(),
    ID_vina: wine.ID_vina || '',
    Naziv_vina: wine.Naziv_vina || '',
    Letnik: wine.Letnik || '',
    Kolicina_L: bottledQty,
    Stevilo_kosov: '',
    Volumen_kos_L: '',
    Serija_lot: '',
    Status: 'Na zalogi',
    QR_oznaka: stockId,
    Opomba:
      'Prenos iz cisterne ' +
      (sourceTank || currentTank || '') +
      ' po postopku ' +
      (procedureObj.Postopek || 'Stekleničenje')
  };

  appendObject_(CONFIG.SHEETS.WAREHOUSE, stockObj);

  const remaining = Math.max(0, currentQty - bottledQty);

  const updates = {
    Trenutna_kolicina_L: remaining
  };

  if (remaining === 0) {
    updates.Trenutna_cisterna = '';
    updates.Status = 'Ustekleničeno';
  }

  setCellByKey_(
    CONFIG.SHEETS.WINES,
    'ID_vina',
    wine.ID_vina,
    updates
  );

  return {
    stockId: stockId,
    bottledQty: bottledQty,
    remainingQty: remaining
  };
}


function valuesAsObjects_(name) {
  const sh = sheet_(name);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  const data = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = data.shift().map(String);

  return data
    .filter(row => row.some(v => String(v).trim() !== ''))
    .map(row => {
      const o = {};
      headers.forEach((h, i) => o[h] = row[i]);
      return o;
    });
}

function headers_(name) {
  const sh = sheet_(name);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

function appendObject_(name, obj) {
  const sh = sheet_(name);
  const headers = headers_(name);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sh.appendRow(row);
  return obj;
}

function setCellByKey_(sheetName, keyColumn, keyValue, updates) {
  const sh = sheet_(sheetName);
  const data = sh.getDataRange().getValues();

  if (data.length < 2) return false;

  const headers = data[0].map(String);
  const keyIndex = headers.indexOf(keyColumn);

  if (keyIndex < 0) {
    throw new Error('Manjka ključni stolpec ' + keyColumn);
  }

  const updateIndexes = {};
  Object.keys(updates).forEach(k => {
    const idx = headers.indexOf(k);
    if (idx >= 0) updateIndexes[k] = idx;
  });

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][keyIndex]) === String(keyValue)) {
      Object.keys(updateIndexes).forEach(k => {
        sh.getRange(r + 1, updateIndexes[k] + 1).setValue(updates[k]);
      });
      return true;
    }
  }

  return false;
}

function uuid_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function today_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Europe/Ljubljana',
    'yyyy-MM-dd'
  );
}

function nowTime_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Europe/Ljubljana',
    'HH:mm'
  );
}

function normalizeNumber_(value) {
  if (value === null || value === undefined || value === '') return '';
  const n = Number(String(value).replace(',', '.'));
  return isNaN(n) ? value : n;
}


function ensureWineSupplierColumn_() {
  const sh = sheet_(CONFIG.SHEETS.WINES);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(String);

  if (!headers.includes('Dobavitelj_grozdja')) {
    sh.getRange(1, lastCol + 1).setValue('Dobavitelj_grozdja');
    sh.getRange(1, lastCol + 1)
      .setFontWeight('bold')
      .setBackground('#1F4E3D')
      .setFontColor('#FFFFFF');
  }
}

function activeWines_() {
  return valuesAsObjects_(CONFIG.SHEETS.WINES)
    .filter(w => String(w.Aktivno || 'DA').toUpperCase() !== 'NE')
    .sort((a, b) =>
      String(a.Naziv_vina || '').localeCompare(
        String(b.Naziv_vina || ''),
        'sl'
      )
    );
}

function activeTanks_() {
  return valuesAsObjects_(CONFIG.SHEETS.TANKS)
    .filter(t => String(t.Aktivna || 'DA').toUpperCase() !== 'NE')
    .sort((a, b) => {
      const na = Number(a.Naziv);
      const nb = Number(b.Naziv);

      if (!isNaN(na) && !isNaN(nb)) return na - nb;

      return String(a.Naziv || '').localeCompare(
        String(b.Naziv || ''),
        'sl'
      );
    });
}

function getBootstrapData() {
  requireAllowedUser_();
  ensureSalesSheet_();
  ensureWarehouseSheet_();
  ensureGrapePurchaseSheet_();
  ensureWineSupplierColumn_();

  const lookups = valuesAsObjects_(CONFIG.SHEETS.LOOKUPS);

  const listFromColumn = key =>
    [...new Set(
      lookups
        .map(r => r[key])
        .filter(v => String(v).trim() !== '')
    )];

  return {
    today: today_(),
    nowTime: nowTime_(),
    currentUser: getCurrentUserEmail_(),
    webAppUrl: ScriptApp.getService().getUrl() || '',
    wines: activeWines_(),
    tanks: activeTanks_(),
    warehouse: warehouseItems_(),
    sales: salesItems_(),
    grapePurchases: grapePurchases_(),
    grapeOrigins: grapeOrigins_(),
    grapeSuppliers: grapeSuppliers_(),
    lookups: {
      procedures: listFromColumn('Postopki'),
      products: listFromColumn('Preparati'),
      units: listFromColumn('Enote'),
      analysisTypes: listFromColumn('Tipi_analiz'),
      statuses: listFromColumn('Statusi'),
      wineNames: listFromColumn('Sorte')
    }
  };
}

function getWineDetail(idVina) {
  requireAllowedUser_();
  ensureSalesSheet_();

  const wines = activeWines_();
  const wine = wines.find(w => String(w.ID_vina) === String(idVina));

  if (!wine) throw new Error('Vino ni bilo najdeno.');

  const procedures = valuesAsObjects_(CONFIG.SHEETS.PROCEDURES)
    .filter(x => String(x.ID_vina) === String(idVina))
    .map(x => ({ ...x, _type: 'Postopek' }));

  const analyses = valuesAsObjects_(CONFIG.SHEETS.ANALYSES)
    .filter(x => String(x.ID_vina) === String(idVina))
    .map(x => ({ ...x, _type: 'Analiza' }));

  const transfers = valuesAsObjects_(CONFIG.SHEETS.TRANSFERS)
    .filter(x => String(x.ID_vina) === String(idVina))
    .map(x => ({ ...x, _type: 'Pretok' }));

  const sales = valuesAsObjects_(CONFIG.SHEETS.SALES)
    .filter(x => String(x.ID_vina) === String(idVina))
    .map(x => ({ ...x, _type: 'Prodaja' }));

  const history = [...procedures, ...analyses, ...transfers, ...sales]
    .sort((a, b) => String(b.Datum || '').localeCompare(String(a.Datum || '')));

  return { wine, history };
}

function saveWine(form) {
  requireAllowedUser_();

  const obj = {
    ID_vina: uuid_('VIN'),
    Letnik: normalizeNumber_(form.Letnik),
    Naziv_vina: form.Naziv_vina || '',
    Tip: form.Tip || '',
    Datum_sprejema: form.Datum_sprejema || today_(),
    Izvor_vinograd: form.Izvor_vinograd || '',
    Dobavitelj_grozdja: form.Dobavitelj_grozdja || '',
    Zacetna_kolicina_L: normalizeNumber_(form.Zacetna_kolicina_L),
    Trenutna_cisterna: form.Trenutna_cisterna || '',
    Status: form.Status || 'Mošt',
    Trenutna_kolicina_L: normalizeNumber_(form.Zacetna_kolicina_L),
    Zacetni_Oe: normalizeNumber_(form.Zacetni_Oe),
    Ciljni_sladkor_gL: normalizeNumber_(form.Ciljni_sladkor_gL),
    Opomba: form.Opomba || '',
    Aktivno: 'DA'
  };

  appendObject_(CONFIG.SHEETS.WINES, obj);

  return {
    ok: true,
    message: 'Vino je shranjeno.',
    item: obj
  };
}

function saveProcedure(form) {
  requireAllowedUser_();

  if (isBottlingProcedure_(form.Postopek)) {
    throw new Error(
      'Stekleničenje je ločen postopek. Uporabi gumb "Stekleničenje" pri vinu ali cisterni.'
    );
  }

  const obj = {
    ID_postopka: uuid_('POS'),
    Datum: form.Datum || today_(),
    Cas: form.Cas || nowTime_(),
    ID_vina: form.ID_vina || '',
    ID_cisterne: form.ID_cisterne || '',
    Postopek: form.Postopek || '',
    Preparat: form.Preparat || '',
    Kolicina_preparata: normalizeNumber_(form.Kolicina_preparata),
    Enota: form.Enota || '',
    Kolicina_vina_L: normalizeNumber_(form.Kolicina_vina_L),
    Temperatura_C: normalizeNumber_(form.Temperatura_C),
    Oe_Brix: normalizeNumber_(form.Oe_Brix),
    Prosti_SO2_mgL: normalizeNumber_(form.Prosti_SO2_mgL),
    Izvajalec: form.Izvajalec || getCurrentUserEmail_(),
    Opomba: form.Opomba || ''
  };

  appendObject_(CONFIG.SHEETS.PROCEDURES, obj);

  return {
    ok: true,
    message: 'Postopek je shranjen.',
    item: obj
  };
}

function saveBottling(form) {
  requireAllowedUser_();
  ensureWarehouseSheet_();
  ensureBottlingSheet_();

  const wineId = String(form.ID_vina || '');
  const tankId = String(form.ID_cisterne || '');

  const wine = activeWines_()
    .find(w => String(w.ID_vina) === wineId);

  if (!wine) {
    throw new Error('Vino za stekleničenje ni bilo najdeno.');
  }

  if (!tankId || String(wine.Trenutna_cisterna || '') !== tankId) {
    throw new Error('Izbrano vino trenutno ni v izbrani cisterni.');
  }

  const currentQty = Number(
    String(wine.Trenutna_kolicina_L || '0').replace(',', '.')
  ) || 0;

  const pieces = Number(normalizeNumber_(form.Stevilo_kosov) || 0);
  const volumePerPiece = Number(normalizeNumber_(form.Volumen_kos_L) || 0);

  if (pieces <= 0) {
    throw new Error('Vnesi število napolnjenih kosov.');
  }

  if (volumePerPiece <= 0) {
    throw new Error('Vnesi volumen embalaže v litrih.');
  }

  const bottledQty = pieces * volumePerPiece;

  if (bottledQty > currentQty + 0.0001) {
    throw new Error(
      'Stekleničena količina (' + bottledQty.toFixed(2) +
      ' L) je večja od trenutne količine v cisterni (' +
      currentQty.toFixed(2) + ' L).'
    );
  }

  const lot = String(form.Serija_lot || '').trim() ||
    ('LOT-' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || 'Europe/Ljubljana',
        'yyyyMMdd'
      ) +
      '-' +
      String(wine.ID_vina || '').slice(-4).toUpperCase()
    );

  const stockId = uuid_('SKL');
  const bottlingId = uuid_('POL');
  const procedureId = uuid_('POS');

  const packaging = String(form.Tip_embalaze || 'Steklenica');

  const bottlingObj = {
    ID_polnitve: bottlingId,
    Datum: form.Datum || today_(),
    ID_vina: wine.ID_vina || '',
    Nacin: packaging,
    Volumen_embalaze_L: volumePerPiece,
    Stevilo_kosov: pieces,
    Skupaj_L: bottledQty,
    Serija_lot: lot,
    Prosti_SO2_mgL: normalizeNumber_(form.Prosti_SO2_mgL),
    Sladkor_gL: normalizeNumber_(form.Sladkor_gL),
    Opomba: form.Opomba || ''
  };

  appendObject_(CONFIG.SHEETS.BOTTLING, bottlingObj);

  const procedureObj = {
    ID_postopka: procedureId,
    Datum: form.Datum || today_(),
    Cas: nowTime_(),
    ID_vina: wine.ID_vina || '',
    ID_cisterne: tankId,
    Postopek: 'Stekleničenje',
    Preparat: '',
    Kolicina_preparata: '',
    Enota: '',
    Kolicina_vina_L: bottledQty,
    Temperatura_C: '',
    Oe_Brix: '',
    Prosti_SO2_mgL: normalizeNumber_(form.Prosti_SO2_mgL),
    Izvajalec: getCurrentUserEmail_(),
    Opomba:
      'LOT ' + lot + ' · ' +
      pieces + ' kos × ' + volumePerPiece + ' L' +
      (form.Opomba ? ' · ' + form.Opomba : '')
  };

  appendObject_(CONFIG.SHEETS.PROCEDURES, procedureObj);

  const stockObj = {
    ID_zaloge: stockId,
    Datum: form.Datum || today_(),
    ID_vina: wine.ID_vina || '',
    Naziv_vina: wine.Naziv_vina || '',
    Letnik: wine.Letnik || '',
    Tip_embalaze: packaging,
    Kolicina_L: bottledQty,
    Stevilo_kosov: pieces,
    Zacetno_stevilo_kosov: pieces,
    Volumen_kos_L: volumePerPiece,
    Serija_lot: lot,
    Status: 'Na zalogi',
    QR_oznaka: stockId,
    Opomba:
      'Stekleničeno iz cisterne ' + tankId +
      (form.Opomba ? ' · ' + form.Opomba : '')
  };

  appendObject_(CONFIG.SHEETS.WAREHOUSE, stockObj);

  const remaining = Math.max(0, currentQty - bottledQty);

  const updates = {
    Trenutna_kolicina_L: remaining
  };

  if (remaining <= 0.0001) {
    updates.Trenutna_kolicina_L = 0;
    updates.Trenutna_cisterna = '';
    updates.Status = 'Ustekleničeno';
  }

  setCellByKey_(
    CONFIG.SHEETS.WINES,
    'ID_vina',
    wine.ID_vina,
    updates
  );

  return {
    ok: true,
    message:
      'Stekleničenje je shranjeno: ' +
      pieces + ' kos × ' + volumePerPiece + ' L = ' +
      bottledQty.toFixed(2) + ' L. ' +
      'LOT: ' + lot + '. ' +
      (remaining <= 0.0001
        ? 'Cisterna je sproščena.'
        : 'V cisterni ostane ' + remaining.toFixed(2) + ' L.'),
    item: bottlingObj,
    stock: stockObj
  };
}


function saveAnalysis(form) {
  requireAllowedUser_();

  const obj = {
    ID_analize: uuid_('ANA'),
    Datum: form.Datum || today_(),
    ID_vina: form.ID_vina || '',
    Tip_analize: '',
    Temperatura_C: normalizeNumber_(form.Temperatura_C),
    Oe_Brix: normalizeNumber_(form.Oe_Brix),
    pH: normalizeNumber_(form.pH),
    Skupne_kisline_gL: normalizeNumber_(form.Skupne_kisline_gL),
    Prosti_SO2_mgL: normalizeNumber_(form.Prosti_SO2_mgL),
    Skupni_SO2_mgL: normalizeNumber_(form.Skupni_SO2_mgL),
    Sladkor_gL: normalizeNumber_(form.Sladkor_gL),
    Alkohol_vol: normalizeNumber_(form.Alkohol_vol),
    Hlapne_kisline_gL: normalizeNumber_(form.Hlapne_kisline_gL),
    Opomba: form.Opomba || ''
  };

  appendObject_(CONFIG.SHEETS.ANALYSES, obj);

  return {
    ok: true,
    message: 'Analiza je shranjena.',
    item: obj
  };
}

function saveTransfer(form) {
  requireAllowedUser_();

  const amount = normalizeNumber_(form.Kolicina_L);
  const loss = normalizeNumber_(form.Izguba_L) || 0;

  const obj = {
    ID_pretoka: uuid_('PRE'),
    Datum: form.Datum || today_(),
    ID_vina: form.ID_vina || '',
    Iz_cisterne: form.Iz_cisterne || '',
    V_cisterno: form.V_cisterno || '',
    Kolicina_L: amount,
    Izguba_L: loss,
    Razlog: form.Razlog || '',
    Izvajalec: form.Izvajalec || '',
    Opomba: form.Opomba || ''
  };

  appendObject_(CONFIG.SHEETS.TRANSFERS, obj);

  const fullTransfer =
    String(form.Celoten_pretok) === 'true' ||
    form.Celoten_pretok === true;

  if (fullTransfer) {
    setCellByKey_(
      CONFIG.SHEETS.WINES,
      'ID_vina',
      form.ID_vina,
      {
        Trenutna_cisterna: form.V_cisterno || '',
        Trenutna_kolicina_L:
          Math.max(
            0,
            Number(amount || 0) - Number(loss || 0)
          )
      }
    );
  }

  return {
    ok: true,
    message: fullTransfer
      ? 'Pretok je shranjen in trenutna cisterna vina je posodobljena.'
      : 'Pretok je shranjen.',
    item: obj
  };
}

function saveSale(form) {
  requireAllowedUser_();
  ensureSalesSheet_();
  ensureWarehouseSheet_();

  const saleType = String(form.Vrsta_prodaje || 'Prodaja iz skladišča').trim();
  const isTankSale =
    saleType === 'Prodaja iz cisterne' ||
    saleType === 'Odprema iz cisterne';

  const unitPrice = Number(normalizeNumber_(form.Cena_enota_EUR) || 0);
  let totalPrice = normalizeNumber_(form.Cena_skupaj_EUR);

  if (isTankSale) {
    const quantity = Number(normalizeNumber_(form.Kolicina_L) || 0);
    let tankId = form.ID_cisterne || '';

    if (!tankId && form.ID_vina) {
      const wineForTank = activeWines_()
        .find(w => String(w.ID_vina) === String(form.ID_vina));
      if (wineForTank) tankId = wineForTank.Trenutna_cisterna || '';
    }

    if (!form.ID_vina) {
      throw new Error('Pri prodaji iz cisterne manjka vino.');
    }
    if (!tankId) {
      throw new Error('Pri prodaji iz cisterne manjka cisterna.');
    }
    if (quantity <= 0) {
      throw new Error('Vnesi količino za prodajo v litrih.');
    }

    const wine = activeWines_()
      .find(w => String(w.ID_vina) === String(form.ID_vina));

    if (!wine) {
      throw new Error('Izbrano vino ni bilo najdeno med aktivnimi vini.');
    }

    if (String(wine.Trenutna_cisterna || '') !== String(tankId)) {
      throw new Error('Izbrano vino trenutno ni v izbrani cisterni.');
    }

    const current = Number(
      String(wine.Trenutna_kolicina_L || '0').replace(',', '.')
    ) || 0;

    if (quantity > current + 0.0001) {
      throw new Error(
        'Prodana količina (' + quantity +
        ' L) je večja od trenutne količine v cisterni (' +
        current + ' L).'
      );
    }

    if ((totalPrice === '' || Number(totalPrice) === 0) && unitPrice > 0) {
      totalPrice = quantity * unitPrice;
    }

    const obj = {
      ID_prodaje: uuid_('PRO'),
      Datum: form.Datum || today_(),
      ID_vina: form.ID_vina || '',
      ID_zaloge: '',
      Vrsta_prodaje: 'Prodaja iz cisterne',
      ID_cisterne: tankId,
      Nacin_prodaje: 'Odprto vino / prodaja iz cisterne',
      Serija_lot: '',
      Stevilo_kosov: '',
      Volumen_kos_L: '',
      Kolicina_L: quantity,
      Cena_enota_EUR: unitPrice || '',
      Cena_skupaj_EUR: normalizeNumber_(totalPrice),
      Kupec: form.Kupec || '',
      Kupec_naslov: form.Kupec_naslov || '',
      Kupec_posta: form.Kupec_posta || '',
      Kupec_davcna: form.Kupec_davcna || '',
      Nacin_placila: form.Nacin_placila || '',
      Datum_placila: form.Datum_placila || '',
      Izvajalec: getCurrentUserEmail_(),
      Odstej_od_cisterne: 'DA',
      Opomba: form.Opomba || ''
    };

    appendObject_(CONFIG.SHEETS.SALES, obj);

    const remaining = Math.max(0, current - quantity);
    const updates = { Trenutna_kolicina_L: remaining };

    if (remaining <= 0.0001) {
      updates.Trenutna_kolicina_L = 0;
      updates.Trenutna_cisterna = '';
    }

    setCellByKey_(
      CONFIG.SHEETS.WINES,
      'ID_vina',
      obj.ID_vina,
      updates
    );

    return {
      ok: true,
      message: 'Prodaja iz cisterne je shranjena in količina vina je zmanjšana.',
      item: obj
    };
  }

  // Prodaja iz skladišča je vedno vezana na točno ID_zaloge / LOT.
  const stockId = String(form.ID_zaloge || '').trim();

  if (!stockId) {
    throw new Error('Izberi zalogo oziroma LOT za prodajo.');
  }

  const stock = warehouseItems_()
    .find(x => String(x.ID_zaloge) === stockId);

  if (!stock) {
    throw new Error('Izbrana zaloga ni bila najdena.');
  }

  const availablePieces = Number(
    String(stock.Stevilo_kosov || '0').replace(',', '.')
  ) || 0;

  const pieces = Number(normalizeNumber_(form.Stevilo_kosov) || 0);
  const volumePerPiece = Number(
    String(stock.Volumen_kos_L || '0').replace(',', '.')
  ) || 0;

  if (availablePieces <= 0) {
    throw new Error('Izbrani LOT nima več zaloge.');
  }

  if (pieces <= 0) {
    throw new Error('Vnesi število prodanih kosov.');
  }

  if (pieces > availablePieces) {
    throw new Error(
      'Na zalogi je samo ' + availablePieces +
      ' kos, prodati pa želiš ' + pieces + ' kos.'
    );
  }

  if (volumePerPiece <= 0) {
    throw new Error(
      'Pri izbrani zalogi manjka volumen embalaže. ' +
      'Preveri zapis v zavihku SKLADISCE.'
    );
  }

  const quantity = pieces * volumePerPiece;

  if ((totalPrice === '' || Number(totalPrice) === 0) && unitPrice > 0) {
    totalPrice = pieces * unitPrice;
  }

  const obj = {
    ID_prodaje: uuid_('PRO'),
    Datum: form.Datum || today_(),
    ID_vina: stock.ID_vina || '',
    ID_zaloge: stock.ID_zaloge || '',
    Vrsta_prodaje: 'Prodaja iz skladišča',
    ID_cisterne: '',
    Nacin_prodaje: stock.Tip_embalaze || 'Skladišče',
    Serija_lot: stock.Serija_lot || '',
    Stevilo_kosov: pieces,
    Volumen_kos_L: volumePerPiece,
    Kolicina_L: quantity,
    Cena_enota_EUR: unitPrice || '',
    Cena_skupaj_EUR: normalizeNumber_(totalPrice),
    Kupec: form.Kupec || '',
    Kupec_naslov: form.Kupec_naslov || '',
    Kupec_posta: form.Kupec_posta || '',
    Kupec_davcna: form.Kupec_davcna || '',
    Nacin_placila: form.Nacin_placila || '',
    Datum_placila: form.Datum_placila || '',
    Izvajalec: getCurrentUserEmail_(),
    Odstej_od_cisterne: 'NE',
    Opomba: form.Opomba || ''
  };

  appendObject_(CONFIG.SHEETS.SALES, obj);

  const remainingPieces = availablePieces - pieces;
  const remainingLiters = remainingPieces * volumePerPiece;

  setCellByKey_(
    CONFIG.SHEETS.WAREHOUSE,
    'ID_zaloge',
    stockId,
    {
      Stevilo_kosov: remainingPieces,
      Kolicina_L: remainingLiters,
      Status: remainingPieces > 0 ? 'Na zalogi' : 'Prodano'
    }
  );

  return {
    ok: true,
    message:
      'Prodaja iz skladišča je shranjena. ' +
      'LOT ' + (stock.Serija_lot || stockId) +
      ': ostane ' + remainingPieces + ' kos.',
    item: obj
  };
}
