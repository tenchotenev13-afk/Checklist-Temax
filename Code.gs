// ТеМАХ Контролна карта — Apps Script (коригирана версия)
// Поставете ЦЕЛИЯ този код в единствения файл в Apps Script
// Важно: след деплой изберете "Execute as: Me" и "Who has access: Anyone"

var SS_ID      = '1yIYEhc69iJYryU0VooPWSj5WGOgRjs9mcJ7qlQpmlg8';
var SS_RAW     = 'Всички проверки';
var SS_SUMMARY = 'Резюме по магазин';

var SEC_IDS = ['s1','s2','s3','s4','s5','s6','s7','s8','s9','s10','s11','s12','s13','s14','s15','s16','s17'];

// Колони в листа "Всички проверки" (1-базирани):
// 1:Дата/час  2:Магазин  3:Дата  4:Управител  5:Обща%  6:Бонус  7:Блокиращ
// 8-24: s1-s17 (обща% на сектор = СРЕДНО(общи%, специф.%))
// 25-41: s1-s17 общи% (от Лист 1)
// 42-58: s1-s17 специф.% (от Лист 2)
// 59:Блок.1  60:Блок.2  61:Блок.3  62:Блок.4
// 63:Забележки  64:Повторни грешки  65:Срок/Отговорник

var COL_COMBINED_START = 8;   // s1 комбинирана %
var COL_GEN_START      = 25;  // s1 обща %
var COL_SPEC_START     = 42;  // s1 специфична %
var COL_BLOCK_START    = 59;  // Блок.1
var COL_NOTES          = 63;
var COL_REPEAT         = 64;
var COL_DEADLINE       = 65;
var TOTAL_COLS         = 65;

// ── MAIN ENTRY POINT ────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e.parameter ? e.parameter.action   : null;
    var cb     = e.parameter ? e.parameter.callback : null;
    var data   = e.parameter ? e.parameter.data     : null;

    // Таблото иска всички данни
    if (action === 'getAll') {
      return sendJSON(getAllRows(), cb);
    }

    // Чеклистът изпраща нова проверка
    if (data) {
      // Apps Script вече декодира URL параметрите — не декодираме втори път!
      try {
        var parsed = JSON.parse(data);
        if (!parsed.shop || !parsed.date) {
          return sendJSON({status:'error', message:'Липсват задължителни полета (shop/date)'}, cb);
        }
        writeRow(parsed);
        return sendJSON({status:'ok', message:'Записано успешно'}, cb);
      } catch(parseErr) {
        // Логваме грешката в отделен лист за диагностика
        try {
          var ss = SpreadsheetApp.openById(SS_ID);
          var errSheet = ss.getSheetByName('_errors_') || ss.insertSheet('_errors_');
          errSheet.appendRow([new Date(), 'parse error', parseErr.toString(), String(data).slice(0,200)]);
        } catch(e) {}
        return sendJSON({status:'error', message:'JSON parse грешка: ' + parseErr.toString()}, cb);
      }
    }

    // Тест на записването — ?action=testWrite
    if (action === 'testWrite') {
      try {
        var ss = SpreadsheetApp.openById(SS_ID);
        var testSheet = ss.getSheetByName('_test_');
        if (!testSheet) testSheet = ss.insertSheet('_test_');
        testSheet.appendRow(['тест', new Date().toString()]);
        return sendJSON({status:'ok', message:'Записването работи! Провери лист _test_ в Sheets.'}, cb);
      } catch(err) {
        return sendJSON({status:'error', message:'Грешка при запис: ' + err.toString()}, cb);
      }
    }

    // Тест на пълния writeRow — ?action=testSave
    if (action === 'testSave') {
      try {
        var fakeData = {
          shop: 'ТЕСТмагазин', date: '2026-05-12', manager: 'Тест Тестов',
          avgPct: 0.85, bonusOk: true, blockFail: false,
          blocking: {b1:'DA', b2:'DA', b3:'DA', b4:'DA'},
          sectorData: (function(){
            var sd = {};
            ['s1','s2','s3','s4','s5','s6','s7','s8','s9','s10','s11','s12','s13','s14','s15','s16','s17'].forEach(function(s){
              sd[s] = {combinedPct: 0.85, genPct: 0.83, specPct: 0.87};
            });
            return sd;
          })(),
          notes: 'тест коментар', repeatIssues: '', deadlines: ''
        };
        writeRow(fakeData);
        return sendJSON({status:'ok', message:'testSave OK — провери лист Всички проверки в Sheets!'}, cb);
      } catch(err) {
        return sendJSON({status:'error', message:'testSave грешка: ' + err.toString()}, cb);
      }
    }

    // Изпращане на пълен имейл отчет — ?action=sendEmail&to=...&shop=...&date=...
    if (action === 'sendEmail') {
      try {
        var to      = e.parameter.to || '';
        var shop    = e.parameter.shop || '';
        var date    = e.parameter.date || '';
        var manager = e.parameter.manager || '';
        if (!to) return sendJSON({status:'error', message:'Липсва имейл адрес'}, cb);

        // Четем данните от Sheets
        var ss2    = SpreadsheetApp.openById(SS_ID);
        var sheet2 = ss2.getSheetByName(SS_RAW);
        var entry  = null;
        if (sheet2 && sheet2.getLastRow() > 1) {
          var lastColE  = sheet2.getLastColumn();
          var readColsE = Math.min(lastColE, TOTAL_COLS);
          var rowsE     = sheet2.getRange(2, 1, sheet2.getLastRow()-1, readColsE).getValues();
          // Намираме по магазин + дата
          for (var ri = rowsE.length - 1; ri >= 0; ri--) {
            var rowDate = rowsE[ri][2];
            if (rowDate instanceof Date) {
              rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } else {
              rowDate = String(rowDate).slice(0, 10);
            }
            if (String(rowsE[ri][1]).trim() === shop.trim() && rowDate === date) {
              entry = rowsE[ri]; break;
            }
          }
          // Fallback: последен запис за магазина
          if (!entry) {
            for (var ri2 = rowsE.length - 1; ri2 >= 0; ri2--) {
              if (String(rowsE[ri2][1]).trim() === shop.trim()) {
                entry = rowsE[ri2]; break;
              }
            }
          }
        }

        var secNames = ['Плочки и ламинат','Строителство','Баня и ВиК','Текстил',
          'Подови настилки','Дом. потреби','Бои и облицовки','Ел. и градин. машини',
          'Градина','Градин. мебели','Крепежи, обков и врати','Входна зона и фасада',
          'Вътрешен склад','Външен склад','Плац','Касова зона и бюра','Осветление'];

        var colorF = function(p) {
          if (p===null||p===undefined||p==='') return '#94A3B8';
          return p>=85?'#0A7C59':p>=60?'#D97706':'#C0392B';
        };

        var overallPct = entry ? parsePercent(entry[4]) : null;
        var bonusOk    = entry ? entry[5]==='ДА ✅' : false;
        var bonusColor = bonusOk ? '#0A7C59' : '#C0392B';

        // Сектори
        var secRows = '';
        if (entry) {
          for (var si = 0; si < 17; si++) {
            var combined = parsePercent(entry[COL_COMBINED_START-1+si]);
            var gen      = parsePercent(entry[COL_GEN_START-1+si]);
            var spec     = parsePercent(entry[COL_SPEC_START-1+si]);
            var bgS  = combined===null?'#F8FAFC':combined>=85?'#E6F9F1':combined>=60?'#FEF3C7':'#FDE8E8';
            var colS = colorF(combined);
            var genSpec = (gen!==null?gen+'%':'—') + ' / ' + (spec!==null?spec+'%':'—');
            var status  = combined===null?'—':combined>=85?'✅ ОК':combined>=60?'⚠ Внимание':'❌ Слаб';
            secRows += '<tr style="background:'+bgS+'">'
              +'<td style="padding:6px 10px;font-size:11px;border-bottom:1px solid #E2E8F0">'+secNames[si]+'</td>'
              +'<td style="padding:6px 10px;font-size:11px;text-align:center;color:#64748B;border-bottom:1px solid #E2E8F0">'+genSpec+'</td>'
              +'<td style="padding:6px 10px;font-size:12px;font-weight:700;text-align:center;color:'+colS+';border-bottom:1px solid #E2E8F0">'+(combined!==null?combined+'%':'—')+'</td>'
              +'<td style="padding:6px 10px;font-size:11px;color:#64748B;border-bottom:1px solid #E2E8F0">'+status+'</td>'
              +'</tr>';
          }
        }

        // Блокиращи
        var blockRows = '';
        if (entry) {
          var bTexts = [
            'Касови продажби. При под 30% от средния брой касови продажби от справката в Т-бюлетина — бонусът отпада.',
            'Проверка на всички цели палета (изтекъл срок)',
            'Всички с униформи и баджове + проверка от камери',
            'Изрядна документация и пожарогасители в срок'
          ];
          for (var bi = 0; bi < 4; bi++) {
            var bVal = entry[COL_BLOCK_START-1+bi]||'-';
            var bCol = bVal==='DA'?'#0A7C59':bVal==='NE'?'#C0392B':'#94A3B8';
            var bTxt = bVal==='DA'?'✅ ДА':bVal==='NE'?'❌ НЕ':'—';
            blockRows += '<tr><td style="padding:5px 10px;font-size:11px;border-bottom:1px solid #F1F5F9">'+bTexts[bi]+'</td>'
              +'<td style="padding:5px 10px;font-weight:700;color:'+bCol+';border-bottom:1px solid #F1F5F9">'+bTxt+'</td></tr>';
          }
        }

        var notes     = entry?(entry[COL_NOTES-1]||''):'';
        var repeats   = entry?(entry[COL_REPEAT-1]||''):'';
        var deadlines = entry?(entry[COL_DEADLINE-1]||''):'';

        // Забележки — таблица като в PDF отчета
        var notesHtml = '';
        if ((notes&&notes!=='-') || (repeats&&repeats!=='-') || (deadlines&&deadlines!=='-')) {
          // Събираме всички критерии с коментари
          var noteItems = [];
          if (notes&&notes!=='-') {
            notes.split(' | ').forEach(function(n) {
              if (!n) return;
              var colonIdx = n.indexOf(': ');
              var criterion = colonIdx>0 ? n.slice(0,colonIdx) : n;
              var comment   = colonIdx>0 ? n.slice(colonIdx+2) : '';
              var isRepeat  = repeats && repeats.indexOf(criterion) >= 0;
              var deadline  = '';
              if (deadlines&&deadlines!=='-') {
                deadlines.split(' | ').forEach(function(d){
                  if (d.indexOf(criterion+': ')===0) deadline = d.slice(criterion.length+2);
                });
              }
              noteItems.push({criterion:criterion, comment:comment, isRepeat:isRepeat, deadline:deadline});
            });
          }
          if (noteItems.length) {
            notesHtml += '<div style="font-size:11px;font-weight:700;color:#374151;margin:14px 0 6px;text-transform:uppercase;letter-spacing:.04em">📝 Забележки и коментари</div>';
            notesHtml += '<table style="width:100%;border-collapse:collapse">';
            notesHtml += '<tr style="background:#0D1B2A">'
              +'<th style="padding:6px 8px;font-size:9px;color:#94A3B8;text-align:left;font-weight:700;text-transform:uppercase">Критерий</th>'
              +'<th style="padding:6px 8px;font-size:9px;color:#94A3B8;text-align:left;font-weight:700;text-transform:uppercase">Коментар</th>'
              +'<th style="padding:6px 8px;font-size:9px;color:#94A3B8;text-align:center;font-weight:700;text-transform:uppercase">Повт. грешка</th>'
              +'<th style="padding:6px 8px;font-size:9px;color:#94A3B8;text-align:left;font-weight:700;text-transform:uppercase">Срок / Отговорник</th>'
              +'</tr>';
            noteItems.forEach(function(item, idx) {
              var bg = idx%2===0?'#F8FAFC':'#FFFFFF';
              var repeatCell = item.isRepeat
                ? '<span style="color:#C0392B;font-weight:700">🔁 ДА</span>'
                : '<span style="color:#94A3B8">НЕ</span>';
              notesHtml += '<tr style="background:'+bg+';border-bottom:1px solid #E2E8F0">'
                +'<td style="padding:6px 8px;font-size:10px;font-weight:600;color:#374151">'+item.criterion+'</td>'
                +'<td style="padding:6px 8px;font-size:10px;color:#374151">'+item.comment+'</td>'
                +'<td style="padding:6px 8px;font-size:10px;text-align:center">'+repeatCell+'</td>'
                +'<td style="padding:6px 8px;font-size:10px;color:'+( item.deadline?'#B45309':'#94A3B8')+'">'+( item.deadline||'—')+'</td>'
                +'</tr>';
            });
            notesHtml += '</table>';
          }
        }

        var subject = 'ТеМАХ Контролна карта — '+shop+' — '+date;
        var html = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto">'
          +'<div style="background:#0D1B2A;padding:18px 22px;border-radius:8px 8px 0 0;display:flex;justify-content:space-between;align-items:center">'
          +'<div><div style="font-size:18px;font-weight:800;color:#fff">ТеМАХ — Контролна карта</div>'
          +'<div style="font-size:11px;color:#64748B;margin-top:2px">'+date+' · '+(manager||'—')+'</div></div>'
          +'<div style="text-align:right"><div style="font-size:26px;font-weight:800;color:'+bonusColor+'">'+(overallPct!==null?overallPct+'%':'—')+'</div>'
          +'<div style="font-size:12px;font-weight:700;color:'+bonusColor+'">'+(bonusOk?'✅ БОНУС ОДОБРЕН':'✗ БЕЗ БОНУС')+'</div></div>'
          +'</div>'
          +'<div style="background:#fff;padding:10px 22px;border:1px solid #E2E8F0;border-top:none">'
          +'<span style="font-size:15px;font-weight:700;color:#0D1B2A">'+shop+'</span></div>'
          +(blockRows?'<table style="width:100%;border-collapse:collapse;border:1px solid #E2E8F0;border-top:none">'
          +'<tr style="background:#FFF5F5"><th colspan="2" style="padding:7px 10px;font-size:10px;font-weight:700;color:#C0392B;text-align:left;text-transform:uppercase">⛔ Блокиращи критерии</th></tr>'
          +blockRows+'</table>':'')
          +'<table style="width:100%;border-collapse:collapse;border:1px solid #E2E8F0;border-top:none">'
          +'<tr style="background:#0D1B2A">'
          +'<th style="padding:7px 10px;font-size:10px;color:#94A3B8;text-align:left;font-weight:700;text-transform:uppercase">Сектор</th>'
          +'<th style="padding:7px 10px;font-size:10px;color:#94A3B8;text-align:center;font-weight:700;text-transform:uppercase">Общи / Специф.</th>'
          +'<th style="padding:7px 10px;font-size:10px;color:#94A3B8;text-align:center;font-weight:700;text-transform:uppercase">Обща %</th>'
          +'<th style="padding:7px 10px;font-size:10px;color:#94A3B8;text-align:left;font-weight:700;text-transform:uppercase">Статус</th>'
          +'</tr>'+secRows+'</table>'
          +'<div style="padding:14px;border:1px solid #E2E8F0;border-top:none">'+notesHtml
          +'<p style="margin-top:14px;font-size:10px;color:#94A3B8">Изпратено автоматично от ТеМАХ Контролна карта · TeMAX</p>'
          +'</div></div>';

        MailApp.sendEmail({to:to, subject:subject, htmlBody:html});
        return sendJSON({status:'ok', message:'Имейлът е изпратен към '+to}, cb);
      } catch(err) {
        return sendJSON({status:'error', message:'Грешка при имейл: '+err.toString()}, cb);
      }
    }

    // Записване на снимки — ?action=addPhotos&shop=...&date=...&photos=...
    if (action === 'addPhotos') {
      try {
        var shopP   = e.parameter.shop || '';
        var dateP   = e.parameter.date || '';
        var photosJ = e.parameter.photos || '{}';
        var ss3     = SpreadsheetApp.openById(SS_ID);
        var sheet3  = ss3.getSheetByName(SS_RAW);
        if (sheet3 && sheet3.getLastRow() > 1) {
          var lastC3 = sheet3.getLastColumn();
          var rows3  = sheet3.getRange(2, 1, sheet3.getLastRow()-1, Math.min(lastC3, 3)).getValues();
          for (var ri3 = rows3.length - 1; ri3 >= 0; ri3--) {
            var rd3 = rows3[ri3][2];
            if (rd3 instanceof Date) rd3 = Utilities.formatDate(rd3, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            else rd3 = String(rd3).slice(0,10);
            if (String(rows3[ri3][1]).trim() === shopP.trim() && rd3 === dateP) {
              // Expand to col 66 if needed
              if (sheet3.getMaxColumns() < 66) sheet3.insertColumnsAfter(sheet3.getMaxColumns(), 66 - sheet3.getMaxColumns());
              sheet3.getRange(ri3 + 2, 66).setValue(photosJ);
              return sendJSON({status:'ok', message:'Снимките са записани'}, cb);
            }
          }
        }
        return sendJSON({status:'error', message:'Записът не е намерен'}, cb);
      } catch(err) {
        return sendJSON({status:'error', message:'Грешка: ' + err.toString()}, cb);
      }
    }

    // Проверка дали API работи
    return sendJSON({status:'ok', message:'ТеМАХ API е активен!'}, cb);

  } catch(err) {
    // Важно: при грешка също връщаме JSONP, иначе callback не се извиква
    return sendJSON({status:'error', message:err.toString()}, cb);
  }
}

function sendJSON(obj, cb) {
  var json = JSON.stringify(obj);
  var out  = cb ? cb + '(' + json + ')' : json;
  var mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(out).setMimeType(mime);
}


// ── POST handler (от чеклиста — fetch POST) ─────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    writeRow(data);
    var output = ContentService.createTextOutput(JSON.stringify({status:'ok'}));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch(err) {
    var output = ContentService.createTextOutput(JSON.stringify({status:'error', message:err.toString()}));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

// ── ЧЕТЕНЕ НА ДАННИ ЗА ТАБЛОТО ──────────────────────────────────
// КОРЕКЦИЯ: чете TOTAL_COLS колони (беше само 30 — липсваха notes, repeatIssues, deadlines)
function getAllRows() {
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SS_RAW);
  if (!sheet || sheet.getLastRow() < 2) return {status:'ok', data:[]};

  var lastRow  = sheet.getLastRow() - 1;
  var lastCol  = sheet.getLastColumn(); // динамично — не надвишаваме броя колони
  var readCols = Math.min(lastCol, TOTAL_COLS);
  var rows     = sheet.getRange(2, 1, lastRow, readCols).getValues();

  var data = rows
    .filter(function(r) { return r[1]; }) // само редове с магазин
    .map(function(r) {
      // Комбинирана % по сектор (СРЕДНО на общи + специфични)
      var sectors = {};
      SEC_IDS.forEach(function(sid, i) {
        var val = r[COL_COMBINED_START - 1 + i];
        sectors[sid] = parsePercent(val);
      });

      // Общи % и специф. % по сектор (за по-детайлен анализ)
      var sectorsGen = {}, sectorsSpec = {};
      SEC_IDS.forEach(function(sid, i) {
        sectorsGen[sid]  = parsePercent(r[COL_GEN_START  - 1 + i]);
        sectorsSpec[sid] = parsePercent(r[COL_SPEC_START - 1 + i]);
      });

      // Блокиращи
      var blocking = {
        b1: r[COL_BLOCK_START - 1]     || '-',
        b2: r[COL_BLOCK_START - 1 + 1] || '-',
        b3: r[COL_BLOCK_START - 1 + 2] || '-',
        b4: r[COL_BLOCK_START - 1 + 3] || '-',
      };

      return {
        savedAt:      r[0] ? new Date(r[0]).toISOString() : '',
        shop:         String(r[1] || ''),  // force string — числови имена (123) се четат като number
        // Конвертираме Date обект към локална дата (България UTC+3)
        // Без това "2026-05-13" → "2026-05-12T21:00:00.000Z" → показва се като 12-ти
        date: (function(v){
          if (!v) return '';
          if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          return String(v).slice(0, 10);
        })(r[2]),
        manager:      r[3] || '',
        overallPct:   parsePercent(r[4]),           // обща % (СРЕДНО на 17 сектора)
        bonusOk:      r[5] === 'ДА ✅',
        blockFail:    r[6] === 'ДА ⛔',
        sectors:      sectors,
        sectorsGen:   sectorsGen,
        sectorsSpec:  sectorsSpec,
        blocking:     blocking,
        // КОРЕКЦИЯ: тези полета липсваха — таблото не можеше да филтрира по тях
        notes:        r[COL_NOTES    - 1] || '',
        repeatIssues: r[COL_REPEAT   - 1] || '',
        deadlines:    r[COL_DEADLINE - 1] || '',
        photos:       (function(v){ try{ return v ? JSON.parse(v) : {}; } catch(e){ return {}; } })(r[65]),
      };
    });

  return {status:'ok', data:data};
}

// Помощна функция — обработва всички формати от Google Sheets:
// "87%" (string) → 87  |  0.87 (decimal от Sheets) → 87  |  87 (integer) → 87
function parsePercent(val) {
  if (val === null || val === undefined || val === '' || val === '-') return null;
  if (typeof val === 'number') {
    // Google Sheets конвертира "87%" → 0.87 при appendRow
    return val > 1 ? Math.round(val) : Math.round(val * 100);
  }
  var str = String(val).replace('%','').trim();
  if (!str || str === '-') return null;
  var num = parseFloat(str);
  if (isNaN(num)) return null;
  return num > 1 ? Math.round(num) : Math.round(num * 100);
}

// ── ЗАПИС НА НОВА ПРОВЕРКА ──────────────────────────────────────
function writeRow(d) {
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SS_RAW);
  if (!sheet) {
    sheet = ss.insertSheet(SS_RAW);
    try { addRawHeaders(sheet); } catch(he) { Logger.log('Header error: ' + he); }
  }
  if (sheet.getLastRow() === 0) {
    try { addRawHeaders(sheet); } catch(he) { Logger.log('Header error: ' + he); }
  }

  // Обща % (СРЕДНО на 17 сектора × 100, закръглено)
  var overallPct = (d.avgPct !== null && d.avgPct !== undefined)
    ? Math.round(d.avgPct * 100)
    : '';

  // Базов ред: Дата/час, Магазин, Дата, Управител, Обща%, Бонус, Блокиращ
  var row = [
    new Date(),
    d.shop    || '',
    d.date    || '',
    d.manager || '',
    overallPct,
    d.bonusOk   ? 'ДА ✅' : 'НЕ ✗',
    d.blockFail ? 'ДА ⛔' : 'НЕ',
  ];

  // Колони 8-24: комбинирана % по сектор — като числа (не string с %)
  SEC_IDS.forEach(function(sid) {
    var sd = d.sectorData && d.sectorData[sid];
    row.push((sd && sd.combinedPct != null) ? Math.round(sd.combinedPct * 100) : '');
  });

  // Колони 25-41: общи % по сектор
  SEC_IDS.forEach(function(sid) {
    var sd = d.sectorData && d.sectorData[sid];
    row.push((sd && sd.genPct != null) ? Math.round(sd.genPct * 100) : '');
  });

  // Колони 42-58: специфични % по сектор
  SEC_IDS.forEach(function(sid) {
    var sd = d.sectorData && d.sectorData[sid];
    row.push((sd && sd.specPct != null) ? Math.round(sd.specPct * 100) : '');
  });

  // Поддържа и новия формат (d.blocking) и стария (d.state.blocking)
  var bl = d.blocking || (d.state && d.state.blocking) || {};
  row.push(
    bl['b1'] || '-',
    bl['b2'] || '-',
    bl['b3'] || '-',
    bl['b4'] || '-'
  );

  // Забележки, повторни грешки, срок
  row.push(
    d.notes        || '-',
    d.repeatIssues || '-',
    d.deadlines    || '-'
  );

  // Запис на реда
  sheet.appendRow(row);

  // Форматиране на новия ред
  var lr  = sheet.getLastRow();
  var bg  = d.bonusOk ? '#E6F9F1' : (d.avgPct && d.avgPct < 0.6 ? '#FDE8E8' : '#FEF9E7');
  sheet.getRange(lr, 1, 1, row.length).setBackground(bg);
  sheet.getRange(lr, 5).setFontWeight('bold');                                              // обща %
  sheet.getRange(lr, 6).setFontColor(d.bonusOk ? '#0A7C59' : '#C0392B').setFontWeight('bold'); // бонус

  // Актуализация на резюмето по магазин
  updateSummary(d, ss, overallPct);
}

// ── РЕЗЮМЕ ПО МАГАЗИН ───────────────────────────────────────────
function updateSummary(d, ss, overallPct) {
  var sum = ss.getSheetByName(SS_SUMMARY);
  if (!sum) { sum = ss.insertSheet(SS_SUMMARY); addSumHeaders(sum); }
  if (sum.getLastRow() === 0) addSumHeaders(sum);

  var shop = d.shop || '';
  var sr   = -1;
  if (sum.getLastRow() > 1) {
    var col = sum.getRange(2, 1, sum.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      if (col[i][0] === shop) { sr = i + 2; break; }
    }
  }

  var sumOverallPct = (d.avgPct !== null && d.avgPct !== undefined) ? Math.round(d.avgPct * 100) : '';
  var sumRow = [shop, d.date || '', d.manager || '', sumOverallPct, d.bonusOk ? 'ДА ✅' : 'НЕ ✗'];

  // Комбинирана % по сектор — като числа
  SEC_IDS.forEach(function(sid) {
    var sd = d.sectorData && d.sectorData[sid];
    sumRow.push((sd && sd.combinedPct != null) ? Math.round(sd.combinedPct * 100) : '');
  });

  if (sr > 0) {
    sum.getRange(sr, 1, 1, sumRow.length).setValues([sumRow]);
  } else {
    sum.appendRow(sumRow);
    sr = sum.getLastRow();
  }

  var bg = d.bonusOk ? '#E6F9F1' : '#FEF9E7';
  sum.getRange(sr, 1, 1, sumRow.length).setBackground(bg);
  sum.getRange(sr, 4).setFontWeight('bold');
  sum.getRange(sr, 5).setFontColor(d.bonusOk ? '#0A7C59' : '#C0392B').setFontWeight('bold');
}

// ── ХЕДЪРИ ──────────────────────────────────────────────────────
function addRawHeaders(s) {
  var secNames = ['Плочки','Строителство','Баня','Текстил','Подови настилки',
    'Дом. потреби','Бои','Ел. машини','Градина','Градин. мебели','Крепежи',
    'Входна зона','Вът. склад','Външ. склад','Плац','Касова зона','Осветление'];

  var h = ['Дата/час','Магазин','Дата','Управител','Обща %','Бонус','Блокиращ'];

  // Комбинирани % (СРЕДНО на общи + специф.)
  secNames.forEach(function(n) { h.push(n + ' %'); });

  // Общи % (от Лист 1)
  secNames.forEach(function(n) { h.push(n + ' (общи%)'); });

  // Специфични % (от Лист 2)
  secNames.forEach(function(n) { h.push(n + ' (специф.%)'); });

  // Блокиращи
  h.push('Блок.1','Блок.2','Блок.3','Блок.4');

  // Забележки
  h.push('Забележки','Повторни грешки','Срок/Отговорник');

  // Разширяваме Sheet-а до необходимите колони (нов Sheet има само 26 по default)
  var neededCols = h.length;
  var currentCols = s.getMaxColumns();
  if (currentCols < neededCols) {
    s.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  s.getRange(1, 1, 1, h.length).setValues([h])
   .setBackground('#0D1B2A')
   .setFontColor('#FFFFFF')
   .setFontWeight('bold')
   .setFontSize(10);

  s.setFrozenRows(1);
  s.setFrozenColumns(2);

  // Ширини на колоните
  s.setColumnWidth(1, 130); // дата/час
  s.setColumnWidth(2, 160); // магазин
  s.setColumnWidth(3, 90);  // дата
  s.setColumnWidth(4, 140); // управител
  s.setColumnWidth(5, 70);  // обща %
  s.setColumnWidth(6, 80);  // бонус
  s.setColumnWidth(7, 80);  // блокиращ

  for (var i = 8; i <= 58; i++) s.setColumnWidth(i, 75);
  for (var j = 59; j <= 62; j++) s.setColumnWidth(j, 65);
  s.setColumnWidth(63, 300);
  s.setColumnWidth(64, 200);
  s.setColumnWidth(65, 200);

  // Цветни групи заглавия
  s.getRange(1, COL_COMBINED_START, 1, 17).setBackground('#1864AB'); // комб. сини
  s.getRange(1, COL_GEN_START,      1, 17).setBackground('#2F9E44'); // общи зелени
  s.getRange(1, COL_SPEC_START,     1, 17).setBackground('#9C36B5'); // специф. лилави
  s.getRange(1, COL_BLOCK_START,    1,  4).setBackground('#C92A2A'); // блокиращи червени
  s.getRange(1, COL_NOTES,          1,  3).setBackground('#495057'); // забележки сиви
}

function addSumHeaders(s) {
  var secNames = ['Плочки','Строителство','Баня','Текстил','Подови настилки',
    'Дом. потреби','Бои','Ел. машини','Градина','Градин. мебели','Крепежи',
    'Входна зона','Вят. склад','Външ. склад','Плац','Касова зона','Осветление'];

  var h = ['Магазин','Последна проверка','Управител','Обща %','Бонус'];
  secNames.forEach(function(n) { h.push(n + ' %'); });

  // Разширяваме Sheet-а до необходимите колони (нов Sheet има само 26 по default)
  var neededCols = h.length;
  var currentCols = s.getMaxColumns();
  if (currentCols < neededCols) {
    s.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  var neededColsS = h.length;
  var currentColsS = s.getMaxColumns();
  if (currentColsS < neededColsS) {
    s.insertColumnsAfter(currentColsS, neededColsS - currentColsS);
  }

  s.getRange(1, 1, 1, h.length).setValues([h])
   .setBackground('#0D1B2A')
   .setFontColor('#FFFFFF')
   .setFontWeight('bold')
   .setFontSize(10);

  s.setFrozenRows(1);
  s.setColumnWidth(1, 160);
  s.setColumnWidth(2, 110);
  s.setColumnWidth(3, 140);
  s.setColumnWidth(4, 75);
  s.setColumnWidth(5, 80);
  for (var i = 6; i <= h.length; i++) s.setColumnWidth(i, 75);
}
