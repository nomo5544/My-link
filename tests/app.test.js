// Запуск: npm install && npm test   (потрібен jsdom; bibleText.json має лежати в корені проєкту)
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { JSDOM } = require('jsdom');

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const script = fs.readFileSync(path.join(root, 'textPage.js'), 'utf8');
const importScript = fs.readFileSync(path.join(root, 'importFile.js'), 'utf8');
const fixture = (name) => fs.readFileSync(path.join(__dirname, 'fixtures', name));
const toArrayBuffer = (b) => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
const bible = JSON.parse(fs.readFileSync(path.join(root, 'bibleText.json'), 'utf8'));

async function makeApp(store, touch) {
    const dom = new JSDOM(html, { runScripts: 'outside-only', url: 'http://localhost/' });
    await new Promise(r => dom.window.addEventListener('load', r));
    if (touch) dom.window.matchMedia = (q) => ({ matches: !/hover:\s*hover/.test(q) });
    for (const k in (store || {})) dom.window.localStorage.setItem(k, store[k]);
    dom.window.fflate = require('fflate');
    dom.window.TextDecoder = require('util').TextDecoder; // у справжньому браузері він є
    dom.window.eval(script + '\nwindow.__bookLookup = bookLookup;');
    dom.window.eval(importScript);
    return dom.window;
}

let passed = 0, failed = 0;
const queue = [];
function test(name, fn) { queue.push({ name, fn }); }
function section(title) { queue.push({ title }); }
// масиви з jsdom належать іншому realm, тому порівнюємо через JSON
const same = (a, b) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b));

let w;
const refs = (text) => Array.from(w.findReferences(text)).map(r => [r.book, r.chapter, r.verses].join('|'));

section('Розпізнавання посилань');
test('базові форми', () => {
    same(refs('Матвія 5:3'), ['Від Матвія|5|3']);
    same(refs('Буття 1.1'), ['Буття|1|1']);
    same(refs('2Кор 3.4-6'), ['2 до Коринтян|3|4-6']);
    same(refs('Рим. 8:28'), ['До Римлян|8|28']);
});
test('«1Петр 4:11» та «1 Петра 4:11»', () => {
    same(refs('1Петр 4:11'), ['1 Петра|4|11']);
    same(refs('1 Петра 4:11'), ['1 Петра|4|11']);
});
test('«Об\'явлення» з прямим і типографським апострофом', () => {
    same(refs("Об'явлення 21:4").map(x => x.split('|').slice(1).join('|')), ['21|4']);
    same(refs('Об’явлення 21:4').map(x => x.split('|').slice(1).join('|')), ['21|4']);
});
test('списки віршів і довге тире', () => {
    same(refs('Мт 5:3,5'), ['Від Матвія|5|3,5']);
    same(refs('Мт 5:3—5'), ['Від Матвія|5|3-5']);
    same(refs('Мт 5:3-5, 7'), ['Від Матвія|5|3-5,7']);
});
test('короткі скорочення розпізнаються без обмежень: «як 5», «ам 3», «об 8», «іс 53»', () => {
    same(refs('Я купив як 5 яблук').map(x => x.split('|').slice(1).join('|')), ['5|']);
    same(refs('ам 3').map(x => x.split('|').slice(1).join('|')), ['3|']);
    same(refs('Це було об 8 годині').map(x => x.split('|').slice(1).join('|')), ['8|']);
    same(refs('іс 53:5').map(x => x.split('|').slice(1).join('|')), ['53|5']);
    assert.ok(refs('Я купив як 5 яблук')[0].startsWith('Якова'), refs('Я купив як 5 яблук')[0]);
    assert.ok(refs('ам 3')[0].startsWith('Амос'));
    assert.ok(refs('Як 5:16')[0].startsWith('Якова'));
});
test('слова, що лише містять скорочення, як і раніше, не чіпаються', () => {
    same(refs('Усіх 3 рази'), []);
    same(refs('Сам 5 разів'), []);
});
test('слово-частина не дає Амос (Филип’янам)', () => {
    const r = refs('Филип’янам 4:13');
    assert.ok(r.every(x => !x.startsWith('Амос')), r.join());
});
section('DOM: посилання, розділене тегами, та атрибути');
test('<b>Мт</b> 5:3 розпізнається', () => {
    const el = w.document.createElement('div');
    el.innerHTML = '<p><b>Мт</b> 5:3 і ще</p>';
    w.linkifyReferences(el);
    const spans = el.querySelectorAll('.bible-link');
    assert.ok(spans.length >= 1);
    assert.strictEqual(spans[0].dataset.book, 'Від Матвія');
    assert.strictEqual(el.textContent, 'Мт 5:3 і ще');
});
test('атрибути не чіпаються', () => {
    const el = w.document.createElement('div');
    el.innerHTML = '<span title="Ос 5 Мт 5:3" data-x="Пс 23:1">текст</span>';
    w.linkifyReferences(el);
    assert.strictEqual(el.querySelectorAll('.bible-link').length, 0);
    assert.strictEqual(el.firstChild.getAttribute('title'), 'Ос 5 Мт 5:3');
});
test('кілька посилань в одному вузлі', () => {
    const el = w.document.createElement('div');
    el.textContent = 'Пс 23:1 та Ів 3:16 і Рим 8:28';
    w.linkifyReferences(el);
    assert.strictEqual(el.querySelectorAll('.bible-link').length, 3);
    assert.strictEqual(el.textContent, 'Пс 23:1 та Ів 3:16 і Рим 8:28');
});

section('Санітизація');
test('прибирає скрипти, картинки, обробники подій', () => {
    const out = w.sanitizeHtml('<p onclick="x()">Привіт<script>alert(1)</script><img src=x onerror=alert(1)></p>').innerHTML;
    assert.strictEqual(out, '<p>Привіт</p>');
});
test('javascript: у посиланнях блокується, https дозволений', () => {
    assert.ok(!/href/.test(w.sanitizeHtml('<a href="javascript:alert(1)">x</a>').innerHTML));
    assert.ok(/href="https:\/\/a\.b"/.test(w.sanitizeHtml('<a href="https://a.b">x</a>').innerHTML));
});
test('стилі: лишаються кольори, url() прибирається', () => {
    const out = w.sanitizeHtml('<span style="color: red; background: url(x); font-weight: bold; position: fixed">a</span>').innerHTML;
    assert.ok(out.includes('color: red') && out.includes('font-weight: bold'));
    assert.ok(!out.includes('url') && !out.includes('position'));
});
test('форматування зберігається: шрифт, розмір, відступ, <font>, вирівнювання', () => {
    const out = w.sanitizeHtml('<p align="center" style="text-indent: 20px; margin: 0"><font color="blue" face="Arial" size="4">a</font> <span style="font-family: \'Times New Roman\', serif; font-size: 14pt; color: rgb(192, 0, 0); font-weight: 700; font-style: italic">b</span></p>').innerHTML;
    for (const part of ['align="center"', 'text-indent: 20px', '<font', 'color="blue"', 'face="Arial"', 'Times New Roman', 'font-size: 14pt', 'rgb(192, 0, 0)', 'font-weight: 700', 'font-style: italic']) {
        assert.ok(out.includes(part), 'втрачено: ' + part + '\n' + out);
    }
});
test('небезпечні стилі все одно відсікаються', () => {
    const out = w.sanitizeHtml('<p style="position: fixed; background: url(http://x); width: expression(alert(1)); color: red">a</p>').innerHTML;
    assert.ok(out.includes('color: red') && !/position|url|expression/.test(out), out);
});
test('Google Docs: списки, інтервали, таблиця, картинки зберігаються', () => {
    const html = '<b style="font-weight:normal;" id="docs-internal-guid-1"><h1 dir="ltr" style="line-height:1.38;margin-top:20pt;margin-bottom:6pt;"><span style="font-size:20pt;font-family:Arial,sans-serif;font-weight:400;white-space:pre-wrap;">Заголовок</span></h1>' +
        '<ul style="margin-top:0;margin-bottom:0;padding-inline-start:48px;"><li style="list-style-type:disc;font-size:11pt;" aria-level="1"><p style="line-height:1.38;text-align:justify;margin-top:0pt;margin-bottom:0pt;"><span style="font-size:11pt;">Пункт</span></p></li></ul>' +
        '<table style="border:none;border-collapse:collapse;table-layout:fixed;width:468pt"><colgroup><col width="100" /><col width="300" /></colgroup><tbody><tr><td style="border-left:solid #000000 1pt;vertical-align:top;padding:5pt 5pt 5pt 5pt;">A</td></tr></tbody></table>' +
        '<img src="https://lh7.googleusercontent.com/abc" width="200"><img src="data:image/png;base64,AAAA"></b>';
    const out = w.sanitizeHtml(html).innerHTML;
    for (const part of ['margin-top: 20pt', 'margin-bottom: 6pt', 'line-height: 1.38', 'padding-inline-start: 48px',
        'list-style-type: disc', 'text-align: justify', 'table-layout: fixed', 'border-left: solid #000000 1pt',
        'padding: 5pt 5pt 5pt 5pt', '<col width="100">', 'https://lh7.googleusercontent.com/abc', 'font-weight: normal']) {
        assert.ok(out.includes(part), 'втрачено: ' + part + '\n' + out);
    }
    assert.ok(!out.includes('data:image'), 'data: картинки мають відсікатись');
});
test('розміри з Docs масштабуються пропорційно, стовпці стають відсотками', () => {
    const el = w.sanitizeHtml('<h1 style="margin-top:20pt"><span style="font-size:20pt">A</span></h1><p style="margin:0pt 0pt 10pt 36pt"><span style="font-size:11pt">B</span></p><table><colgroup><col width="100"><col width="300"></colgroup><tbody><tr><td>x</td><td>y</td></tr></tbody></table><h2>Без розміру</h2>');
    w.scaleDocUnits(el);
    const html = el.innerHTML;
    assert.ok(html.includes('font-size: calc(20pt * var(--k-pt))'), html);
    assert.ok(html.includes('calc(36pt * var(--k-pt))'), html);
    assert.ok(/<col style="width: 25%;?">/.test(html) && /<col style="width: 75%;?">/.test(html), html);
    assert.ok(el.querySelector('h2').classList.contains('plain-h'));
    assert.ok(!el.querySelector('h1').classList.contains('plain-h'));
});
test('Word: одиниці cm/in переводяться в масштабовані pt', () => {
    const el = w.sanitizeHtml('<p style="margin: 0cm 0cm 8pt 1.27cm; text-indent: -0.5in">a</p>');
    w.scaleDocUnits(el);
    const h = el.innerHTML;
    assert.ok(h.includes('calc(0pt * var(--k-pt))') && h.includes('calc(36pt * var(--k-pt))') && h.includes('calc(-36pt * var(--k-pt))'), h);
});
test('Word: абзаци-пункти стають справжніми вкладеними списками', () => {
    const bullet = (lvl, mark, text, ml) => `<p class="MsoListParagraph" style="margin-left:${ml}pt;text-indent:-18pt"><span style="font-family:Symbol"><span>${mark}<span style="font:7pt 'Times New Roman'">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><b>${text}</b></p>`;
    const num = (n, text) => `<p class="MsoListParagraph" style="margin-left:36pt;text-indent:-18pt"><span>${n}<span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span>${text}</p>`;
    const html = '<p class="MsoNormal">Вступ</p>' + bullet(1, '·', 'Перший', 36) + bullet(2, 'o', 'Вкладений', 72) + bullet(1, '·', 'Другий', 36) +
        '<p class="MsoNormal">Між списками</p>' + num('1.', 'Раз') + num('2.', 'Два');
    const out = w.sanitizeHtml(html);
    assert.strictEqual(out.querySelectorAll('ul').length, 2, out.innerHTML);   // зовнішній + вкладений
    assert.strictEqual(out.querySelectorAll('ol').length, 1);
    assert.strictEqual(out.querySelectorAll('li').length, 5);
    assert.strictEqual(out.querySelector('ul > li > ul > li').textContent, 'Вкладений');
    assert.strictEqual(out.querySelector('ul').getAttribute('style'), 'list-style-type: disc');
    assert.strictEqual(out.querySelector('ul ul').getAttribute('style'), 'list-style-type: circle');
    assert.strictEqual(out.querySelector('ol').getAttribute('style'), 'list-style-type: decimal');
    assert.ok(!/·|Symbol|&nbsp;/.test(out.innerHTML.replace(/\u00a0/g, '&nbsp;')) && !out.textContent.includes('\u00a0\u00a0'), 'маркери-символи мають зникнути: ' + out.innerHTML);
    assert.ok(out.querySelector('li b'), 'жирний у пункті зберігається');
    assert.strictEqual(out.querySelector('p').textContent, 'Вступ');
});
test('Word: звичайний абзац із цифрою на початку не стає списком', () => {
    const out = w.sanitizeHtml('<p class="MsoNormal">1. Звичайний текст</p>');
    assert.strictEqual(out.querySelectorAll('ol, ul').length, 0);
});
test('панелі форматування (кнопки «Ж») у діалозі немає', () => {
    assert.strictEqual(w.document.getElementById('boldBtn'), null);
    assert.strictEqual(w.document.querySelectorAll('.editor-toolbar').length, 0);
});
test('CSS: блок із текстом не є областю прокрутки (інакше тачпад не гортає сторінку над текстом)', () => {
    const css = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
    const rules = [...css.matchAll(/(^|\})\s*([^{}]+)\{([^{}]*)\}/g)].filter(m => /^\s*#textcontent\s*$/.test(m[2]));
    assert.ok(rules.length > 0);
    for (const r of rules) assert.ok(!/overflow/.test(r[3]), 'У #textcontent не має бути overflow: ' + r[3]);
});
test('CSS: діалог вище за кнопку «+»', () => {
    const css = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
    const modalZ = +/\.modal\s*\{[^}]*z-index:\s*(\d+)/.exec(css)[1];
    const plusZ = +/\.add-tab-btn\s*\{[^}]*z-index:\s*(\d+)/.exec(css)[1];
    assert.ok(modalZ > plusZ, `modal ${modalZ} має бути більше за ${plusZ}`);
});
test('Word (.docx): форматування, списки, таблиця, посилання', async () => {
    const html = await w.BibleImport.docxToHtml(toArrayBuffer(fixture('sample.docx')));
    const d = w.sanitizeHtml(html); // те саме, що відбувається при збереженні
    const out = d.innerHTML;
    assert.ok(d.querySelector('h1') && d.querySelector('h2'), 'заголовки');
    assert.ok(/<span style="font-weight: 700">жирний<\/span>/.test(out) && /font-style: italic">курсив/.test(out), 'жирний і курсив');
    assert.ok(/text-decoration: underline">підкреслений/.test(out) && /line-through">закреслений/.test(out), 'підкреслення й закреслення');
    assert.ok(/color: #C00000/i.test(out) && /background-color: #ffff00/.test(out), 'колір і виділення');
    assert.ok(/font-family: 'Times New Roman',serif|font-family: 'Times New Roman', serif/.test(out) && out.includes('font-size: 14pt'), 'інший шрифт і розмір');
    assert.ok(d.querySelector('sup') && d.querySelector('sup').textContent === '2', 'верхній індекс');
    assert.ok(out.includes('text-align: center') && out.includes('text-align: justify') && out.includes('text-indent: 35.45pt'), 'вирівнювання та відступ');
    assert.ok(d.querySelectorAll('ul').length >= 2 && d.querySelectorAll('ol').length === 1 && d.querySelectorAll('li').length === 7, 'списки: ' + d.querySelectorAll('li').length);
    assert.ok(out.includes('list-style-type: disc') && out.includes('list-style-type: decimal'), 'тип маркерів');
    assert.ok(d.querySelector('table td[colspan="2"]') && d.querySelector('table td[rowspan="2"]'), 'об\'єднання клітинок');
    assert.ok(/background-color: #FFF2CC/i.test(out) && /border-top: 0.5pt solid #000000/.test(out), 'заливка та рамки клітинок');
    assert.ok(d.querySelector('a[href="https://example.com/"]'), 'гіперпосилання');
    assert.ok(d.textContent.includes('Мт 5:3-5'));
    // біблійні посилання знаходяться в імпортованому тексті
    w.linkifyReferences(w.scaleDocUnits(d));
    assert.ok(d.querySelectorAll('.bible-link').length >= 1, 'посилання на вірші підсвічуються');
});
test('Word (.docx): тип нумерації та контекстний інтервал як у Word', async () => {
    const html = await w.BibleImport.docxToHtml(toArrayBuffer(fixture('sample.docx')));
    const d = w.sanitizeHtml(html);
    const items = [...d.querySelectorAll('ul > li')].map(li => li.getAttribute('style'));
    assert.ok(/margin-bottom: 0pt/.test(items[0]), 'між однаковими пунктами немає інтервалу');
    assert.ok(/margin-bottom: 10pt/.test(items[1]), 'після останнього пункту інтервал є');
});
test('Word (.docx): помилки зрозумілі', async () => {
    await assert.rejects(() => w.BibleImport.docxToHtml(new Uint8Array([1, 2, 3, 4]).buffer), /не схоже на документ Word/);
});
test('TXT: UTF-8 та windows-1251, порожні рядки зберігаються', async () => {
    const utf = w.BibleImport.txtToHtml(w.BibleImport.decodeText(toArrayBuffer(fixture('sample.txt'))));
    assert.strictEqual((utf.match(/<p /g) || []).length, 4);
    assert.ok(utf.includes('<p style="margin: 0"><br></p>') && utf.includes('Мт 5:3'));
    assert.ok(w.BibleImport.decodeText(toArrayBuffer(fixture('cp1251.txt'))).startsWith('Привіт, світе!'));
});
test('Файли: .gdoc, .doc та невідомі формати пояснюються', async () => {
    const f = (name) => new w.File(['x'], name);
    await assert.rejects(() => w.BibleImport.fileToHtml(f('конспект.gdoc')), /ярлик Google Диска.*\.docx/s);
    await assert.rejects(() => w.BibleImport.fileToHtml(f('a.doc')), /\.doc.*\.docx/s);
    await assert.rejects(() => w.BibleImport.fileToHtml(f('a.xlsx')), /не підтримується/);
    const ok = await w.BibleImport.fileToHtml(new w.File([fixture('simple.docx')], 'Мій урок.docx'));
    assert.strictEqual(ok.title, 'Мій урок');
    assert.ok(ok.html.includes('Просто один абзац.'));
});
test('Діалог: кнопка «Завантажити файл» і приховане поле вибору файлу', () => {
    const btn = w.document.getElementById('importFileBtn');
    assert.ok(btn && w.document.getElementById('fileInput').hidden);
    assert.ok(/docx/.test(w.document.getElementById('fileInput').accept) && /pdf/.test(w.document.getElementById('fileInput').accept));
    assert.strictEqual(btn.parentElement.nextElementSibling, w.document.getElementById('inputArea'));
});
test('Смуга прогресу: прихована без сторінок, видима з відкритою сторінкою', async () => {
    const a = await makeApp();
    assert.strictEqual(a.document.getElementById('timer-progress').hidden, true);
    const b = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    assert.strictEqual(b.document.getElementById('timer-progress').hidden, false);
    assert.strictEqual(b.document.getElementById('pwa-timer').hidden, false);
});
test('Смуга прогресу ходить синхронно з таймером: старт, стоп, скидання', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    const fill = () => a.document.querySelector('#timer-progress .timer-progress-fill').style.width;
    const real = a.Date.now;
    let now = 1000000;
    a.Date.now = () => now;
    a.document.getElementById('timer-toggle').click();                       // старт
    assert.ok(a.document.getElementById('timer-toggle').classList.contains('stop'));
    now += 15 * 60000; a.updateTimerDisplay();                                // 15 хв із 30
    assert.strictEqual(fill(), '50%');
    assert.strictEqual(a.document.getElementById('timer-display').textContent, '15:00');
    a.document.getElementById('timer-toggle').click();                       // стоп
    const frozen = fill();
    now += 10 * 60000; a.updateTimerDisplay();
    assert.strictEqual(fill(), frozen, 'після зупинки смуга стоїть');
    a.document.getElementById('timer-toggle').click();                       // знову старт
    now += 20 * 60000; a.updateTimerDisplay();                                // 15+20 = 35 хв > 30
    assert.strictEqual(fill(), '100%');
    assert.ok(a.document.getElementById('timer-progress').classList.contains('overtime'));
    a.document.getElementById('resetTimerBtn').click();                      // скидання
    assert.strictEqual(fill(), '0%');
    assert.ok(!a.document.getElementById('timer-progress').classList.contains('overtime'));
    assert.strictEqual(a.document.getElementById('timer-toggle').className, 'timer-btn start');
    a.Date.now = real;
});
test('Час смуги: кнопка «Ліміт», панель, збереження, швидкі значення', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    const $ = (id) => a.document.getElementById(id);
    assert.strictEqual($('timer-limit-btn').textContent, 'Ліміт: 30 хв', 'за замовчуванням 30 хв');
    assert.strictEqual($('timer-limit-pop').hidden, true);
    $('timer-limit-btn').click();
    assert.strictEqual($('timer-limit-pop').hidden, false, 'кнопка відкриває панель');
    assert.strictEqual($('timer-limit-input').value, '30');
    let now = 5000000; a.Date.now = () => now;
    $('timer-toggle').click();
    $('timer-limit-input').value = '10';
    $('timer-limit-save').click();
    assert.strictEqual($('timer-limit-pop').hidden, true, 'після збереження панель закривається');
    assert.strictEqual($('timer-limit-btn').textContent, 'Ліміт: 10 хв');
    assert.strictEqual(a.localStorage.getItem('bible_timer_minutes'), '10');
    now += 5 * 60000; a.updateTimerDisplay();
    assert.strictEqual(a.document.querySelector('.timer-progress-fill').style.width, '50%', 'смуга одразу рахується від нового часу');
    $('timer-limit-btn').click(); $('timer-limit-input').value = '7,5'; $('timer-limit-save').click();
    assert.strictEqual($('timer-limit-btn').textContent, 'Ліміт: 7.5 хв');
    $('timer-limit-btn').click();
    a.document.querySelector('.timer-limit-preset[data-minutes="45"]').click();
    assert.strictEqual($('timer-limit-btn').textContent, 'Ліміт: 45 хв');
    assert.strictEqual($('timer-limit-pop').hidden, true, 'швидке значення зберігається одразу');
    $('timer-progress').click(); assert.strictEqual($('timer-limit-pop').hidden, false);
    $('timer-progress').click(); assert.strictEqual($('timer-limit-pop').hidden, true);
    $('timer-display').click(); assert.strictEqual($('timer-limit-pop').hidden, false);
    $('timer-limit-input').value = '99';
    $('timer-limit-input').dispatchEvent(new a.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    assert.strictEqual($('timer-limit-pop').hidden, true);
    assert.strictEqual(a.localStorage.getItem('bible_timer_minutes'), '45');
});
test('Час смуги: некоректні значення відхиляються', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    const $ = (id) => a.document.getElementById(id);
    const openPanel = () => { if ($('timer-limit-pop').hidden) $('timer-limit-btn').click(); };
    for (const bad of ['', 'abc', '0', '-5', '0.1', '1441', '99999']) {
        openPanel(); $('timer-limit-input').value = bad; $('timer-limit-save').click();
        assert.strictEqual($('timer-limit-pop').hidden, false, `«${bad}» не повинно прийматись`);
        assert.ok($('timer-limit-input').classList.contains('invalid'));
        assert.strictEqual(a.localStorage.getItem('bible_timer_minutes'), null);
    }
    for (const good of ['0.5', '1', '90', '1440']) {
        openPanel(); $('timer-limit-input').value = good; $('timer-limit-save').click();
        assert.strictEqual(a.localStorage.getItem('bible_timer_minutes'), good);
    }
});
test('Кнопка «Ліміт» і панель зникають разом із таймером', async () => {
    const a = await makeApp();
    assert.strictEqual(a.document.getElementById('timer-limit-btn').hidden, true);
    const b = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    assert.strictEqual(b.document.getElementById('timer-limit-btn').hidden, false);
    b.document.getElementById('timer-limit-btn').click();
    b.confirm = () => true; b.deletePage(0);
    assert.strictEqual(b.document.getElementById('timer-limit-pop').hidden, true, 'панель закрилась');
    assert.strictEqual(b.document.getElementById('timer-limit-btn').hidden, true);
});
test('Кнопка скидання коротко затемнюється й повертається', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>x</p>' }]) });
    const btn = a.document.getElementById('resetTimerBtn');
    btn.click();
    assert.ok(btn.classList.contains('pressed'), 'одразу після натискання');
    await new Promise(r => setTimeout(r, 400));
    assert.ok(!btn.classList.contains('pressed'), 'потім повертається до свого кольору');
});
test('CSS: стани кнопок таймера та симетрична смуга прогресу', () => {
    const css = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
    assert.ok(/#timer-toggle\.stop[^{]*\{[^}]*background-color/.test(css), 'кнопка STOP темна, поки йде відлік');
    assert.ok(/\.timer-btn\.pressed[^{]*\{[^}]*background-color/.test(css));
    const bar = /\.timer-progress\s*\{([^}]*)\}/.exec(css)[1];
    const left = /left:\s*(\d+)px/.exec(bar)[1], right = /right:\s*(\d+)px/.exec(bar)[1];
    assert.strictEqual(left, right, 'однакові відступи зліва й справа');
    assert.ok(/position:\s*fixed/.test(bar) && /top:\s*\d+px/.test(bar));
});
test('переноси рядків між блоками не дають порожніх рядків', () => {
    const el = w.sanitizeHtml('<p>a</p>\n<p>b</p>\n<ul>\n<li>c</li>\n</ul>\n');
    assert.strictEqual(el.innerHTML, '<p>a</p><p>b</p><ul><li>c</li></ul>');
});
test('назва з тегами не стає розміткою у вкладці', () => {
    w.pages = [];
    const page = { title: '<img src=x onerror=alert(1)>', content: '<p>a</p>' };
    w.eval('pages.push(' + JSON.stringify(page) + ')');
    w.renderTabs();
    assert.strictEqual(w.document.querySelectorAll('#side-tabs img').length, 0);
});

section('Вірші');
test('діапазон, список і розділ', () => {
    assert.strictEqual(w.getVerses('Від Матвія', '5', '3-5').verses.length, 3);
    same(w.getVerses('Від Матвія', '5', '3,5').verses.map(v => v.num), [3, 5]);
    const ch = w.getVerses('Від Матвія', '5', '');
    assert.strictEqual(ch.verses.length, 10);
    assert.strictEqual(ch.more, true);
});
test('усі книги зі словника існують у bibleText.json', () => {
    const books = new Set(Object.keys(bible).map(k => k.replace(/ \d+:\d+$/, '')));
    const used = new Set(Object.values(w.__bookLookup));
    const missing = [...used].filter(b => !books.has(b));
    same(missing, []);
});

section('Таймер');
test('прихований, поки немає сторінок', async () => {
    const a = await makeApp();
    assert.strictEqual(a.document.getElementById('pwa-timer').hidden, true);
});
test('з\'являється після збереження сторінки (Enter) і ховається після видалення останньої', async () => {
    const a = await makeApp();
    a.confirm = () => true;
    a.alert = (m) => { throw new Error(m); };
    const doc = a.document;
    a.openAddDialog();
    doc.getElementById('pageTitle').value = 'Тест';
    doc.getElementById('inputArea').innerHTML = '<p>Мт 5:3</p>';
    const ev = new a.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    doc.getElementById('inputArea').dispatchEvent(ev);
    assert.strictEqual(doc.getElementById('pwa-timer').hidden, false);
    assert.strictEqual(doc.querySelectorAll('.side-tab').length, 1);
    assert.ok(doc.querySelector('#textcontent .bible-link'));
    a.deletePage(0);
    assert.strictEqual(doc.getElementById('pwa-timer').hidden, true);
    assert.strictEqual(doc.querySelectorAll('.side-tab').length, 0);
});
test('Enter з порожнім текстом нічого не зберігає', async () => {
    const a = await makeApp();
    a.openAddDialog();
    a.document.getElementById('inputArea').dispatchEvent(new a.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    assert.strictEqual(a.document.querySelectorAll('.side-tab').length, 0);
    assert.strictEqual(a.document.getElementById('pwa-timer').hidden, true);
});
test('збережені сторінки відкриваються при запуску й показують таймер', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>Ів 3:16</p>' }]) });
    assert.strictEqual(a.document.getElementById('pwa-timer').hidden, false);
    assert.strictEqual(a.document.querySelectorAll('.bible-link').length, 1);
});
test('пошкоджений localStorage не ламає застосунок', async () => {
    const a = await makeApp({ bible_pages: '{зламано' });
    assert.strictEqual(a.document.querySelectorAll('.side-tab').length, 0);
    assert.strictEqual(a.localStorage.getItem('bible_pages_corrupt_backup'), '{зламано');
});
test('тап по посиланню показує підказку з текстом вірша (сенсорний режим)', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>Ів 3:16</p>' }]) }, true);
    a.setBibleData(bible);
    a.document.querySelector('.bible-link').dispatchEvent(new a.MouseEvent('click', { bubbles: true }));
    const tip = a.document.getElementById('bible-tooltip');
    assert.ok(tip && tip.textContent.length > 20);
});

test('наведення миші показує й ховає підказку (режим з мишею)', async () => {
    const a = await makeApp({ bible_pages: JSON.stringify([{ title: 'A', content: '<p>Ів 3:16</p>' }]) });
    a.setBibleData(bible);
    const link = a.document.querySelector('.bible-link');
    link.dispatchEvent(new a.MouseEvent('mouseover', { bubbles: true }));
    assert.ok(a.document.getElementById('bible-tooltip'));
    link.dispatchEvent(new a.MouseEvent('mouseout', { bubbles: true }));
    assert.strictEqual(a.document.getElementById('bible-tooltip'), null);
});
test('після багатьох перемикань вкладок слухачі не накопичуються', async () => {
    const pages = [1, 2, 3].map(i => ({ title: 'P' + i, content: '<p>Ів 3:16</p>' }));
    const a = await makeApp({ bible_pages: JSON.stringify(pages) });
    a.setBibleData(bible);
    for (let i = 0; i < 20; i++) a.loadPage(i % 3);
    let shown = 0;
    const origCreate = a.document.createElement.bind(a.document);
    a.document.createElement = (t) => { if (t === 'div') shown++; return origCreate(t); };
    a.document.querySelector('.bible-link').dispatchEvent(new a.MouseEvent('mouseover', { bubbles: true }));
    assert.strictEqual(shown, 1, 'підказка створена ' + shown + ' разів');
});

test('у розмітці немає кнопок експорту/імпорту', () => {
    assert.strictEqual(w.document.getElementById('export-btn'), null);
    assert.strictEqual(w.document.getElementById('import-btn'), null);
});

(async () => {
    w = await makeApp();
    w.setBibleData(bible);
    for (const item of queue) {
        if (item.title) { console.log(item.title); continue; }
        try { await item.fn(); passed++; console.log('  ok   ' + item.name); }
        catch (e) { failed++; console.log('  FAIL ' + item.name + '\n       ' + e.message.split('\n')[0] + '\n       ' + ((e.stack || '').split('\n').find(l => l.includes('app.test.js')) || '').trim()); }
    }
    console.log(`\nПройшло: ${passed}, не пройшло: ${failed}`);
    process.exit(failed ? 1 : 0);
})();
