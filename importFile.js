/* Імпорт файлів у діалог додавання тексту: .docx, .pdf, .txt → HTML.
   Усе працює локально, без інтернету (бібліотеки лежать у папці vendor/).
   Завантажується лише тоді, коли користувач вперше натискає «Завантажити файл». */
(function (global) {
    const MAX_FILE_BYTES = 40 * 1024 * 1024;

    /* ---------- Допоміжне ---------- */
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const num = (n) => String(+(+n).toFixed(2));

    const scriptPromises = {};
    function loadScript(src) {
        if (!scriptPromises[src]) {
            scriptPromises[src] = new Promise((resolve, reject) => {
                const s = document.createElement('script');
                s.src = src;
                s.onload = resolve;
                s.onerror = () => { delete scriptPromises[src]; reject(new Error('Не вдалося завантажити ' + src)); };
                document.head.appendChild(s);
            });
        }
        return scriptPromises[src];
    }

    function readBuffer(file) {
        if (file.arrayBuffer) return file.arrayBuffer();
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = () => resolve(r.result);
            r.onerror = () => reject(new Error('Не вдалося прочитати файл'));
            r.readAsArrayBuffer(file);
        });
    }

    /* =====================================================================
       TXT
       ===================================================================== */
    function decodeText(buffer) {
        const bytes = new Uint8Array(buffer);
        if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
        if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
        try {
            return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^\uFEFF/, '');
        } catch (e) {
            return new TextDecoder('windows-1251').decode(bytes); // старі українські/російські txt
        }
    }

    function txtToHtml(text) {
        const lines = text.replace(/\r\n?/g, '\n').split('\n');
        while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
        if (!lines.length) throw new Error('Файл порожній');
        return lines.map(l => l.trim() === '' ? '<p style="margin: 0"><br></p>' : `<p style="margin: 0">${esc(l)}</p>`).join('');
    }

    /* =====================================================================
       DOCX
       ===================================================================== */
    const kids = (el, name) => (el ? Array.from(el.children).filter(c => c.nodeName === name) : []);
    const kid = (el, name) => (el ? Array.from(el.children).find(c => c.nodeName === name) : null) || null;
    const wval = (el) => (el ? el.getAttribute('w:val') : null);
    const isOn = (el) => { const v = el.getAttribute('w:val'); return !(v === '0' || v === 'false' || v === 'off'); };
    const toInt = (v) => (v == null || v === '' ? undefined : parseInt(v, 10));

    function merge() {
        const r = {};
        for (const o of arguments) for (const k in o) if (o[k] !== undefined) r[k] = o[k];
        return r;
    }

    const HIGHLIGHT = {
        yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
        darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
        darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff'
    };

    function themeFont(rf, st) {
        const t = rf.getAttribute('w:asciiTheme') || rf.getAttribute('w:hAnsiTheme');
        if (!t || !st.theme) return undefined;
        return /major/i.test(t) ? st.theme.major : st.theme.minor;
    }

    function readRPr(rPr, st) {
        const o = {};
        if (!rPr) return o;
        let e;
        if ((e = kid(rPr, 'w:b'))) o.bold = isOn(e);
        if ((e = kid(rPr, 'w:i'))) o.italic = isOn(e);
        if ((e = kid(rPr, 'w:u'))) o.underline = wval(e) !== 'none';
        if ((e = kid(rPr, 'w:strike'))) o.strike = isOn(e);
        if ((e = kid(rPr, 'w:dstrike')) && isOn(e)) o.strike = true;
        if ((e = kid(rPr, 'w:caps'))) o.caps = isOn(e);
        if ((e = kid(rPr, 'w:smallCaps'))) o.smallCaps = isOn(e);
        if ((e = kid(rPr, 'w:vanish'))) o.vanish = isOn(e);
        if ((e = kid(rPr, 'w:color'))) { const v = wval(e); o.color = !v || v === 'auto' ? null : '#' + v; }
        if ((e = kid(rPr, 'w:highlight'))) { const v = wval(e); o.bg = v === 'none' ? null : (HIGHLIGHT[v] || undefined); }
        if ((e = kid(rPr, 'w:shd'))) { const f = e.getAttribute('w:fill'); if (f && f !== 'auto') o.bg = '#' + f; }
        if ((e = kid(rPr, 'w:sz'))) { const v = toInt(wval(e)); if (v) o.size = v / 2; }
        if ((e = kid(rPr, 'w:rFonts'))) {
            const f = e.getAttribute('w:ascii') || e.getAttribute('w:hAnsi') || themeFont(e, st);
            if (f) o.font = f;
        }
        if ((e = kid(rPr, 'w:vertAlign'))) o.vert = wval(e);
        return o;
    }

    function readPPr(pPr, st) {
        const o = {};
        if (!pPr) return o;
        let e;
        if ((e = kid(pPr, 'w:pStyle'))) o.styleId = wval(e);
        if ((e = kid(pPr, 'w:jc'))) o.jc = wval(e);
        if ((e = kid(pPr, 'w:ind'))) {
            const g = (n) => { const v = toInt(e.getAttribute('w:' + n)); return v === undefined ? undefined : v / 20; };
            const left = g('left') !== undefined ? g('left') : g('start');
            const right = g('right') !== undefined ? g('right') : g('end');
            if (left !== undefined) o.left = left;
            if (right !== undefined) o.right = right;
            if (g('hanging') !== undefined) o.indent = -g('hanging');
            else if (g('firstLine') !== undefined) o.indent = g('firstLine');
        }
        if ((e = kid(pPr, 'w:spacing'))) {
            const b = toInt(e.getAttribute('w:before')), a = toInt(e.getAttribute('w:after'));
            if (b !== undefined) o.before = b / 20;
            if (a !== undefined) o.after = a / 20;
            const line = toInt(e.getAttribute('w:line'));
            if (line !== undefined) { o.line = line; o.lineRule = e.getAttribute('w:lineRule') || 'auto'; }
        }
        if ((e = kid(pPr, 'w:numPr'))) {
            const id = kid(e, 'w:numId'), lvl = kid(e, 'w:ilvl');
            if (id) o.numId = wval(id);
            if (lvl) o.ilvl = toInt(wval(lvl));
        }
        if ((e = kid(pPr, 'w:contextualSpacing'))) o.contextual = isOn(e);
        if ((e = kid(pPr, 'w:shd'))) { const f = e.getAttribute('w:fill'); if (f && f !== 'auto') o.bg = '#' + f; }
        if ((e = kid(pPr, 'w:rPr'))) { const r = readRPr(e, st); if (r.size) o.markSize = r.size; }
        return o;
    }

    function readBorders(el) {
        const o = {};
        if (!el) return o;
        const map = { top: 'w:top', left: 'w:left', bottom: 'w:bottom', right: 'w:right', insideH: 'w:insideH', insideV: 'w:insideV' };
        const alt = { left: 'w:start', right: 'w:end' };
        for (const k in map) {
            const b = kid(el, map[k]) || (alt[k] ? kid(el, alt[k]) : null);
            if (b) o[k] = { val: b.getAttribute('w:val'), sz: toInt(b.getAttribute('w:sz')), color: b.getAttribute('w:color') };
        }
        return o;
    }

    function readCellMar(el) {
        const o = {};
        if (!el) return o;
        for (const [k, names] of [['top', ['w:top']], ['bottom', ['w:bottom']], ['left', ['w:left', 'w:start']], ['right', ['w:right', 'w:end']]]) {
            const b = kid(el, names[0]) || (names[1] ? kid(el, names[1]) : null);
            if (b) { const v = toInt(b.getAttribute('w:w')); if (v !== undefined) o[k] = v / 20; }
        }
        return o;
    }

    function parseTheme(xml) {
        if (!xml) return null;
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const latin = (tag) => {
            const f = doc.getElementsByTagName(tag)[0];
            const l = f && f.getElementsByTagName('a:latin')[0];
            return l ? l.getAttribute('typeface') : undefined;
        };
        return { major: latin('a:majorFont'), minor: latin('a:minorFont') };
    }

    function parseStyles(xml, st) {
        const result = { defaults: { rPr: {}, pPr: {} }, byId: {}, defPara: null };
        if (!xml) return result;
        const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
        const dd = kid(root, 'w:docDefaults');
        if (dd) {
            result.defaults.rPr = readRPr(kid(kid(dd, 'w:rPrDefault'), 'w:rPr'), st);
            result.defaults.pPr = readPPr(kid(kid(dd, 'w:pPrDefault'), 'w:pPr'), st);
        }
        for (const s of kids(root, 'w:style')) {
            const id = s.getAttribute('w:styleId');
            const type = s.getAttribute('w:type');
            const tblPr = kid(s, 'w:tblPr');
            result.byId[id] = {
                id, type,
                name: wval(kid(s, 'w:name')) || '',
                basedOn: wval(kid(s, 'w:basedOn')),
                rPr: readRPr(kid(s, 'w:rPr'), st),
                pPr: readPPr(kid(s, 'w:pPr'), st),
                borders: readBorders(kid(tblPr, 'w:tblBorders')),
                cellMar: readCellMar(kid(tblPr, 'w:tblCellMar'))
            };
            if (s.getAttribute('w:default') === '1' && type === 'paragraph') result.defPara = id;
        }
        return result;
    }

    function styleChain(styles, id) {
        const chain = [];
        const seen = new Set();
        while (id && styles.byId[id] && !seen.has(id)) {
            seen.add(id);
            chain.unshift(styles.byId[id]);
            id = styles.byId[id].basedOn;
        }
        return chain;
    }

    function parseNumbering(xml, st) {
        const result = { abs: {}, nums: {} };
        if (!xml) return result;
        const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
        for (const an of kids(root, 'w:abstractNum')) {
            const lvls = {};
            for (const lvl of kids(an, 'w:lvl')) {
                const ppr = readPPr(kid(lvl, 'w:pPr'), st);
                lvls[toInt(lvl.getAttribute('w:ilvl'))] = {
                    fmt: wval(kid(lvl, 'w:numFmt')) || 'decimal',
                    text: wval(kid(lvl, 'w:lvlText')) || '',
                    start: toInt(wval(kid(lvl, 'w:start'))) || 1,
                    left: ppr.left, indent: ppr.indent
                };
            }
            result.abs[an.getAttribute('w:abstractNumId')] = lvls;
        }
        for (const n of kids(root, 'w:num')) {
            const overrides = {};
            for (const ov of kids(n, 'w:lvlOverride')) {
                const so = kid(ov, 'w:startOverride');
                if (so) overrides[toInt(ov.getAttribute('w:ilvl'))] = toInt(wval(so));
            }
            result.nums[n.getAttribute('w:numId')] = { abs: wval(kid(n, 'w:abstractNumId')), overrides };
        }
        return result;
    }

    function parseRels(xml) {
        const rels = {};
        if (!xml) return rels;
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        for (const r of Array.from(doc.getElementsByTagName('Relationship'))) {
            if (r.getAttribute('TargetMode') === 'External') rels[r.getAttribute('Id')] = r.getAttribute('Target');
        }
        return rels;
    }

    const NUM_FMT = {
        decimal: 'decimal', decimalZero: 'decimal-leading-zero', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha',
        lowerRoman: 'lower-roman', upperRoman: 'upper-roman'
    };

    function fontStack(name) {
        const clean = name.replace(/[^\p{L}\p{N}\s-]/gu, '').trim();
        if (!clean) return 'sans-serif';
        let generic = 'sans-serif';
        if (/courier|consolas|mono|lucida console/i.test(clean)) generic = 'monospace';
        else if (/times|georgia|cambria|garamond|palatino|book|serif|minion|constantia|sylfaen|century/i.test(clean) && !/sans/i.test(clean)) generic = 'serif';
        return `'${clean}',${generic}`;
    }

    // Спільні для абзацу властивості, які успадковують усі фрагменти тексту
    function inheritedCss(base) {
        const css = [];
        if (base.font) css.push('font-family:' + fontStack(base.font));
        css.push(`font-size:${num(base.size || 10)}pt`);
        if (base.bold) css.push('font-weight:700');
        if (base.italic) css.push('font-style:italic');
        if (base.color) css.push('color:' + base.color);
        return css;
    }

    // Для фрагмента записуємо лише відмінності від абзацу
    function runCss(eff, base) {
        const css = [];
        if (eff.font && eff.font !== base.font) css.push('font-family:' + fontStack(eff.font));
        if (eff.size && eff.size !== base.size) css.push(`font-size:${num(eff.size)}pt`);
        if (!!eff.bold !== !!base.bold) css.push(eff.bold ? 'font-weight:700' : 'font-weight:400');
        if (!!eff.italic !== !!base.italic) css.push(eff.italic ? 'font-style:italic' : 'font-style:normal');
        const deco = [];
        if (eff.underline) deco.push('underline');
        if (eff.strike) deco.push('line-through');
        if (deco.length) css.push('text-decoration:' + deco.join(' '));
        if (eff.color && eff.color !== base.color) css.push('color:' + eff.color);
        if (eff.bg) css.push('background-color:' + eff.bg);
        if (eff.caps) css.push('text-transform:uppercase');
        if (eff.smallCaps) css.push('font-variant:small-caps');
        return css.join(';');
    }

    function runHtml(r, base, st) {
        const rPr = kid(r, 'w:rPr');
        const rs = wval(kid(rPr, 'w:rStyle'));
        const charChain = rs ? styleChain(st.styles, rs) : [];
        const eff = merge(base, ...charChain.map(s => s.rPr), readRPr(rPr, st));
        if (eff.vanish) return '';
        let content = '';
        for (const c of Array.from(r.children)) {
            switch (c.nodeName) {
                case 'w:t': content += esc(c.textContent); break;
                case 'w:tab': content += '\t'; break;
                case 'w:br': { const t = c.getAttribute('w:type'); if (t !== 'page' && t !== 'column') content += '<br>'; break; }
                case 'w:cr': content += '<br>'; break;
                case 'w:noBreakHyphen': content += '\u2011'; break;
                default: break;
            }
        }
        if (!content) return '';
        if (eff.vert === 'superscript') content = `<sup>${content}</sup>`;
        else if (eff.vert === 'subscript') content = `<sub>${content}</sub>`;
        const css = runCss(eff, base);
        return css ? `<span style="${css}">${content}</span>` : content;
    }

    function inlineHtml(parent, base, st) {
        let html = '';
        for (const el of Array.from(parent.children)) {
            switch (el.nodeName) {
                case 'w:r': html += runHtml(el, base, st); break;
                case 'w:hyperlink': {
                    const rid = el.getAttribute('r:id');
                    const href = rid ? st.rels[rid] : null;
                    const inner = inlineHtml(el, base, st);
                    html += href && /^(https?:|mailto:)/i.test(href) ? `<a href="${esc(href)}">${inner}</a>` : inner;
                    break;
                }
                case 'w:ins': case 'w:smartTag': case 'w:fldSimple': case 'w:customXml': case 'w:moveTo':
                    html += inlineHtml(el, base, st); break;
                case 'w:sdt': { const c = kid(el, 'w:sdtContent'); if (c) html += inlineHtml(c, base, st); break; }
                default: break; // w:pPr, w:del, закладки, правопис тощо
            }
        }
        return html;
    }

    // Властивості абзацу: усе, що не задано, береться зі стилю, таблиці та значень за замовчуванням
    function paraInfo(p, st) {
        const pPrEl = kid(p, 'w:pPr');
        const direct = readPPr(pPrEl, st);
        const chain = styleChain(st.styles, direct.styleId || st.styles.defPara);
        const eff = merge(st.styles.defaults.pPr, ...st.tableChain.map(s => s.pPr), ...chain.map(s => s.pPr), direct);
        const base = merge(st.styles.defaults.rPr, ...st.tableChain.map(s => s.rPr), ...chain.map(s => s.rPr));
        const name = chain.length ? chain[chain.length - 1].name.toLowerCase() : '';
        const hm = /^heading\s*([1-6])$/.exec(name);
        return { eff, base, direct, styleKey: direct.styleId || st.styles.defPara, tag: hm ? 'h' + hm[1] : 'p', inner: inlineHtml(p, base, st) };
    }

    function paraCss(info, isListItem) {
        const { eff, base } = info;
        const css = inheritedCss(info.inner ? base : merge(base, { size: eff.markSize || base.size }));
        if (!isListItem) {
            if (eff.left) css.push(`margin-left:${num(eff.left)}pt`);
            if (eff.right) css.push(`margin-right:${num(eff.right)}pt`);
            if (eff.indent) css.push(`text-indent:${num(eff.indent)}pt`);
        }
        css.push(`margin-top:${num(info.noTop ? 0 : (eff.before || 0))}pt`);
        css.push(`margin-bottom:${num(info.noBottom ? 0 : (eff.after || 0))}pt`);
        if (eff.line) {
            if (eff.lineRule === 'exact' || eff.lineRule === 'atLeast') css.push(`line-height:${num(eff.line / 20)}pt`);
            else css.push(`line-height:${num(eff.line / 240 * 1.2)}`); // «одинарний» інтервал Word ≈ 1.2 висоти шрифту
        } else css.push('line-height:1.2');
        if (eff.jc === 'center') css.push('text-align:center');
        else if (eff.jc === 'right' || eff.jc === 'end') css.push('text-align:right');
        else if (eff.jc === 'both' || eff.jc === 'distribute') css.push('text-align:justify');
        if (eff.bg) css.push(`background-color:${eff.bg}`);
        return css.join(';');
    }

    function listInfo(info, st) {
        const { eff } = info;
        if (!eff.numId || eff.numId === '0') return null;
        const numDef = st.numbering.nums[eff.numId];
        if (!numDef) return null;
        const ilvl = eff.ilvl || 0;
        const lv = (st.numbering.abs[numDef.abs] || {})[ilvl] || { fmt: 'bullet', text: '•', start: 1 };
        const ordered = lv.fmt !== 'bullet' && lv.fmt !== 'none';
        let type;
        if (ordered) type = NUM_FMT[lv.fmt] || 'decimal';
        else if (lv.fmt === 'none') type = 'none';
        else if (/^[o○◦]$/.test(lv.text)) type = 'circle';
        else if (/^[▪■□§\uF0A7]$/.test(lv.text)) type = 'square';
        else type = 'disc';
        let start = numDef.overrides[ilvl] !== undefined ? numDef.overrides[ilvl] : lv.start;
        let left = info.direct.left !== undefined ? info.direct.left : (lv.left !== undefined ? lv.left : eff.left);
        if (left === undefined || left === 0) left = 18 * (ilvl + 1) + 18;
        return { numId: eff.numId, ilvl, tag: ordered ? 'ol' : 'ul', type, start, left, ordered };
    }

    function borderCss(b) {
        if (!b || b.val === 'nil' || b.val === 'none') return 'none';
        const w = Math.max((b.sz || 4) / 8, 0.5);
        const c = !b.color || b.color === 'auto' ? '#000000' : '#' + b.color;
        return `${num(w)}pt solid ${c}`;
    }

    function tableHtml(tbl, st) {
        const tblPr = kid(tbl, 'w:tblPr');
        const chain = styleChain(st.styles, wval(kid(tblPr, 'w:tblStyle')));
        const borders = merge(...chain.map(s => s.borders), readBorders(kid(tblPr, 'w:tblBorders')));
        const mar = merge({ left: 5.4, right: 5.4, top: 0, bottom: 0 }, ...chain.map(s => s.cellMar), readCellMar(kid(tblPr, 'w:tblCellMar')));
        const grid = kids(kid(tbl, 'w:tblGrid'), 'w:gridCol').map(g => (toInt(g.getAttribute('w:w')) || 0) / 20);
        const rows = kids(tbl, 'w:tr');

        // 1) координати клітинок та об'єднання по вертикалі
        const owners = [];
        const matrix = rows.map((tr, ri) => {
            let col = 0;
            return kids(tr, 'w:tc').map((tc) => {
                const tcPr = kid(tc, 'w:tcPr');
                const span = toInt(wval(kid(tcPr, 'w:gridSpan'))) || 1;
                const vm = kid(tcPr, 'w:vMerge');
                const cell = { tc, tcPr, col, span, ri, rowspan: 1, skip: false };
                if (vm && (vm.getAttribute('w:val') || 'continue') === 'continue') {
                    if (owners[col]) { owners[col].rowspan++; cell.skip = true; }
                } else owners[col] = cell;
                col += span;
                return cell;
            });
        });
        const ncols = grid.length || Math.max(1, ...matrix.map(r => r.reduce((a, c) => a + c.span, 0)));

        // 2) розмітка
        const totalW = grid.reduce((a, b) => a + b, 0);
        const jc = wval(kid(tblPr, 'w:jc'));
        let html = `<table style="border-collapse:collapse;table-layout:fixed;${totalW ? 'width:' + num(totalW) + 'pt;' : 'width:100%;'}` +
            `${jc === 'center' ? 'margin-left:auto;margin-right:auto;' : ''}">`;
        if (grid.length) html += '<colgroup>' + grid.map(w => `<col width="${Math.max(1, Math.round(w / 0.75))}">`).join('') + '</colgroup>';
        html += '<tbody>';
        matrix.forEach((cells, ri) => {
            html += '<tr>';
            for (const c of cells) {
                if (c.skip) continue;
                const tcb = readBorders(kid(c.tcPr, 'w:tcBorders'));
                const lastRow = ri + c.rowspan - 1 >= rows.length - 1;
                const lastCol = c.col + c.span >= ncols;
                const edge = {
                    top: ri === 0 ? borders.top : borders.insideH,
                    bottom: lastRow ? borders.bottom : borders.insideH,
                    left: c.col === 0 ? borders.left : borders.insideV,
                    right: lastCol ? borders.right : borders.insideV
                };
                const b = merge(edge, tcb);
                const cm = merge(mar, readCellMar(kid(c.tcPr, 'w:tcMar')));
                const css = [
                    `border-top:${borderCss(b.top)}`, `border-right:${borderCss(b.right)}`,
                    `border-bottom:${borderCss(b.bottom)}`, `border-left:${borderCss(b.left)}`,
                    `padding:${num(cm.top)}pt ${num(cm.right)}pt ${num(cm.bottom)}pt ${num(cm.left)}pt`
                ];
                const shd = kid(c.tcPr, 'w:shd');
                const fill = shd && shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') css.push(`background-color:#${fill}`);
                const va = wval(kid(c.tcPr, 'w:vAlign'));
                css.push(`vertical-align:${va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top'}`);
                const saved = st.tableChain;
                st.tableChain = chain;
                const inner = convertBlocks(c.tc, st);
                st.tableChain = saved;
                html += `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''} style="${css.join(';')}">${inner || '<p style="margin:0"><br></p>'}</td>`;
            }
            html += '</tr>';
        });
        return html + '</tbody></table>';
    }

    // Перетворює вміст контейнера (тіло документа або клітинка таблиці) на HTML зі справжніми списками
    function convertBlocks(container, st) {
        let out = '';
        const stack = [];
        const closeLists = () => { while (stack.length) out += `</li></${stack.pop().tag}>`; };

        function listItem(info, li) {
            while (stack.length && stack[stack.length - 1].level > li.ilvl) out += `</li></${stack.pop().tag}>`;
            let top = stack[stack.length - 1];
            if (top && top.level === li.ilvl && (top.numId !== li.numId || top.tag !== li.tag)) {
                out += `</li></${stack.pop().tag}>`;
                top = stack[stack.length - 1];
            }
            // лічильник нумерації (продовжується, навіть якщо список перервав звичайний абзац)
            const key = li.numId + ':' + li.ilvl;
            for (const k of Object.keys(st.counters)) {
                const [id, lv] = k.split(':');
                if (id === li.numId && +lv > li.ilvl) delete st.counters[k];
            }
            st.counters[key] = st.counters[key] === undefined ? li.start : st.counters[key] + 1;
            if (top && top.level === li.ilvl) out += '</li>';
            else {
                const parentLeft = top ? top.left : 0;
                const pad = Math.max(14, li.left - parentLeft);
                out += `<${li.tag} style="margin:0;padding-left:${num(pad)}pt;list-style-type:${li.type}"` +
                    `${li.ordered && st.counters[key] !== 1 ? ` start="${st.counters[key]}"` : ''}>`;
                stack.push({ level: li.ilvl, tag: li.tag, numId: li.numId, left: li.left });
            }
            out += `<li style="${paraCss(info, true)}">${info.inner || '<br>'}`;
        }

        function walk(parent) {
            const els = Array.from(parent.children);
            const infos = new Map();
            els.forEach((el) => { if (el.nodeName === 'w:p') infos.set(el, paraInfo(el, st)); });
            const sameStyle = (a, b) => a && b && a.eff.contextual && b.eff.contextual && a.styleKey === b.styleKey;
            els.forEach((el, i) => {
                switch (el.nodeName) {
                    case 'w:p': {
                        const info = infos.get(el);
                        info.noTop = sameStyle(info, infos.get(els[i - 1]));
                        info.noBottom = sameStyle(info, infos.get(els[i + 1]));
                        const li = listInfo(info, st);
                        if (li) listItem(info, li);
                        else {
                            closeLists();
                            out += `<${info.tag} style="${paraCss(info, false)}">${info.inner || '<br>'}</${info.tag}>`;
                        }
                        break;
                    }
                    case 'w:tbl': closeLists(); out += tableHtml(el, st); break;
                    case 'w:sdt': { const c = kid(el, 'w:sdtContent'); if (c) walk(c); break; }
                    case 'w:customXml': case 'w:ins': case 'w:moveTo': walk(el); break;
                    default: break; // w:sectPr тощо
                }
            });
        }
        walk(container);
        closeLists();
        return out;
    }

    async function docxToHtml(buffer) {
        if (!global.fflate) await loadScript('vendor/fflate.min.js');
        const wanted = new Set(['word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels', 'word/theme/theme1.xml']);
        let files;
        try {
            files = global.fflate.unzipSync(new Uint8Array(buffer), { filter: (f) => wanted.has(f.name) }); // картинки не розпаковуємо
        } catch (e) {
            throw new Error('Не вдалося відкрити файл: це не схоже на документ Word (.docx).');
        }
        const read = (n) => (files[n] ? global.fflate.strFromU8(files[n]) : null);
        const docXml = read('word/document.xml');
        if (!docXml) throw new Error('Це не схоже на документ Word (.docx): немає word/document.xml.');

        const st = { theme: parseTheme(read('word/theme/theme1.xml')), tableChain: [], counters: {} };
        st.styles = parseStyles(read('word/styles.xml'), st);
        st.numbering = parseNumbering(read('word/numbering.xml'), st);
        st.rels = parseRels(read('word/_rels/document.xml.rels'));

        const doc = new DOMParser().parseFromString(docXml, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) throw new Error('Документ пошкоджений і не читається.');
        const body = doc.getElementsByTagName('w:body')[0];
        if (!body) throw new Error('У документі немає тексту.');
        const html = convertBlocks(body, st);
        if (!html.replace(/<[^>]*>/g, '').trim()) throw new Error('У документі немає тексту.');
        return html;
    }

    /* =====================================================================
       PDF
       ===================================================================== */
    const BULLET_START = /^([•·▪■○◦\-–—*]|\d{1,3}[.)]|[a-zA-Zа-яіїєґА-ЯІЇЄҐ][.)])\s/;

    function modeOf(values) {
        const counts = new Map();
        let best = values[0], bestN = 0;
        for (const v of values) {
            const n = (counts.get(v) || 0) + 1;
            counts.set(v, n);
            if (n > bestN) { best = v; bestN = n; }
        }
        return best;
    }

    // Збирає рядки сторінки з фрагментів тексту, які віддає pdf.js
    function pageLines(items, fontInfo) {
        const pieces = items
            .filter(it => typeof it.str === 'string' && it.str.length)
            .map(it => {
                const t = it.transform;
                const fi = fontInfo(it.fontName);
                return { str: it.str, x: t[4], y: t[5], size: Math.hypot(t[0], t[1]) || 10, w: it.width || 0, bold: !!fi.bold, italic: !!fi.italic };
            });
        pieces.sort((a, b) => b.y - a.y || a.x - b.x);
        const lines = [];
        for (const p of pieces) {
            const line = lines.find(l => Math.abs(l.y - p.y) <= 0.45 * Math.min(l.size, p.size));
            if (line) line.items.push(p);
            else lines.push({ y: p.y, size: p.size, items: [p] });
        }
        const result = [];
        for (const l of lines) {
            l.items.sort((a, b) => a.x - b.x);
            const segs = [];
            let prev = null;
            for (const it of l.items) {
                let text = it.str;
                if (prev && it.x - (prev.x + prev.w) > 0.15 * it.size && !/\s$/.test(prev.str) && !/^\s/.test(text)) text = ' ' + text;
                const last = segs[segs.length - 1];
                if (last && last.bold === it.bold && last.italic === it.italic && Math.abs(last.size - it.size) < 0.6) last.text += text;
                else segs.push({ text, size: it.size, bold: it.bold, italic: it.italic });
                prev = it;
            }
            const text = segs.map(s => s.text).join('');
            if (!text.trim()) continue;
            const first = l.items[0], lastIt = l.items[l.items.length - 1];
            result.push({
                y: l.y, size: Math.max(...l.items.map(i => i.size)), x0: first.x, x1: lastIt.x + lastIt.w,
                segs, text, allBold: segs.every(s => s.bold || !s.text.trim())
            });
        }
        result.sort((a, b) => b.y - a.y);
        return result;
    }

    function pageParagraphs(lines, vp) {
        // номери сторінок у верхньому/нижньому полі
        lines = lines.filter(l => !(/^\d{1,4}$/.test(l.text.trim()) && (l.y < 0.07 * vp.height || l.y > 0.93 * vp.height)));
        if (!lines.length) return [];
        const baseLeft = modeOf(lines.map(l => Math.round(l.x0 / 2) * 2));
        const sorted = lines.map(l => l.x1).sort((a, b) => a - b);
        const maxRight = sorted[Math.floor(sorted.length * 0.9)];
        // звичайний міжрядковий проміжок сторінки (медіана) — більший проміжок означає новий абзац
        const gaps = [];
        for (let i = 1; i < lines.length; i++) {
            const g = lines[i - 1].y - lines[i].y;
            if (g > 0 && g < 3 * Math.max(lines[i].size, lines[i - 1].size)) gaps.push(g);
        }
        gaps.sort((a, b) => a - b);
        const lead = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 0;
        const paras = [];
        let cur = null, prev = null;
        for (const l of lines) {
            let startNew = !cur;
            if (cur) {
                const gap = prev.y - l.y;
                const normal = lead || 1.2 * Math.max(prev.size, l.size);
                const shortEnding = prev.x1 < baseLeft + 0.4 * (maxRight - baseLeft) && /[.!?…:;»”")]$/.test(prev.text.trim());
                startNew =
                    gap > 1.3 * normal ||
                    BULLET_START.test(l.text.trim()) ||
                    (l.x0 - baseLeft > 0.9 * l.size && prev.x0 - baseLeft < 0.9 * prev.size) ||
                    shortEnding ||
                    Math.abs(l.size - prev.size) > 1.2 ||
                    prev.allBold !== l.allBold;
            }
            if (startNew) { cur = { lines: [], baseLeft }; paras.push(cur); }
            cur.lines.push(l);
            prev = l;
        }
        paras.forEach(p => { p.vpWidth = vp.width; });
        return paras;
    }

    function paraHtml(p, bodySize) {
        // розмір абзацу = найпоширеніший за кількістю символів
        const weights = new Map();
        for (const l of p.lines) for (const s of l.segs) {
            const k = Math.round(s.size * 2) / 2;
            weights.set(k, (weights.get(k) || 0) + s.text.length);
        }
        const size = [...weights.entries()].sort((a, b) => b[1] - a[1])[0][0];
        const x0Min = Math.min(...p.lines.map(l => l.x0));
        const centered = p.lines.length <= 3 && p.lines.every(l => Math.abs((l.x0 + l.x1) / 2 - p.vpWidth / 2) < 0.04 * p.vpWidth) &&
            x0Min - p.baseLeft > 0.08 * p.vpWidth;
        let inner = '';
        let lastStyle = null;
        let buf = '';
        const flush = () => {
            if (!buf) return;
            const css = [];
            if (lastStyle.bold) css.push('font-weight:700');
            if (lastStyle.italic) css.push('font-style:italic');
            if (Math.abs(lastStyle.size - size) > 0.8) css.push(`font-size:${num(lastStyle.size)}pt`);
            inner += css.length ? `<span style="${css.join(';')}">${esc(buf)}</span>` : esc(buf);
            buf = '';
        };
        p.lines.forEach((l, li) => {
            l.segs.forEach((s, si) => {
                let text = s.text;
                if (li > 0 && si === 0) {
                    // перенос слова в кінці рядка склеюємо, інакше ставимо пробіл
                    if (/[A-Za-zА-Яа-яІіЇїЄєҐґ]-$/.test(buf) && /^[a-zа-яіїєґ]/.test(text)) buf = buf.slice(0, -1);
                    else if (!/\s$/.test(buf)) text = ' ' + text;
                }
                const key = { bold: s.bold, italic: s.italic, size: Math.round(s.size * 2) / 2 };
                if (lastStyle && (lastStyle.bold !== key.bold || lastStyle.italic !== key.italic || lastStyle.size !== key.size)) flush();
                lastStyle = key;
                buf += text;
            });
        });
        flush();
        const css = [`font-size:${num(size)}pt`, 'line-height:1.2', `margin-top:0`, `margin-bottom:${num(size * (size > bodySize * 1.15 ? 0.8 : 0.55))}pt`];
        if (centered) css.push('text-align:center');
        else {
            const offset = x0Min - p.baseLeft;
            if (offset > 1.5 * size) css.push(`margin-left:${num(offset)}pt`);
            const firstIndent = p.lines[0].x0 - x0Min;
            if (firstIndent > 0.9 * size) css.push(`text-indent:${num(firstIndent)}pt`);
        }
        return `<p style="${css.join(';')}">${inner}</p>`;
    }

    async function pdfToHtml(buffer, onProgress) {
        if (!global.pdfjsLib) await loadScript('vendor/pdf.min.js');
        const lib = global.pdfjsLib;
        lib.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.js', document.baseURI).href;
        let pdf;
        try {
            pdf = await lib.getDocument({ data: new Uint8Array(buffer), isEvalSupported: false }).promise;
        } catch (e) {
            throw new Error(e && e.name === 'PasswordException' ? 'PDF захищений паролем.' : 'Не вдалося відкрити PDF: файл пошкоджений або це не PDF.');
        }
        const allParas = [];
        const sizes = [];
        for (let n = 1; n <= pdf.numPages; n++) {
            if (onProgress) onProgress(n, pdf.numPages);
            const page = await pdf.getPage(n);
            const vp = page.getViewport({ scale: 1 });
            const tc = await page.getTextContent();
            try { await page.getOperatorList(); } catch (e) { /* шрифти не обов'язкові */ }
            const cache = {};
            const fontInfo = (name) => {
                if (cache[name]) return cache[name];
                let r = {};
                try {
                    const f = page.commonObjs.get(name);
                    const fname = (f && (f.name || f.loadedName)) || '';
                    r = { bold: !!(f && f.bold) || /bold|black|heavy|semibold|demi/i.test(fname), italic: !!(f && f.italic) || /italic|oblique/i.test(fname) };
                } catch (e) { /* ще не завантажений */ }
                return (cache[name] = r);
            };
            const lines = pageLines(tc.items, fontInfo);
            lines.forEach(l => l.segs.forEach(s => sizes.push(Math.round(s.size * 2) / 2)));
            const paras = pageParagraphs(lines, vp);
            if (paras.length) paras[paras.length - 1].pageEnd = true; // може продовжитись на наступній сторінці
            for (const p of paras) allParas.push(p);
            page.cleanup();
        }
        if (!allParas.length) throw new Error('У PDF не знайдено тексту. Якщо це скан (картинка), програма не вміє його розпізнавати.');
        const bodySize = modeOf(sizes);
        // абзац, що обірвався на межі сторінок, продовжуємо
        const merged = [];
        for (const p of allParas) {
            const last = merged[merged.length - 1];
            const lastText = last ? last.lines[last.lines.length - 1].text.trim() : '';
            const firstText = p.lines[0].text.trim();
            if (last && last.pageEnd && !/[.!?…:;»”"')\]]$/.test(lastText) && /^[a-zа-яіїєґ]/.test(firstText) &&
                Math.abs(last.lines[0].size - p.lines[0].size) < 1) {
                last.lines = last.lines.concat(p.lines);
                last.pageEnd = !!p.pageEnd;
            } else merged.push(p);
        }
        return merged.map(p => paraHtml(p, bodySize)).join('');
    }

    /* =====================================================================
       Вхідна точка
       ===================================================================== */
    async function fileToHtml(file, onProgress) {
        const name = file.name || '';
        const ext = ((name.match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();
        const title = name.replace(/\.[^.]+$/, '');
        if (ext === 'gdoc' || ext === 'gsheet' || ext === 'gslides') {
            throw new Error('Файл .gdoc — це лише ярлик Google Диска без самого тексту, його неможливо прочитати. ' +
                'Відкрийте документ у Google Docs: Файл → Завантажити → Microsoft Word (.docx) — і завантажте отриманий файл.');
        }
        if (ext === 'doc') {
            throw new Error('Старий формат Word (.doc) не підтримується. Відкрийте файл у Word або Google Docs і збережіть як .docx.');
        }
        if (!['docx', 'txt', 'pdf'].includes(ext)) {
            throw new Error('Цей формат не підтримується. Можна завантажувати: Word (.docx), PDF (.pdf), текст (.txt).');
        }
        if (file.size > MAX_FILE_BYTES) throw new Error('Файл завеликий (понад 40 МБ).');
        const buffer = await readBuffer(file);
        let html;
        if (ext === 'docx') html = await docxToHtml(buffer);
        else if (ext === 'txt') html = txtToHtml(decodeText(buffer));
        else html = await pdfToHtml(buffer, onProgress);
        return { html, title, ext };
    }

    global.BibleImport = { fileToHtml, docxToHtml, txtToHtml, pdfToHtml, decodeText };
})(typeof window !== 'undefined' ? window : globalThis);
