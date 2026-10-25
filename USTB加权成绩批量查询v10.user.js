// ==UserScript==
// @name         USTB-BYYT-加权成绩·课程排名·批量查询
// @namespace    https://github.com/wishes1129/ustb-grade-enhancer
// @version      10.1.0
// @description  适配2026年改版本研一体化教务系统：实时加权计算、单课课程排名、单人/Excel批量查询、成绩页表格联动
// @author       wishes1129 (汪其年)
// @match        https://byyt.ustb.edu.cn/*
// @grant        unsafeWindow
// @require      https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
// @run-at       document-idle
// @license      MIT
// ==/UserScript==

(function () {
    'use strict';
    const log = (...a) => console.log('[USTB-v10]', ...a);
    if (window.top !== window.self) return; // 只在顶层窗口注入

    // ==================== 接口层 ====================
    const HEADERS = { 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest', 'RoleCode': '01' };
    async function post(path, body) {
        const res = await fetch(path, { method: 'POST', credentials: 'include', headers: HEADERS, body: JSON.stringify(body) });
        const j = await res.json().catch(() => null);
        if (!j) throw new Error('接口无响应: ' + path);
        return j;
    }
    // 拿学籍+方案号
    async function getXs(xh) {
        const j = await post('/cjgl/cjzhtjcx/cjcx/getXs', { xjidorxh: xh, pylx: '1' });
        const s = j && j.content && j.content[0];
        if (!s || !s.fah) throw new Error(xh + ' 学籍/方案未找到');
        return s;
    }
    // 拉取全部成绩+单课排名（grcjcxByXjid 单请求，支持任意学号；xscjlb=主修/辅修；pm/zrs=课程排名）
    async function getGrades(xh) {
        const j = await post('/cjgl/grcjcx/grcjcxByXjid', { xn: null, xq: null, kcmc: null, cxbj: '-1', pylx: '1', current: 1, pageSize: 500, xscjlb: null, sffx: null, xjid: xh });
        const list = j && j.content && j.content.list;
        if (!Array.isArray(list)) throw new Error(xh + ' 成绩查询失败');
        const rows = [];
        for (const r of list) {
            const score = (r.zpcj != null && r.zpcj !== '') ? r.zpcj : (r.zzcj != null ? r.zzcj : r.cj);
            if (score == null || score === '') continue; // 未出成绩的跳过
            rows.push({
                kcdm: r.kcdm, kcmc: r.kcmc, xf: r.xf, cj: String(score),
                xnxq: r.xnxqmc || r.xnxq || '', kclb: r.kclbmc || r.kclb || '', kcxz: r.kcxzmc || r.kcxz || '',
                kclbdm: r.kclbdm, xscjlb: safeText(r.xscjlb), pm: r.pm, zrs: r.zrs
            });
        }
        return { rows: rows, raw: list };
    }

    // ==================== 页面表格联动（复刻旧版 init(学号) 体验） ====================
    function walkForTableComp(rootApp) {
        let target = null;
        const walk = (v, depth) => {
            if (!v || depth > 8 || target) return;
            const d = v.$data || {};
            if (d.grcjcxcolumns1 && Array.isArray(d.data)) { target = v; return; }
            (v.$children || []).forEach(c => walk(c, depth + 1));
        };
        walk(rootApp, 0);
        return target;
    }
    function findTableComp() {
        try {
            // 1) 顶层窗口（直接打开 /cjgl/grcjcx/go/1 的场景）
            if (window.app) {
                const t = walkForTableComp(window.app);
                if (t) return t;
            }
            // 2) 主界面外壳：表格组件在 iframe 内（同源，可直接访问）
            const iframes = document.querySelectorAll("iframe");
            for (const f of iframes) {
                try {
                    const w2 = f.contentWindow;
                    if (w2 && w2.app) {
                        const t = walkForTableComp(w2.app);
                        if (t) return t;
                    }
                } catch (e) { }
            }
        } catch (e) { }
        return null;
    }
    // 把查询到的学生成绩切换进页面表格（备注列显示排名）
    function swapTable(rawRows, xh) {
        const comp = findTableComp();
        if (!comp) return false;
        try {
            comp.xjid = xh;
            comp.data = rawRows;
            comp.data.forEach(r => { if (r.pm != null && r.pm !== '' && r.zrs) r.cjbzmc = '排名 ' + r.pm + '/' + r.zrs; });
            return true;
        } catch (e) { return false; }
    }
    // 不在成绩页时，调用外壳函数打开“学生学习进度查询”标签页
    function openGrcjcxTab() {
        try {
            const fn = (window.parent && window.parent.bnavbtnFn) || window.bnavbtnFn;
            if (typeof fn === 'function') { fn('', 'u' + Date.now(), '学生学习进度查询', '/cjgl/grcjcx/go/1', ''); return true; }
        } catch (e) { }
        return false;
    }
    // 带重试的切换：页面加载慢时每秒重试，最多 attempts 次
    function swapTableWithRetry(rawRows, xh, attempts, done) {
        if (swapTable(rawRows, xh)) { if (done) done(true); return; }
        if (attempts <= 0) { if (done) done(false); return; }
        setTimeout(() => swapTableWithRetry(rawRows, xh, attempts - 1, done), 1000);
    }

    // ==================== 排除规则（沿用旧逻辑，字段名已映射） ====================
    const KCLBDM_EXCLUDE = { '55': '专业拓展', '27': '创新创业', '18': '国防公益' };
    const KCDM_EXCLUDE = new Set([
        '11501021', '11501022', '11501061', '11501062', '1150113',
        '1151002', '312010Y',
        'M159028', 'M159056', 'M159084', 'M159098', 'M159060',
        'M159065', 'M159066', 'M159069', 'M159072', 'M159081',
        'M159085', 'M159092', 'M159094', 'M159095', 'M159101',
        'M159103', 'M159106', 'M320009', 'M320025', 'M360042',
        'M360045', 'M360046', 'M360049', 'M360052', 'M360054',
        'M360056',
        '1099184', '1099096', '1099002', '1099150', '1099158',
        '1089123', '1089145', '1029005', '1029013', '10390028',
        '1189009', '1259004', '1289026', '1289071',
        '11590091', '11590094', '11590121', '11590124', '1159069',
        '1159071', '1159072', '1309001', '1309003',
        '1079009', '1070368',
        '70301192', '70706020', '71801007',
        '4070004', '4070014', '4071023',
        '2070001', '2070002', '2070153', '2070229', '2071006'
    ]);
    const KCDM_KEEP = new Set(['1120101', '308010A', '20602431', '20602432', '309000B', '321010C', '3211001', '306010D']);
    function safeText(v) { return v == null ? '' : String(v).trim(); }
    function toNum(v) { if (v == null || v === '') return null; const n = Number(v); return isFinite(n) ? n : null; }
    function matchAny(text, keywords) {
        const raw = safeText(text);
        if (!raw) return null;
        const parenMatches = [];
        raw.replace(/[（(]([^）)]*)[）)]/g, (m, inner) => { parenMatches.push(inner); return ''; });
        const stripped = raw.replace(/[（(].*?[）)]/g, '').trim();
        for (const kw of keywords) {
            const k = safeText(kw);
            if (!k) continue;
            if (raw.includes(k)) return k;
            for (const inner of parenMatches) if (inner.includes(k)) return k;
            if (stripped.includes(k)) return k;
        }
        return null;
    }
    function getExcludeReason(item) {
        const kcdm = safeText(item.kcdm);
        const name = safeText(item.kcmc);
        if (safeText(item.xscjlb).includes('辅修')) return '辅修';
        if (KCDM_EXCLUDE.has(kcdm)) return '手动排除(' + kcdm + ')';
        if (KCDM_KEEP.has(kcdm)) return null;
        const xf = toNum(item.xf);
        if (xf == null || xf <= 0) return '学分≤0';
        const scoreText = safeText(item.cj);
        if (scoreText === '通过') return '成绩=通过';
        const kclbmc = safeText(item.kclb);
        const kclbdm = safeText(item.kclbdm);
        if (kclbdm && KCLBDM_EXCLUDE[kclbdm]) return '类型[' + KCLBDM_EXCLUDE[kclbdm] + ']';
        if (matchAny(kclbmc, ['专业拓展', '素质拓展', '创新创业', '国防公益'])) return '类型[' + kclbmc + ']';
        const allText = [name, kclbmc, safeText(item.kcxz)].join(' | ');
        const gl = matchAny(allText, ['辅修', '二学位', '第二学位', '缓考', '旷考', '作弊', '取消资格', '缺考', '形势与政策']);
        if (gl) return '全局[' + gl + ']';
        return null;
    }
    function calc(rows, xnxqFilter) {
        let totalS = 0, totalC = 0, valid = 0, excl = 0;
        const details = [];
        const seen = new Set();
        for (const item of rows) {
            const key = (item.kcdm || '') + '|' + (item.xnxq || '');
            if (seen.has(key)) continue;
            seen.add(key);
            const xnxq = safeText(item.xnxq);
            if (xnxqFilter && !xnxq.startsWith(xnxqFilter)) continue;
            const reason = getExcludeReason(item);
            const xf = toNum(item.xf);
            const scoreText = safeText(item.cj);
            const score = toNum(scoreText);
            let counted = false;
            if (reason) { excl++; }
            else if (xf && xf > 0 && score !== null) { totalS += xf * score; totalC += xf; valid++; counted = true; }
            else { excl++; }
            details.push({
                xnxq: xnxq, kcdm: safeText(item.kcdm), kcmc: safeText(item.kcmc),
                xf: xf, score: scoreText, counted: counted ? '是' : '否',
                pm: safeText(item.pm), zrs: safeText(item.zrs), xscjlb: safeText(item.xscjlb),
                reason: reason || (counted ? '' : (score === null ? '成绩非数字' : '其他'))
            });
        }
        return { score: totalC > 0 ? totalS / totalC : null, credits: totalC, valid: valid, excluded: excl, total: details.length, details: details };
    }

    // ==================== SheetJS 按需加载（多CDN兜底） ====================
    function ensureXLSX() {
        return new Promise((resolve, reject) => {
            if (typeof XLSX !== 'undefined') { resolve(); return; }
            const cdns = [
                'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
                'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js',
                'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
            ];
            let i = 0;
            const tryNext = () => {
                if (i >= cdns.length) { reject(new Error('SheetJS 三个CDN都加载失败，请检查网络')); return; }
                const s = document.createElement('script');
                s.src = cdns[i];
                s.onload = () => { if (typeof XLSX !== 'undefined') { log('XLSX 已从', cdns[i].split('/')[2], '加载'); resolve(); } else { i++; tryNext(); } };
                s.onerror = () => { i++; tryNext(); };
                document.head.appendChild(s);
            };
            tryNext();
        });
    }

    // ==================== 面板 ====================
    let panel = null;
    function ensurePanel() {
        if (panel && panel.parentNode) return panel;
        if (!document.body) return null;
        panel = document.createElement('div');
        panel.id = 'ustb-panel-v10';
        panel.style.cssText = 'position:fixed;bottom:10px;right:10px;z-index:2147483646;background:rgba(0,0,0,0.88);color:#fff;border-radius:8px;font-size:13px;line-height:1.6;width:430px;box-shadow:0 2px 10px rgba(0,0,0,0.5);user-select:none;overflow:hidden;';
        panel.innerHTML =
            '<div id="ustb-header" style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;cursor:move;background:rgba(255,255,255,0.1);font-weight:bold;">' +
                '<span>📊 USTB 成绩增强 v10</span>' +
                '<button id="ustb-close" style="background:transparent;border:none;color:#fff;font-size:18px;cursor:pointer;padding:0 4px;line-height:1;">×</button>' +
            '</div>' +
            '<div id="ustb-content" style="padding:10px 12px;">等待操作…</div>' +
            '<div id="ustb-footer" style="padding:0 12px 12px 12px;border-top:1px solid rgba(255,255,255,0.08);padding-top:10px;">' +
                '<input id="ustb-xh" type="text" placeholder="输入单个学号查询" style="width:100%;padding:4px 8px;font-size:12px;border:1px solid #555;background:#222;color:#fff;border-radius:4px;box-sizing:border-box;margin-bottom:6px;">' +
                '<input id="ustb-xnxq" type="text" placeholder="学年筛选（留空=全部，如 2024-2025）" style="width:100%;padding:4px 8px;font-size:12px;border:1px solid #555;background:#222;color:#fff;border-radius:4px;box-sizing:border-box;margin-bottom:8px;">' +
                '<button id="ustb-one-btn" style="width:100%;padding:6px;font-size:12px;cursor:pointer;border:none;background:#1890ff;color:#fff;border-radius:4px;font-weight:bold;margin-bottom:6px;">🔍 查询单人加权</button>' +
                '<button id="ustb-batch-btn" style="width:100%;padding:6px;font-size:12px;cursor:pointer;border:none;background:#52c41a;color:#fff;border-radius:4px;font-weight:bold;">📁 批量查询（导入 Excel）</button>' +
            '</div>';
        document.body.appendChild(panel);
        document.getElementById('ustb-close').addEventListener('click', () => {
            panel.style.display = 'none';
            if (!document.getElementById('ustb-reopen')) {
                const r = document.createElement('button');
                r.id = 'ustb-reopen'; r.textContent = '📊';
                r.style.cssText = 'position:fixed;bottom:10px;right:10px;z-index:2147483646;width:40px;height:40px;border-radius:50%;background:rgba(0,0,0,0.8);color:#fff;border:none;cursor:pointer;font-size:18px;';
                r.addEventListener('click', () => { panel.style.display = 'block'; r.remove(); });
                document.body.appendChild(r);
            }
        });
        document.getElementById('ustb-one-btn').addEventListener('click', onOneClick);
        document.getElementById('ustb-batch-btn').addEventListener('click', onBatchClick);
        let drag = false, sx, sy, ox, oy;
        const header = document.getElementById('ustb-header');
        header.addEventListener('mousedown', e => {
            if (e.target.id === 'ustb-close') return;
            drag = true; sx = e.clientX; sy = e.clientY;
            const rect = panel.getBoundingClientRect(); ox = rect.left; oy = rect.top;
            panel.style.right = 'auto'; panel.style.bottom = 'auto';
            panel.style.left = ox + 'px'; panel.style.top = oy + 'px';
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
            e.preventDefault();
        });
        function onMove(e) { if (!drag) return; panel.style.left = (ox + e.clientX - sx) + 'px'; panel.style.top = (oy + e.clientY - sy) + 'px'; }
        function onUp() { drag = false; document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
        return panel;
    }
    function setContent(html) {
        const c = document.getElementById('ustb-content');
        if (c) c.innerHTML = html;
    }

    // ==================== 单人查询 ====================
    async function onOneClick() {
        const xh = safeText(document.getElementById('ustb-xh').value);
        if (!xh) { alert('请输入学号'); return; }
        const xnxq = safeText(document.getElementById('ustb-xnxq').value);
        setContent('⏳ 查询 ' + xh + ' …');
        try {
            const { rows, raw } = await getGrades(xh);
            const stats = calc(rows, xnxq);
            const rankCount = rows.filter(r => r.pm != null && r.pm !== '').length;
            const swapped = swapTable(raw || [], xh);
            if (!swapped) {
                openGrcjcxTab();
                setContent(
                    '<div>✅ ' + xh + '（共 ' + stats.total + ' 门记录，' + rankCount + ' 门有排名）</div>' +
                    '<div style="margin-top:4px;">📈 加权均分：<b style="font-size:16px;">' + (stats.score != null ? stats.score.toFixed(4) : 'N/A') + '</b></div>' +
                    '<div>学分 ' + stats.credits.toFixed(1) + ' ｜ 计入 ' + stats.valid + ' 门 ｜ 排除 ' + stats.excluded + ' 门</div>' +
                    '<div style="color:#9fe;margin-top:4px;">📋 正在打开成绩页并切换表格（最长等10秒）…</div>'
                );
                swapTableWithRetry(raw || [], xh, 10, (ok) => {
                    setContent(
                        '<div>✅ ' + xh + '（共 ' + stats.total + ' 门记录，' + rankCount + ' 门有排名）</div>' +
                        '<div style="margin-top:4px;">📈 加权均分：<b style="font-size:16px;">' + (stats.score != null ? stats.score.toFixed(4) : 'N/A') + '</b></div>' +
                        '<div>学分 ' + stats.credits.toFixed(1) + ' ｜ 计入 ' + stats.valid + ' 门 ｜ 排除 ' + stats.excluded + ' 门</div>' +
                        (ok ? '<div style="color:#9fe;margin-top:4px;">📋 表格已切换为该学生成绩（备注列含排名）</div>' : '<div style="color:#ffb;font-size:12px;">表格未能自动切换——请手动点开“学生学习进度查询”页后再查一次</div>')
                    );
                });
                return;
            }
            setContent(
                '<div>✅ ' + xh + '（共 ' + stats.total + ' 门记录，' + rankCount + ' 门有排名）</div>' +
                '<div style="margin-top:4px;">📈 加权均分：<b style="font-size:16px;">' + (stats.score != null ? stats.score.toFixed(4) : 'N/A') + '</b></div>' +
                '<div>学分 ' + stats.credits.toFixed(1) + ' ｜ 计入 ' + stats.valid + ' 门 ｜ 排除 ' + stats.excluded + ' 门' + (xnxq ? ' ｜ 筛选 ' + xnxq : '') + '</div>' +
                '<div style="color:#9fe;margin-top:4px;">📋 上方表格已切换为该学生成绩（备注列含排名）</div>'
            );
        } catch (e) {
            setContent('❌ 查询失败：' + e.message);
        }
    }

    // ==================== 批量查询 ====================
    let batchRunning = false;
    function readStudentsFromExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    if (!rows.length) { resolve([]); return; }
                    const header = rows[0].map(h => String(h || '').trim());
                    let xhCol = -1, nameCol = -1;
                    for (let i = 0; i < header.length; i++) {
                        if (/学号/.test(header[i])) xhCol = i;
                        if (/姓名/.test(header[i])) nameCol = i;
                    }
                    if (xhCol < 0) xhCol = 0;
                    if (nameCol < 0 && header.length > 1) nameCol = 1;
                    const students = [];
                    for (let i = 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row) continue;
                        const xh = String(row[xhCol] || '').trim();
                        if (!xh || !/^[A-Za-z0-9]/.test(xh)) continue;
                        students.push({ xh: xh, name: nameCol >= 0 ? String(row[nameCol] || '').trim() : '' });
                    }
                    resolve(students);
                } catch (err) { reject(err); }
            };
            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }
    function onBatchClick() {
        if (batchRunning) { alert('批量查询正在进行中'); return; }
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls,.csv';
        input.style.display = 'none';
        document.body.appendChild(input);
        input.addEventListener('change', async () => {
            const file = input.files && input.files[0];
            input.remove();
            if (!file) return;
            try {
                const students = await readStudentsFromExcel(file);
                if (!students.length) { alert('Excel 里没找到学号'); return; }
                const xnxq = safeText(document.getElementById('ustb-xnxq').value);
                if (!confirm('共 ' + students.length + ' 个学生。\n学年筛选：' + (xnxq || '全部') + '\n每人约 2-5 秒（逐类别查询），是否开始？')) return;
                await runBatch(students, xnxq);
            } catch (e) {
                alert('读取 Excel 失败：' + e.message);
            }
        });
        input.click();
    }
    async function runBatch(students, xnxq) {
        batchRunning = true;
        const results = [];
        const btn = document.getElementById('ustb-batch-btn');
        if (btn) { btn.disabled = true; btn.style.background = '#666'; btn.textContent = '⏳ 查询中...'; }
        for (let i = 0; i < students.length; i++) {
            const s = students[i];
            setContent('<div>📊 批量查询中…</div><div>进度：' + (i + 1) + ' / ' + students.length + '</div><div>当前：' + s.xh + ' ' + s.name + '</div>');
            try {
                const { rows, raw } = await getGrades(s.xh);
                swapTable(raw || [], s.xh); // 旧版行为：批量时表格随查询切换
                const stats = calc(rows, xnxq);
                results.push({ xh: s.xh, name: s.name, score: stats.score, credits: stats.credits, valid: stats.valid, excluded: stats.excluded, total: stats.total, error: '', _details: stats.details });
                log('✓', s.xh, s.name, stats.score != null ? stats.score.toFixed(2) : 'null');
            } catch (err) {
                results.push({ xh: s.xh, name: s.name, score: null, credits: 0, valid: 0, excluded: 0, total: 0, error: err.message || String(err), _details: [] });
                log('✗', s.xh, s.name, '失败:', err.message);
            }
            await new Promise(r => setTimeout(r, 500));
        }
        const validList = results.filter(r => r.score != null);
        validList.sort((a, b) => b.score - a.score);
        validList.forEach((r, i) => { r.rank = i + 1; });
        results.forEach(r => { if (r.rank == null) r.rank = ''; });
        try {
            exportResults(results, xnxq);
            const okCount = results.filter(r => r.score != null).length;
            setContent('<div>✅ 批量完成：成功 ' + okCount + ' / 失败 ' + (results.length - okCount) + '</div><div style="font-size:11px;color:#8f8;">Excel 已下载</div>');
        } catch (e) {
            alert('导出失败：' + e.message);
        }
        batchRunning = false;
        if (btn) { btn.disabled = false; btn.style.background = '#52c41a'; btn.textContent = '📁 批量查询（导入 Excel）'; }
    }
    async function exportResults(results, xnxq) {
        await ensureXLSX();
        const sorted = [...results].sort((a, b) => {
            if (a.rank === '' && b.rank !== '') return 1;
            if (a.rank !== '' && b.rank === '') return -1;
            return (Number(a.rank) || 999) - (Number(b.rank) || 999);
        });
        const summary = sorted.map(r => ({
            '学号': r.xh, '姓名': r.name, '排名': r.rank,
            '加权均分': r.score != null ? Number(r.score.toFixed(4)) : '',
            '学分': Number(r.credits.toFixed(1)),
            '有效门数': r.valid, '备注': r.error || ''
        }));
        const ws1 = XLSX.utils.json_to_sheet(summary);
        ws1['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 6 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 30 }];
        const details = [];
        for (const r of sorted) {
            if (!r._details) continue;
            for (const d of r._details) {
                details.push({ '学号': r.xh, '姓名': r.name, '学年学期': d.xnxq, '课程代码': d.kcdm, '课程名称': d.kcmc, '学分': d.xf, '总评成绩': d.score, '课程排名': (d.pm && d.zrs) ? d.pm + '/' + d.zrs : '', '学生成绩类别': d.xscjlb, '是否计入': d.counted, '备注': d.reason });
            }
        }
        const ws2 = XLSX.utils.json_to_sheet(details);
        ws2['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 30 }, { wch: 6 }, { wch: 8 }, { wch: 8 }, { wch: 20 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws1, '汇总');
        XLSX.utils.book_append_sheet(wb, ws2, '明细');
        const ts = new Date();
        const fn = 'USTB批量查询v10' + (xnxq ? '_' + xnxq : '') + '_' + ts.getFullYear() + String(ts.getMonth() + 1).padStart(2, '0') + String(ts.getDate()).padStart(2, '0') + '_' + String(ts.getHours()).padStart(2, '0') + String(ts.getMinutes()).padStart(2, '0') + '.xlsx';
        XLSX.writeFile(wb, fn);
        log('已导出:', fn);
    }

    // ==================== 启动 ====================
    function start() {
        log('启动 v10.0.0，当前页面:', location.href);
        ensurePanel();
        setContent('✅ 已就绪（适配 2026 改版接口）<div style="font-size:11px;color:#aaa;margin-top:4px;">输入学号查单人，或导入 Excel 批量查询。<br>注意：新版接口不再返回课程排名(pm/zrs)，排名徽章功能已移除。</div>');
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
    else start();
})();
