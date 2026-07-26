// ═══════════════════════════════════════════════════════════════════
// PRESET SCANNER — le os .prfpset do Premiere (predefinicoes de efeito) do
// disco, no PROPRIO app (Node), FORA da engine ExtendScript. Assim:
//   - Catalogo de presets fica pronto no BOOT (primeira vez rapida)
//   - fs.watch detecta preset novo salvo e re-scaneia (sem reabrir painel)
// Produz itens no MESMO formato que o plugin ja consumia (matchName =
// "<caminho.prfpset>#<nome>"), entao o apply do plugin continua igual.
//
// Porta fiel do parser do plugin (_lsParseBinTree / _lsXmlUnescape).
// ═══════════════════════════════════════════════════════════════════
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');

function xmlUnescape(s) {
    return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'").replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(+d); })
        .replace(/&amp;/g, '&');
}

// Anda a arvore de bins do .prfpset: BinTreeItem = pasta, TreeItem folha = preset.
// Retorna [{ name, category }] — igual ao que o painel Predefinicoes do Premiere mostra.
function parseBinTree(xml) {
    try {
        const nodes = {};
        const re = /<(BinTreeItem|TreeItem)\s+ObjectID="(\d+)"[^>]*>/g;
        let m;
        const starts = [];
        while ((m = re.exec(xml)) !== null) starts.push({ tag: m[1], id: m[2], idx: m.index });
        if (!starts.length) return [];
        for (let i = 0; i < starts.length; i++) {
            const s = starts[i];
            const end = xml.indexOf('</' + s.tag + '>', s.idx);
            if (end < 0) continue;
            nodes[s.id] = { tag: s.tag, body: xml.slice(s.idx, end) };
        }
        const nameOf = (n) => { const mm = n.body.match(/<Name>([^<]*)<\/Name>/); return mm ? xmlUnescape(mm[1]).trim() : ''; };
        const refsOf = (n) => {
            const out = [];
            const items = n.body.match(/<Items[^>]*>([\s\S]*?)<\/Items>/);
            if (items) { const r = /<Item Index="\d+" ObjectRef="(\d+)"\/>/g; let k; while ((k = r.exec(items[1])) !== null) out.push(k[1]); }
            return out;
        };
        let rootId = null;
        for (const id in nodes) { if (nodes[id].tag === 'BinTreeItem' && nameOf(nodes[id]) === 'Root') { rootId = id; break; } }
        if (!rootId) return [];
        const out = [], seen = {};
        let guard = 0;
        (function walk(nid, folder) {
            if (guard++ > 30000) return;
            const n = nodes[nid]; if (!n) return;
            const nm = nameOf(n);
            if (n.tag === 'BinTreeItem') {
                const skipName = (nm === 'Root' || !nm || n.body.indexOf('.PresetsBin>') >= 0);
                const p = skipName ? folder : folder.concat(nm);
                const kids = refsOf(n);
                for (let i2 = 0; i2 < kids.length; i2++) walk(kids[i2], p);
            } else if (nm) {
                const key = nm + '|' + folder.join('/');
                if (!seen[key]) { seen[key] = 1; out.push({ name: nm, category: folder.join(' / ') }); }
            }
        })(rootId, []);
        return out;
    } catch (e) { return []; }
}

// Descobre os .prfpset do usuario nas pastas de perfil do Premiere.
function findPrfpsetFiles() {
    const out = [];
    const seen = {};
    const home = os.homedir();
    const appdata = process.env.APPDATA || path.join(home, 'AppData', 'Roaming');
    const roots = [
        path.join(appdata, 'Adobe', 'Premiere Pro'),                              // Win Roaming (arquivo mestre)
        path.join(home, 'Documents', 'Adobe', 'Premiere Pro'),                     // Win/Mac Documents
        path.join(home, 'OneDrive', 'Documents', 'Adobe', 'Premiere Pro'),         // Win + OneDrive
        path.join(home, 'OneDrive', 'Documentos', 'Adobe', 'Premiere Pro'),        // OneDrive pt-BR
        path.join(home, 'Library', 'Application Support', 'Adobe', 'Premiere Pro'), // Mac
    ];
    const SKIP = /^(Media Cache|Media Cache Files|Peak Files|Preview Files|CFA Files|Auto-Save|Captured Audio|Adobe Premiere Pro Auto-Save)$/i;
    function walk(dir, depth) {
        if (depth > 5) return;
        let entries;
        try { entries = fs.readdirSync(dir); } catch (e) { return; }
        for (let i = 0; i < entries.length; i++) {
            const nm = entries[i];
            const full = path.join(dir, nm);
            let st; try { st = fs.statSync(full); } catch (e) { continue; }
            if (st.isDirectory()) { if (!SKIP.test(nm)) walk(full, depth + 1); }
            else if (/\.prfpset$/i.test(nm)) {
                const key = full.toLowerCase();
                if (seen[key]) continue;
                seen[key] = true;
                out.push({ file: full, name: nm.replace(/\.prfpset$/i, ''), mtimeMs: st.mtimeMs, size: st.size });
            }
        }
    }
    for (let r = 0; r < roots.length; r++) { try { walk(roots[r], 0); } catch (e) {} }
    // Ordena por versao do Premiere (mais nova primeiro) e depois mtime — quem parsear primeiro
    // ganha na deduplicacao por nome, entao a versao MAIS NOVA do preset prevalece.
    const verOf = (o) => { const mm = String(o.file).match(/Premiere Pro[\\/](\d+(?:\.\d+)?)/i); return mm ? parseFloat(mm[1]) : 0; };
    out.sort((a, b) => (verOf(b) - verOf(a)) || (b.mtimeMs - a.mtimeMs));
    return out;
}

// Parseia UM .prfpset → itens de preset prontos pro catalogo do Lion Search.
function parsePresetFile(fileInfo) {
    try {
        const data = fs.readFileSync(fileInfo.file);
        let xml;
        if (data.length > 2 && data[0] === 0x1f && data[1] === 0x8b) xml = zlib.gunzipSync(data).toString('utf8');
        else xml = data.toString('utf8');
        if (!xml) return [];
        const entries = parseBinTree(xml);
        const items = [];
        for (let i = 0; i < entries.length; i++) {
            const pn = entries[i].name;
            if (!pn || pn.toLowerCase() === fileInfo.name.toLowerCase()) continue; // pula o proprio container
            items.push({
                kind: 'preset',
                name: pn,
                matchName: fileInfo.file + '#' + pn, // caminho do .prfpset + nome (o apply do plugin usa isso)
                category: entries[i].category || 'Presets',
                _appScanned: true,
            });
        }
        return items;
    } catch (e) { return []; }
}

// Scan completo: acha os arquivos + parseia todos → { items, files, ms }
function scanPresets() {
    const t0 = Date.now();
    const files = findPrfpsetFiles(); // ja vem ordenado versao-mais-nova primeiro
    const seen = {};
    const deduped = [];
    for (let i = 0; i < files.length; i++) {
        const fileItems = parsePresetFile(files[i]);
        for (let j = 0; j < fileItems.length; j++) {
            // dedup por nome+pasta (case-insensitive): mesmo preset em 25.0 e 26.0 vira 1 so
            // (o do arquivo mais novo ganha, porque vem primeiro na lista).
            const k = (fileItems[j].name + '|' + fileItems[j].category).toLowerCase();
            if (seen[k]) continue;
            seen[k] = 1; deduped.push(fileItems[j]);
        }
    }
    return { items: deduped, files: files, ms: Date.now() - t0 };
}

module.exports = { scanPresets, findPrfpsetFiles, parsePresetFile, parseBinTree, xmlUnescape };
