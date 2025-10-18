/***** SINGLE-HASH CONFIG (tree) — Simplified (no rowSelector) *****************
 * Each node may specify:
 *  - id:               unique id (required)
 *  - label:            display name for $ mode (required)
 *  - alternate_label:  optional display name for % mode (anonymized or alternate)
 *  - sheet/col:        where to read the value for this node (optional)
 *  - method:           'latestPositive' | 'latestNonEmpty' | 'sumColumn' (default: latestPositive)
 *  - children:         array of child nodes (optional)
 * ...
 ******************************************************************************/

const CONFIG = {
};

/** ==== UI ==== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Net Worth Sankey')
    .addItem('Insert Diagram', 'showNetWorthSankey')
    .addToUi();
}

function showNetWorthSankey() {
  const html = HtmlService.createHtmlOutputFromFile('sankey')
    .setWidth(980)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Insert Diagram');
}

/** ==== Helpers ==== */
const COL_CACHE = {};  // { "<sheet>.<colLetter>": [values...] }
// Leaf sequence counter for stable, contiguous, leaf-driven ordering
let LEAF_SEQ = 0;

function colLetterToIndex(letter) {
  letter = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letter)) throw new Error('Invalid column letter: ' + letter);
  let idx = 0;
  for (let i = 0; i < letter.length; i++) idx = idx * 26 + (letter.charCodeAt(i) - 64);
  return idx; // 1-based
}
function asNumber(v) {
  if (v === '' || v === null) return NaN;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[^0-9.\-]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}
function magnitude(v) {
  const n = asNumber(v);
  return Number.isFinite(n) ? Math.abs(n) : NaN;
}
function getColValues(ss, sheetName, colLetter) {
  const key = `${sheetName}.${colLetter}`;
  if (COL_CACHE[key]) return COL_CACHE[key];
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found.`);
  const lastRow = sh.getLastRow();
  const vals = lastRow >= 1 ? sh.getRange(1, colLetterToIndex(colLetter), lastRow, 1).getValues().flat() : [];
  COL_CACHE[key] = vals;
  return vals;
}

/** Row picking for a single column */
function pickRowIndex(vals, method) {
  method = method || 'latestPositive';
  const last = vals.length;
  if (method === 'sumColumn') return -2; // sentinel for "use column sum"
  if (method === 'latestNonEmpty') {
    for (let i = last - 1; i >= 0; i--) { if (!isNaN(asNumber(vals[i]))) return i; }
    return -1;
  }
  // default: latestPositive
  for (let i = last - 1; i >= 0; i--) {
    const v = magnitude(vals[i]);
    if (Number.isFinite(v) && v > 0) return i;
  }
  return -1;
}

/** ==== Tree evaluation -> {nodes, links} with leaf-driven ordering ==== */
/**
 * DFS that:
 *  - computes each node's label amount (from sheet/col or sum of children),
 *  - builds links, and
 *  - assigns a stable 'order' derived from contiguous deepest-leaf indices
 *    so siblings (and their descendants) stay grouped together recursively.
 */
function evaluateTree(node, ss, depth = 0) {
  // 1) Resolve this node's "label amount" from sheet/col, if provided
  let amount = null;
  if (node.sheet && node.col) {
    const colVals = getColValues(ss, node.sheet, node.col);
    const rowIdx = pickRowIndex(colVals, node.method || 'latestPositive');
    if (rowIdx === -2) {
      // sumColumn: sum absolute values
      const sum = colVals.reduce((acc, v) => {
        const m = magnitude(v); return Number.isFinite(m) ? acc + m : acc;
      }, 0);
      amount = sum > 0 ? sum : null;
    } else if (rowIdx >= 0) {
      const m = magnitude(colVals[rowIdx]);
      amount = Number.isFinite(m) ? m : null;
    }
  }

  const nodes = [];
  const links = [];
  const me = { id: node.id, label: node.label, alternate_label: node.alternate_label || null, amount: amount, depth: depth };
  nodes.push(me);

  let childrenTotal = 0;
  let leafMin = Infinity;
  let leafMax = -Infinity;

  if (Array.isArray(node.children) && node.children.length) {
    // Preserve CONFIG.children order in DFS (left-to-right)
    for (const child of node.children) {
      const { nodes: cnodes, links: clinks, total: childTotal, leafRange } = evaluateTree(child, ss, depth + 1);
      nodes.push(...cnodes);
      links.push(...clinks);
      if (Number.isFinite(childTotal) && childTotal > 0) {
        links.push({ source: node.id, target: child.id, value: childTotal });
        childrenTotal += childTotal;
      }
      if (leafRange) {
        leafMin = Math.min(leafMin, leafRange.min);
        leafMax = Math.max(leafMax, leafRange.max);
      }
    }
  } else {
    // This is a leaf → assign strict left-to-right position
    leafMin = leafMax = LEAF_SEQ++;
  }

  // 3) Label fallback: if no explicit amount but has children, display sum(children)
  if ((!Number.isFinite(me.amount) || me.amount === null) && childrenTotal > 0) {
    me.amount = childrenTotal;
  }

  // 4) FLOW up the tree (used for parent→this link)
  const flowOut = (childrenTotal > 0)
    ? childrenTotal
    : (Number.isFinite(me.amount) ? me.amount : 0);

  // 5) Stable node order from deepest-leaf span (midpoint keeps siblings grouped)
  if (Number.isFinite(leafMin) && Number.isFinite(leafMax)) {
    me.order = (leafMin + leafMax) / 2;
    // Optional tiny epsilon for strict, stable tiebreaks by depth:
    // me.order += depth * 1e-6;
  } else {
    // Fallback if no leaves found under this node
    me.order = Number.isFinite(me.amount) ? me.amount : 0;
  }

  return {
    nodes,
    links,
    total: flowOut,
    leafRange: (Number.isFinite(leafMin) && Number.isFinite(leafMax)) ? { min: leafMin, max: leafMax } : null
  };
}

/** Main entry used by HTML */
function getNetWorthPayload() {
  // Clear caches per call
  for (const k in COL_CACHE) delete COL_CACHE[k];
  // Reset leaf sequence so orders are consistent per run
  LEAF_SEQ = 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { nodes, links } = evaluateTree(CONFIG.root, ss);

  // Filter zero/NaN links
  const cleanLinks = links.filter(l => Number.isFinite(l.value) && l.value > 0);

  return { nodes, links: cleanLinks };
}

/** ==== Chunked upload helpers (Drive-backed) ==== */

const TEMP_FOLDER_NAME = 'SankeyTempUploads';

/** Get or create temp folder for chunk files */
function getTempFolder_() {
  const it = DriveApp.getFoldersByName(TEMP_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(TEMP_FOLDER_NAME);
}

/** Start an image upload session; returns {sessionId} */
function startImageUpload() {
  const sessionId = Utilities.getUuid();
  // Optionally pre-create a subfolder per session (not required), we just prefix file names.
  return { sessionId };
}

/** Append a base64 chunk (string) for a session */
function uploadImageChunk(sessionId, seq, base64Chunk) {
  if (!sessionId || typeof seq !== 'number' || !base64Chunk) throw new Error('Invalid chunk');
  const folder = getTempFolder_();
  const name = `${sessionId}_${String(seq).padStart(6, '0')}.part`;
  folder.createFile(name, base64Chunk, MimeType.PLAIN_TEXT);
  return true;
}

/** Finalize: read all chunks, build blob, insert image, then cleanup */
function finalizeImageUpload(sessionId, mimeSubtype, filename) {
  if (!sessionId) throw new Error('Missing sessionId');
  const folder = getTempFolder_();

  // Collect chunks
  const files = [];
  const it = folder.getFiles();
  const prefix = `${sessionId}_`;
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (name.startsWith(prefix) && name.endsWith('.part')) files.push(f);
  }
  if (!files.length) throw new Error('No chunks found for session ' + sessionId);

  // Sort by sequence
  files.sort((a, b) => a.getName().localeCompare(b.getName()));

  // Concatenate base64 text
  let base64 = '';
  files.forEach(f => {
    base64 += f.getBlob().getDataAsString();
  });

  // Decode & create blob
  const mime = 'image/' + (mimeSubtype || 'jpeg');
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime, filename || ('sankey.' + (mimeSubtype === 'png' ? 'png' : 'jpg')));

  // Insert at active cell
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const range = sh.getActiveRange();
  const row = range ? range.getRow() : 1;
  const col = range ? range.getColumn() : 1;
  sh.insertImage(blob, col, row);

  // Cleanup
  files.forEach(f => { try { f.setTrashed(true); } catch(e) {} });

  return true;
}
