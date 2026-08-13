import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import templateBase64 from './template-base64.js';

const ORIGIN = 'https://handsome10219.github.io';
const cors = {
  'Access-Control-Allow-Origin': ORIGIN,
  'Access-Control-Allow-Methods': 'GET,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

function base64Bytes(s) {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function excelSerial(dateText) {
  const [y,m,d] = dateText.split('-').map(Number);
  return Math.floor((Date.UTC(y,m-1,d) - Date.UTC(1899,11,30)) / 86400000);
}
function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function setCellValue(xml, ref, value) {
  const re = new RegExp(`(<c\\b[^>]*\\br="${escRe(ref)}"[^>]*>)([\\s\\S]*?)(<\\/c>)`);
  return xml.replace(re, (all, open, inner, close) => {
    const cleanOpen = open.replace(/\s+t="[^"]*"/g, '');
    const withoutV = inner.replace(/<v>[\s\S]*?<\/v>/g, '');
    const v = value == null || value === '' ? '' : `<v>${Number(value)}</v>`;
    return cleanOpen + withoutV + v + close;
  });
}
function clearCellValue(xml, ref) { return setCellValue(xml, ref, null); }
function forceRecalc(xml) {
  if (/<calcPr\b/.test(xml)) {
    return xml.replace(/<calcPr\b([^>]*)\/?>(?:<\/calcPr>)?/, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>');
  }
  return xml.replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (url.pathname === '/api/health') return Response.json({ ok: true }, { headers: cors });
    if (url.pathname !== '/api/export-excel' || request.method !== 'GET') {
      return Response.json({ error: 'Not found' }, { status: 404, headers: cors });
    }
    try {
      const { results } = await env.DB.prepare(`
        SELECT record_date AS date, weight, bmi, body_fat AS bodyFat,
               muscle, bmr, body_age AS bodyAge, visceral_fat AS visceralFat,
               fat_weight AS fatWeight
        FROM body_records WHERE user_id='main' ORDER BY record_date ASC
      `).all();
      const settings = await env.DB.prepare(`SELECT goal_weight AS goalWeight FROM user_settings WHERE user_id='main'`).first();
      if (!results?.length) return Response.json({ error: '目前沒有可匯出的紀錄' }, { status: 400, headers: cors });
      if (results.length > 64) return Response.json({ error: '目前母版支援最多 64 筆紀錄' }, { status: 400, headers: cors });

      const files = unzipSync(base64Bytes(templateBase64));
      let sheet = strFromU8(files['xl/worksheets/sheet1.xml']);
      // 清除母版原本 A:H 的示範資料，再寫入 D1 紀錄；I:W 的公式、樣式、圖表關聯完全保留。
      for (let row = 2; row <= 65; row++) for (const col of ['A','B','C','D','E','F','G','H']) sheet = clearCellValue(sheet, `${col}${row}`);
      results.forEach((r, i) => {
        const row = i + 2;
        const vals = { A: excelSerial(r.date), B:r.weight, C:r.bmi, D:r.bodyFat, E:r.muscle, F:r.bmr, G:r.bodyAge, H:r.visceralFat };
        for (const [col,val] of Object.entries(vals)) sheet = setCellValue(sheet, `${col}${row}`, val);
      });
      files['xl/worksheets/sheet1.xml'] = strToU8(sheet);

      let settingsXml = strFromU8(files['xl/worksheets/sheet6.xml']);
      settingsXml = setCellValue(settingsXml, 'B3', results[0].weight);
      settingsXml = setCellValue(settingsXml, 'B4', settings?.goalWeight ?? 65);
      files['xl/worksheets/sheet6.xml'] = strToU8(settingsXml);
      files['xl/workbook.xml'] = strToU8(forceRecalc(strFromU8(files['xl/workbook.xml'])));

      const out = zipSync(files, { level: 6 });
      const stamp = new Date().toISOString().slice(0,10);
      return new Response(out, { headers: {
        ...cors,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="weighttracker_${stamp}.xlsx"`,
        'Cache-Control': 'no-store'
      }});
    } catch (e) {
      console.error(e);
      return Response.json({ error: e?.message || String(e) }, { status: 500, headers: cors });
    }
  }
};
