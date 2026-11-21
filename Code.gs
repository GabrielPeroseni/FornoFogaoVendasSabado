/**
 * Forno & Fogão – banco de dados no Google Planilhas
 * Cole este código em Extensões > Apps Script (dentro da planilha),
 * troque a CHAVE abaixo e publique como "App da Web".
 */
const CHAVE = "TROQUE-ESTA-CHAVE";   // mesma chave digitada no sistema

const ABAS = {
  prod: { nome: "Produtos",  cols: ["id","nome","cat","un","cod","preco","ativo","desc"],            texto: ["id","nome","cat","un","ativo","desc"] },
  ped:  { nome: "Comandas",  cols: ["id","tipo","cliente","hora","status","endereco","numero","complemento","bairro","pagamento","taxa","extra","json"], texto: ["id","tipo","cliente","hora","status","endereco","numero","complemento","bairro","pagamento","json"] },
  itens:{ nome: "Itens",     cols: ["comanda","produto_id","produto","qtd","etiqueta"],               texto: ["comanda","produto_id","produto","etiqueta"] },
  est:  { nome: "Estoque",   cols: ["produto_id","produzida"],                                        texto: ["produto_id"] },
  end:  { nome: "Enderecos", cols: ["nome","faixa","bairro","taxa"],                                  texto: ["nome","faixa","bairro"] },
  cfg:  { nome: "Config",    cols: ["chave","valor"],                                                 texto: ["chave"] }
};

function doGet(e)  { return saida_(() => { auth_(e.parameter.k); return e.parameter.action === "all" ? lerTudo_() : { ok: true, rev: rev_() }; }); }
function doPost(e) { return saida_(() => { const b = JSON.parse(e.postData.contents); auth_(b.k); return sincronizar_(b); }); }

function saida_(fn) {
  let r;
  try { r = fn(); } catch (err) { r = { ok: false, err: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function auth_(k) { if (CHAVE === "TROQUE-ESTA-CHAVE" || k !== CHAVE) throw new Error("Chave de acesso inválida."); }

function aba_(k) {
  const d = ABAS[k], ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(d.nome);
  if (!sh) {
    sh = ss.insertSheet(d.nome);
    sh.getRange(1, 1, 1, d.cols.length).setValues([d.cols]).setFontWeight("bold");
    sh.setFrozenRows(1);
    d.cols.forEach((c, i) => { if (d.texto.indexOf(c) >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat("@"); });
  }
  return sh;
}
function dados_(k) {
  const sh = aba_(k), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, ABAS[k].cols.length).getValues();
}
function gravar_(k, linhas) {
  const sh = aba_(k), nc = ABAS[k].cols.length, n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, nc).clearContent();
  if (linhas.length) {
    if (sh.getMaxRows() < linhas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), linhas.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, linhas.length, nc).setValues(linhas);
  }
}
const num_ = v => (v === "" || v === null || v === undefined || isNaN(Number(v))) ? "" : Number(v);
const txt_ = v => (v === null || v === undefined) ? "" : String(v);

function rev_() {
  const sh = aba_("cfg"), v = dados_("cfg").find(r => r[0] === "rev");
  return v ? String(v[1]) : "0";
}
function novaRev_() {
  const sh = aba_("cfg"), linhas = dados_("cfg").filter(r => r[0] !== "rev"), r = String(Date.now());
  linhas.push(["rev", r]); gravar_("cfg", linhas); return r;
}

function lerTudo_() {
  const prod = dados_("prod").map(r => ({ id: r[0], nome: r[1], cat: r[2], un: r[3], cod: num_(r[4]), preco: num_(r[5]), ativo: r[6] || "SIM", desc: r[7] || "" }));
  const itens = {};
  dados_("itens").forEach(r => { (itens[r[0]] = itens[r[0]] || []).push({ pid: r[1], qtd: Number(r[3]) || 0, etq: txt_(r[4]) }); });
  const ped = dados_("ped").map(r => { let o = {}; try { o = JSON.parse(r[12]); } catch (e) { o = { id: r[0], tipo: r[1], cli: r[2], hora: r[3], st: r[4] }; } o.itens = itens[o.id] || []; return o; });
  const est = {}; dados_("est").forEach(r => { est[r[0]] = r[1]; });
  const end = dados_("end").map(r => ({ nome: r[0], faixa: r[1], bairro: r[2], taxa: Number(r[3]) || 0 }));
  return { ok: true, rev: rev_(), prod, ped, est, end };
}

function sincronizar_(b) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    if (b.prod) gravar_("prod", b.prod.map(p => [p.id, p.nome, p.cat, p.un, num_(p.cod), num_(p.preco), p.ativo || "SIM", p.desc || ""]));
    if (b.est)  gravar_("est", Object.keys(b.est).map(id => [id, num_(b.est[id])]));
    if (b.end)  gravar_("end", b.end.map(e => [e.nome, e.faixa || "", e.bairro || "", Number(e.taxa) || 0]));
    if (b.ped) {
      const del = {}, up = {};
      (b.ped.del || []).forEach(id => del[id] = 1);
      (b.ped.up || []).forEach(o => up[o.id] = o);
      const linhas = dados_("ped").filter(r => !del[r[0]] && !up[r[0]]);
      const itens = dados_("itens").filter(r => !del[r[0]] && !up[r[0]]);
      const nomes = {}; dados_("prod").forEach(r => nomes[r[0]] = r[1]);
      Object.keys(up).forEach(id => {
        const o = up[id], copia = Object.assign({}, o); delete copia.itens;
        linhas.push([o.id, o.tipo, o.cli, o.hora, o.st, o.end || "", o.num || "", o.comp || "", o.bairro || "", o.pag || "", Number(o.taxa) || 0, o.extra ? "SIM" : "", JSON.stringify(copia)]);
        (o.itens || []).forEach(i => itens.push([o.id, i.pid, nomes[i.pid] || "", Number(i.qtd) || 0, txt_(i.etq)]));
      });
      linhas.sort((a, c) => String(a[0]).localeCompare(String(c[0])));
      gravar_("ped", linhas); gravar_("itens", itens);
    }
    return { ok: true, rev: novaRev_() };
  } finally { lock.releaseLock(); }
}
