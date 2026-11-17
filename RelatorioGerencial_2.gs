/* ============================================================
   RELATÓRIO SEMANAL — CRM Dellamed
   ------------------------------------------------------------
   Toda segunda-feira, por volta das 8h, o CRM:
     1. tira uma "foto" dos dados (a mesma leitura que o CRM faz ao abrir);
     2. gera um link novo e exclusivo para essa foto;
     3. manda o link por e-mail para RG_CONFIG.destinatario.

   O relatório tem o mesmo formato de antes ("CRM LEADS — Visão Semanal", arquivo
   relatorio.html), mas os dados vêm da foto do CRM, com as mesmas regras do CRM
   (status, data do status, valor, leads ocultos). Os números NÃO mudam depois:
   quem abrir o link na quinta vê o mesmo que na segunda.
   Cada semana tem o seu link; os links antigos continuam abrindo a semana deles
   até serem apagados (ver manterSemanas).

   SEGURANÇA / O QUE CADA LINK CARREGA:
     • Semana de referência = última semana completa (segunda a domingo) antes do envio.
       Se essa semana atravessa a virada do mês, saem DOIS links no mesmo e-mail
       (ex.: 28 a 30/09 no relatório de setembro e 01 a 04/10 no de outubro),
       assim nenhum dia fica sem relatório e o mês de cada link fica travado.
     • Dados identificáveis (razão social, CNPJ, UF, vendedor, motivo da perda) vão
       SÓ para os leads que tiveram movimento na semana de referência.
     • Os demais leads do mês (e da semana anterior, para o comparativo) vão sem
       nenhuma identificação: só time, status, origem, datas e valor — o mínimo para
       os KPIs mensais e o comparativo darem o número certo.
     • Leads fora dessa janela não vão para o link.

   ARQUIVOS: este RelatorioGerencial.gs + o arquivo HTML "relatorio" (relatorio.html).

   INSTALAR (uma vez, pelo editor — botão Executar):
     rg_instalarEnvioSemanal()   → liga o envio de toda segunda às 8h
   TESTAR / GERAR NA HORA:
     rg_enviarAgora()            → gera um relatório agora e manda o e-mail
   DESLIGAR:
     rg_desligarEnvioAutomatico()
   ============================================================ */

var RG_CONFIG = {
  destinatario: 'bianca.lima@dellamed.com.br',
  // Link /exec do CRM (o mesmo que a página do GitHub abre)
  crmUrl: 'https://script.google.com/macros/s/AKfycbwn43yGwRyB5TMq8G1q8WcQJj8dJx5R4eNOmEtmsKT2XDFVPyYqzdnGTGxApT0uAJCGPg/exec',
  manterSemanas: 12,          // quantos relatórios antigos continuam abrindo
  pasta: 'CRM Dellamed - Relatorios semanais',
  timesOcultos: ['Time Invictus'],   // times que não aparecem no relatório (igual ao relatório antigo)
  fuso: 'America/Sao_Paulo'
};

/* ---------- Rodar 1x: liga o envio de toda segunda às 8h ---------- */
function rg_instalarEnvioSemanal() {
  _soEditor_();
  _rg_removerGatilhos_();
  ScriptApp.newTrigger('rg_gerarRelatorioSemanal')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).nearMinute(0)
    .inTimezone(RG_CONFIG.fuso).create();
  Logger.log('Envio semanal ligado: toda segunda por volta das 8h para ' + RG_CONFIG.destinatario + '.');
}

/* ---------- Gera e envia agora (teste ou semana extra) ---------- */
function rg_enviarAgora() {
  _soEditor_();
  return _rg_gerarEEnviar_();
}

/* ---------- Desliga o envio automático ---------- */
function rg_desligarEnvioAutomatico() {
  _soEditor_();
  var n = _rg_removerGatilhos_();
  Logger.log('Envio automático desligado. Gatilhos removidos: ' + n + '. Os links já enviados continuam abrindo.');
}

/* ---------- Chamado pelo gatilho de segunda ----------
   Segurança: só roda se vier do gatilho deste projeto (ou pelo editor).
   Ninguém consegue disparar e-mails chamando esta função pelo navegador. */
function rg_gerarRelatorioSemanal(e) {
  var doGatilho = !!(e && e.triggerUid) && ScriptApp.getProjectTriggers().some(function (t) {
    return t.getUniqueId() === String(e.triggerUid);
  });
  if (!doGatilho) _soEditor_();
  return _rg_gerarEEnviar_();
}
// nome antigo (mensal) — mesmo comportamento, por compatibilidade
function rg_gerarRelatorioMensal(e) { return rg_gerarRelatorioSemanal(e); }

/* ---------- Mostra no log o link do último relatório ---------- */
function rg_link() {
  _soEditor_();
  var lista = _rg_lista_();
  var u = lista.length ? lista[lista.length - 1] : null;
  Logger.log(u ? ('Último relatório (' + (u.semana || u.geradoEm) + '): ' + _rg_url_(u.id)) : 'Nenhum relatório gerado ainda.');
  return u ? _rg_url_(u.id) : '';
}

// ════════════════════ interno ════════════════════

function _rg_gerarEEnviar_() {
  var agora = new Date();
  var hoje = Utilities.formatDate(agora, RG_CONFIG.fuso, 'yyyy-MM-dd');
  var todos = _rg_leadsBase_();
  var lista = _rg_lista_();
  var links = [];
  _rg_periodos_(hoje).forEach(function (per) {
    var foto = _rg_foto_(agora, per, todos);
    var id = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    var arq = _rg_pastaDrive_().createFile('relatorio_semana_' + per.ini + '_' + id.slice(0, 6) + '.json',
                                           JSON.stringify(foto), 'application/json');
    lista.push({ id: id, fileId: arq.getId(), geradoEm: foto.geradoEm, semana: per.rotulo });
    links.push({ url: _rg_url_(id), rotulo: per.rotulo });
  });
  // limpa os mais antigos
  while (lista.length > RG_CONFIG.manterSemanas) {
    var velho = lista.shift();
    try { DriveApp.getFileById(velho.fileId).setTrashed(true); } catch (_) {}
  }
  PropertiesService.getScriptProperties().setProperty('rg_snaps', JSON.stringify(lista));

  var NAVY = '#002f63', RED = '#A6192E', assunto = 'Relatório CRM de Leads Semanal';
  var botoes = links.map(function (l) {
    return '<p style="font-size:13px;color:#002f63;font-weight:700;margin:0 0 8px">Semana: ' + l.rotulo + '</p>' +
      '<a href="' + l.url + '" style="display:inline-block;background:' + RED + ';color:#fff;text-decoration:none;font-weight:700;padding:12px 28px;border-radius:6px;font-size:14px;margin-bottom:18px">ABRIR RELATÓRIO</a>';
  }).join('');
  var aviso = links.length > 1
    ? '<p style="font-size:12px;color:#666;margin:0 0 18px">A semana passada atravessou a virada do mês, por isso vão dois relatórios: um para cada mês.</p>' : '';
  MailApp.sendEmail({
    to: RG_CONFIG.destinatario,
    subject: assunto + ' · ' + links.map(function (l) { return l.rotulo; }).join(' + '),
    name: 'CRM Dellamed — Relatório Gerencial',
    body: links.map(function (l) { return 'Semana ' + l.rotulo + ': ' + l.url; }).join('\n'),
    htmlBody:
      '<div style="font-family:Montserrat,Arial,sans-serif;max-width:520px;margin:0 auto;color:#222">' +
      '<div style="background:' + NAVY + ';padding:20px 24px;border-radius:8px 8px 0 0">' +
      '<h1 style="color:#fff;margin:0;font-size:19px">' + assunto + '</h1></div>' +
      '<div style="border:1px solid #eee;border-top:none;border-radius:0 0 8px 8px;padding:26px 24px;text-align:center">' +
      '<p style="font-size:14px;color:#333;margin:0 0 20px">Clique no link para ter visão completa semanal do CRM LEADS.</p>' +
      aviso + botoes +
      '<p style="font-size:12px;color:#666;margin:4px 0 0">Dados congelados em ' + Utilities.formatDate(agora, RG_CONFIG.fuso, 'dd/MM/yyyy HH:mm') + ' — cada link mostra só a semana dele.</p>' +
      '<p style="font-size:11px;color:#999;margin-top:14px;word-break:break-all">Se o botão não abrir, copie e cole:<br>' +
      links.map(function (l) { return l.rotulo + ': ' + l.url; }).join('<br>') + '</p>' +
      '</div></div>'
  });
  Logger.log('Relatório(s) gerado(s) e enviado(s) para ' + RG_CONFIG.destinatario + ':\n' + links.map(function (l) { return l.rotulo + ' → ' + l.url; }).join('\n'));
  return links.map(function (l) { return l.url; }).join('\n');
}

// Semana de referência = última semana completa (seg a dom) antes de "hoje".
// Se atravessa a virada do mês, vira dois períodos (um por mês).
function _rg_periodos_(hoje) {
  var fim = _rg_addDias_(hoje, -_rg_dow_(hoje));     // segunda desta semana (exclusivo)
  var ini = _rg_addDias_(fim, -7);                    // segunda da semana passada
  var ultimo = _rg_addDias_(fim, -1);
  var pers = [];
  if (ini.slice(0, 7) === ultimo.slice(0, 7)) pers.push({ ini: ini, fim: fim, recortada: false });
  else {
    var vira = ultimo.slice(0, 7) + '-01';
    pers.push({ ini: ini, fim: vira, recortada: true });
    pers.push({ ini: vira, fim: fim, recortada: true });
  }
  return pers.map(function (p) {
    var mesIni = p.ini.slice(0, 7) + '-01';
    var mesFim = _rg_proxMes_(mesIni);
    var antIni = _rg_addDias_(p.ini, -7);
    return {
      ini: p.ini, fim: p.fim, recortada: p.recortada,
      mesIni: mesIni, mesFim: mesFim,
      mesFechado: p.fim === mesFim,                    // a semana fecha o mês
      janelaIni: antIni < mesIni ? antIni : mesIni,    // o que o relatório precisa enxergar
      rotulo: _rg_br_(p.ini) + ' a ' + _rg_br_(_rg_addDias_(p.fim, -1))
    };
  });
}
function _rg_addDias_(iso, n) {
  var p = iso.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)).toISOString().slice(0, 10);
}
function _rg_dow_(iso) {   // segunda = 0 ... domingo = 6
  var p = iso.split('-');
  return (new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay() + 6) % 7;
}
function _rg_proxMes_(mesIni) {
  var p = mesIni.split('-');
  return new Date(Date.UTC(+p[0], +p[1], 1)).toISOString().slice(0, 10);
}
function _rg_br_(iso) { var p = iso.split('-'); return p[2] + '/' + p[1]; }

// Leitura base: mesma leitura do CRM (getData_) e as MESMAS regras do CRM para status,
// data do status, valor e leads ocultos. Fica só na memória do servidor.
function _rg_leadsBase_() {
  var cfg = getConfig_();
  var teams = (cfg.timeComercial || []).map(function (t) { return String(t).trim(); }).filter(String);
  var ocultos = RG_CONFIG.timesOcultos || [];
  var leads = [];
  (getData_('LEADS').leads || []).forEach(function (L) {
    if (String(L['Oculto'] || '') === '1') return;                       // oculto: fora, como no CRM
    var st = _rg_normSt_(L['Status']);
    var data = _rg_dataISO_(L['Data']);
    var at = [];
    try { at = JSON.parse(L['Atividades'] || '[]'); } catch (_) { at = []; }
    if (!Array.isArray(at)) at = [];
    // Histórico: entrada como Aberto + cada mudança de status registrada ("... → X")
    var hist = data ? [{ status: 'Aberto', date: data }] : [];
    at.filter(function (a) { return a && a.type === 'status'; }).forEach(function (a) {
      var m = /→\s*(.+?)\s*$/.exec(String(a.title || ''));
      var d = _rg_dataISO_(a.date);
      if (m && d) hist.push({ status: _rg_normSt_(m[1]), date: d });
    });
    hist.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    var t = String(L['Time Comercial'] || '').trim();
    if (ocultos.indexOf(t) >= 0) return;
    if (!t) t = 'Sem time definido';
    else if (teams.length && teams.indexOf(t) < 0) t = t + ' (fora da lista configurada)';
    leads.push({
      time: t,
      vendedor: String(L['Prospect'] || '').trim() || '(sem proprietário)',
      status: st,
      canal: String(L['Origem'] || '').trim() || 'Não informado',
      dataCadastro: data,
      dataStatus: _rg_stageDate_(L, st, at) || data,
      historico: hist,
      valor: st === 'Convertido' ? _rg_valorPedido_(L) : 0,
      motivoPerda: st === 'Perdido' ? (String(L['Motivo da Perda'] || '').trim() || 'Não informado') : null,
      razaoSocial: String(L['Razão Social'] || L['Lead'] || '').toUpperCase(),
      cnpj: String(L['CNPJ'] || L['Registro'] || ''),
      uf: String(L['UF'] || '').trim().toUpperCase()
    });
  });
  return { teams: teams.filter(function (x) { return ocultos.indexOf(x) < 0; }), ocultos: ocultos, leads: leads };
}

// Foto de UM período: só o que esse link precisa (ver "SEGURANÇA" no topo).
function _rg_foto_(agora, per, base) {
  var noInt = function (d, a, b) { return !!d && d >= a && d < b; };
  var leads = [];
  base.leads.forEach(function (l) {
    var datas = [l.dataCadastro, l.dataStatus].concat(l.historico.map(function (h) { return h.date; }));
    if (!datas.some(function (d) { return noInt(d, per.janelaIni, per.fim); })) return;   // fora da janela: não vai
    var movSemana = l.historico.filter(function (h) { return h.status !== 'Aberto' && noInt(h.date, per.ini, per.fim); });
    var daSemana = noInt(l.dataCadastro, per.ini, per.fim) || movSemana.length > 0 ||
      ((l.status === 'Convertido' || l.status === 'Perdido') && noInt(l.dataStatus, per.ini, per.fim));
    var item = {
      time: l.time, status: l.status, canal: l.canal,
      dataCadastro: l.dataCadastro, dataStatus: l.dataStatus,
      valor: noInt(l.dataStatus, per.janelaIni, per.fim) ? l.valor : 0
    };
    if (daSemana) {               // identificação só para quem se movimentou na semana do link
      item.vendedor = l.vendedor; item.razaoSocial = l.razaoSocial; item.cnpj = l.cnpj; item.uf = l.uf;
      item.motivoPerda = l.motivoPerda; item.historico = movSemana;
    } else {
      item.vendedor = ''; item.razaoSocial = ''; item.cnpj = ''; item.uf = ''; item.motivoPerda = null; item.historico = [];
    }
    leads.push(item);
  });
  return {
    geradoEm: Utilities.formatDate(agora, RG_CONFIG.fuso, 'dd/MM/yyyy HH:mm'),
    hoje: _rg_addDias_(per.fim, -1),
    periodo: { ini: per.ini, fim: per.fim, mesIni: per.mesIni, mesFim: per.mesFim, mesFechado: per.mesFechado, recortada: per.recortada },
    teams: base.teams,
    ocultos: base.ocultos,
    leads: leads
  };
}

// Valor do lead = mesma regra do CRM: coluna Valor do Pedido; se vazia/zerada, soma da lista de Pedidos
function _rg_valorPedido_(L) {
  var p = [];
  try { p = JSON.parse(L['Pedidos'] || '[]'); } catch (_) { p = []; }
  var s = Array.isArray(p) ? p.reduce(function (a, x) { return a + _rg_brlNum_(x && (x.valor != null ? x.valor : x.Valor)); }, 0) : 0;
  s = Math.round(s * 100) / 100;
  if (s > 0) return s;
  return Math.round(_rg_brlNum_(L['Valor do Pedido']) * 100) / 100;
}
// ── mesmas regras do CRM (a_core.js) ──
function _rg_normSt_(s) {
  var k = String(s || '').trim().toLowerCase();
  if (!k) return 'Aberto';
  var velho = { 'novo': 'Aberto', 'contato feito': 'Qualificado', 'ganho': 'Convertido' };
  if (velho[k]) return velho[k];
  var ok = ['Aberto', 'Qualificado', 'Proposta enviada', 'Negociando', 'Convertido', 'Perdido', 'Inativo'];
  for (var i = 0; i < ok.length; i++) if (ok[i].toLowerCase() === k) return ok[i];
  return String(s).trim();
}
// Data do status atual = igual ao stageDate do CRM
function _rg_stageDate_(L, st, at) {
  if (!st || st === 'Aberto') return _rg_dataISO_(L['Data']);
  var re = new RegExp('→\\s*' + st.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$', 'i'), f = '';
  at.forEach(function (a) { if (a && a.type === 'status' && re.test(a.title || '')) f = a.date; });
  return _rg_dataISO_(f) || _rg_dataISO_(L['Data Último Contato']) || _rg_dataISO_(L['Data']);
}
function _rg_dataISO_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, RG_CONFIG.fuso, 'yyyy-MM-dd');
  var s = String(v).trim();
  if (s.indexOf('/') >= 0) {
    var p = s.split(' ')[0].split('/');
    if (p.length !== 3) return '';
    var d = new Date(+p[2], +p[1] - 1, +p[0]);
    if (isNaN(d)) return '';
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
}
function _rg_brlNum_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v == null ? '' : v).replace(/[^\d.,-]/g, '');
  if (!s) return 0;
  // mesma regra do CRM: último separador com 1-2 dígitos depois = decimal; o resto é milhar
  var neg = s.charAt(0) === '-'; s = s.replace(/-/g, '');
  var i = Math.max(s.lastIndexOf(','), s.lastIndexOf('.')), dec = s.length - i - 1, n;
  if (i > -1 && dec >= 1 && dec <= 2) n = parseFloat(s.slice(0, i).replace(/[.,]/g, '') + '.' + s.slice(i + 1));
  else n = parseFloat(s.replace(/[.,]/g, ''));
  n = n || 0; return neg ? -n : n;
}

// Página do relatório (chamada pelo doGet quando o link tem ?relatorio=...)
function rg_pagina_(id) {
  id = String(id || '');
  var item = /^[a-f0-9]{40}$/.test(id) ? _rg_lista_().filter(function (x) { return x.id === id; })[0] : null;
  if (!item) return _rg_aviso_('Link inválido ou expirado', 'Este relatório não existe mais. Peça o link da semana atual.');
  // Relatórios gerados antes desta versão traziam a base inteira: não abrem mais
  // (para cancelar todos de uma vez, o arquivo antigo também pode ir para a lixeira).
  if (!item.semana) return _rg_aviso_('Link antigo desativado', 'Este link é de uma versão anterior do relatório e foi desativado por segurança.');
  var json;
  try {
    var f = DriveApp.getFileById(item.fileId);
    // Para cancelar um link: mande o arquivo dessa semana para a lixeira do Drive
    if (f.isTrashed()) return _rg_aviso_('Link cancelado', 'Este relatório foi retirado do ar.');
    json = f.getBlob().getDataAsString('UTF-8');
  }
  catch (_) { return _rg_aviso_('Relatório indisponível', 'Não foi possível abrir este relatório.'); }
  var foto = JSON.parse(json);
  var js = function (o) {   // JSON seguro dentro de <script>
    return JSON.stringify(o).replace(/</g, '\\u003c').replace(new RegExp(String.fromCharCode(0x2028), 'g'), '\\u2028').replace(new RegExp(String.fromCharCode(0x2029), 'g'), '\\u2029');
  };
  var html = HtmlService.createHtmlOutputFromFile('relatorio').getContent()
    .replace('__RG_LEADS__', function () { return js(foto.leads); })
    .replace('__RG_TEAMS__', function () { return js(foto.teams); })
    .replace('__RG_OCULTOS__', function () { return js(foto.ocultos); })
    .replace('__RG_HOJE__', function () { var p = foto.hoje.split('-'); return p[0] + ',' + (+p[1] - 1) + ',' + (+p[2]); })
    .replace('__RG_GERADO__', function () { return js(foto.geradoEm); })
    .replace('__RG_PERIODO__', function () { return js(foto.periodo); });
  return HtmlService.createHtmlOutput(html)
    .setTitle('CRM LEADS — Visão Gerencial')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    ;
}

function _rg_aviso_(t, m) {
  return HtmlService.createHtmlOutput('<div style="font-family:Arial,sans-serif;max-width:460px;margin:80px auto;text-align:center;color:#32363a">' +
    '<h2 style="color:#0F3C6E">' + t + '</h2><p>' + m + '</p></div>').setTitle('CRM Dellamed');
}
function _rg_url_(id) { return RG_CONFIG.crmUrl + '?relatorio=' + id; }
function _rg_lista_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('rg_snaps') || '[]'); } catch (_) { return []; }
}
function _rg_pastaDrive_() {
  var it = DriveApp.getFoldersByName(RG_CONFIG.pasta);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RG_CONFIG.pasta);
}
function _rg_removerGatilhos_() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'rg_gerarRelatorioSemanal' || fn === 'rg_gerarRelatorioMensal') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}
