// ═══════════════════════════════════════════════════════════════════
//  CRM PIPELINE DELLAMED — Apps Script Backend (Codigo.gs)
//  VERSÃO SEGURA — a página é servida pelo próprio Apps Script.
//
//  O que mudou em relação à versão anterior:
//  • Não existe mais API aberta (doGet/doPost em JSON). A tela conversa
//    com o servidor por google.script.run, sem nenhuma URL no código-fonte.
//  • Login com e-mail + senha. A senha é criada pelo próprio usuário com um
//    código enviado ao e-mail corporativo (prova que o e-mail é dele).
//    Senhas ficam guardadas só como hash (colunas ocultas em USUARIOS).
//  • Sessão com token temporário (expira após 6h sem uso ou 12h no total).
//  • O servidor decide quem é o usuário pela sessão — nunca pelo e-mail
//    que o navegador manda. Ações de admin são checadas no servidor.
//  • Todas as funções internas terminam com "_" (o Apps Script bloqueia
//    chamadas do navegador a funções com esse final).
//  • PIN de administrador removido (o admin também entra com senha).
//
//  Abas necessárias na planilha:
//    LEADS      — dados dos leads
//    WHATSAPP   — leads via WhatsApp
//    PAINEL     — opções de config (motivo perda, canal, linha, time, status)
//    USUARIOS   — Email | Nome | Status | DataCadastro | Admin | ÚltimoAcesso | SenhaHash | SenhaSalt
//    ACESSOS    — Nome | Email  (quem pode SOLICITAR acesso; nome = corporativo exato)
//    PROSPECTS  — Nome | Email | Time
//    ALTERACOES — log de transferências e atribuições
// ═══════════════════════════════════════════════════════════════════

const ADMIN_EMAIL = 'bianca.lima@dellamed.com.br';
const SS = SpreadsheetApp.getActiveSpreadsheet();

const SESSAO_INATIVIDADE_S = 6 * 60 * 60;          // 6h sem uso → sai
const SESSAO_MAXIMA_MS     = 12 * 60 * 60 * 1000;   // 12h no máximo
const HASH_ROUNDS          = 500;

// ─── PÁGINA ────────────────────────────────────────────────────────

function doGet(e) {
  // Link do relatório semanal (?relatorio=<id>) → abre o relatório congelado, sem login
  var rel = e && e.parameter && e.parameter.relatorio;
  if (rel) return rg_pagina_(rel);

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CRM Dellamed')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // Permite abrir o CRM dentro da página do GitHub (o link antigo continua valendo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── FUNÇÕES CHAMADAS PELA TELA (públicas) ─────────────────────────
// Só estas funções podem ser chamadas pelo navegador. Todas validam tudo.

function crmLogin(email, senha) {
  try {
    email = _normEmail_(email);
    if (!email || !senha) return { error: 'Preencha e-mail e senha.' };
    _checaLimite_('login_' + email, 5, 'Muitas tentativas. Aguarde 15 minutos e tente de novo.');

    const reg = _usuarioRegistro_(email);
    if (reg && reg.status === 'pendente') {
      return { pending: true, message: 'Seu acesso está pendente de aprovação. Aguarde a liberação.' };
    }
    if (reg && reg.status === 'aprovado' && !reg.hash) {
      return { error: 'Você ainda não criou sua senha. Clique em "Primeiro acesso ou esqueci a senha".' };
    }
    if (!reg || reg.status !== 'aprovado' || _hashSenha_(senha, reg.salt) !== reg.hash) {
      _somaLimite_('login_' + email, 900);
      Utilities.sleep(600);
      return { error: 'E-mail ou senha incorretos.' };
    }
    _zeraLimite_('login_' + email);

    _sheet_('USUARIOS').getRange(reg.linha, 6).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));

    const token = Utilities.getUuid() + Utilities.getUuid();
    const sessao = { email: reg.email, salt: reg.salt, inicio: Date.now() };
    CacheService.getScriptCache().put('sess_' + token, JSON.stringify(sessao), SESSAO_INATIVIDADE_S);
    return { ok: true, token: token, email: reg.email, nome: reg.nome, admin: reg.admin };
  } catch (err) {
    return { error: err.message };
  }
}

function crmMe(token) {
  try {
    const s = _sessao_(token);
    return { ok: true, email: s.user.email, nome: s.user.nome, admin: s.user.admin };
  } catch (err) {
    return { error: err.message, sessaoExpirada: true };
  }
}

function crmLogout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return { ok: true };
}

// Pedido de acesso de quem ainda não é usuário (mesma regra de antes: aba ACESSOS)
function crmSolicitarAcesso(nome, email) {
  try {
    email = _normEmail_(email);
    const nomeDigitado = String(nome || '').trim();
    if (!nomeDigitado) return { error: 'Preencha seu nome.' };
    if (!email) return { error: 'Preencha o e-mail.' };
    _checaLimite_('solic_' + email, 3, 'Muitas solicitações. Aguarde alguns minutos.');
    _somaLimite_('solic_' + email, 900);

    const reg = _usuarioRegistro_(email);
    if (reg) {
      if (reg.status === 'aprovado') return { ok: true, message: 'Você já tem acesso. Use "Primeiro acesso ou esqueci a senha" para criar sua senha.' };
      if (reg.status === 'pendente') return { pending: true, message: 'Seu acesso está pendente de aprovação. Aguarde a liberação.' };
      return { error: 'Seu acesso foi negado. Entre em contato com a administradora.' };
    }

    const wsA = SS.getSheetByName('ACESSOS');
    if (!wsA) return { error: 'Lista de acessos não configurada. Contate a administradora.' };
    const dataA = wsA.getDataRange().getValues();
    let entrada = null;
    for (let i = 1; i < dataA.length; i++) {
      if (String(dataA[i][1] || '').trim().toLowerCase() === email) {
        entrada = { nome: String(dataA[i][0] || '').trim(), email: email };
        break;
      }
    }
    // Mensagem única para e-mail não autorizado OU nome diferente (não revela quem está na lista)
    if (!entrada || nomeDigitado !== entrada.nome) {
      return { error: 'Não foi possível validar nome e e-mail. Use o nome exato do cadastro corporativo ou fale com a administradora.' };
    }

    _sheet_('USUARIOS').appendRow([email, entrada.nome, 'pendente', new Date(), '']);
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: '[CRM Dellamed] Solicitação de acesso — ' + entrada.nome,
        htmlBody:
          '<h2>Nova solicitação de acesso ao CRM</h2>' +
          '<p><b>Nome:</b> ' + entrada.nome + '</p>' +
          '<p><b>E-mail:</b> ' + email + '</p>' +
          '<p><b>Data:</b> ' + new Date().toLocaleString('pt-BR') + '</p>' +
          '<p>Acesse o CRM para aprovar na aba Usuários.</p>'
      });
    } catch (mailErr) {
      Logger.log('ERRO ao enviar e-mail de solicitação: ' + mailErr.message);
    }
    return { pending: true, message: 'Solicitação enviada! Depois da aprovação, use "Primeiro acesso" para criar sua senha.' };
  } catch (err) {
    return { error: err.message };
  }
}

// Envia código de 6 dígitos para o e-mail (primeiro acesso ou esqueci a senha)
function crmEnviarCodigo(email) {
  const resposta = { ok: true, message: 'Se este e-mail tiver acesso aprovado, você vai receber um código em instantes. Confira também o spam.' };
  try {
    email = _normEmail_(email);
    if (!email) return { error: 'Preencha o e-mail.' };
    _checaLimite_('cod_env_' + email, 3, 'Muitos códigos pedidos. Aguarde 15 minutos.');
    _somaLimite_('cod_env_' + email, 900);

    const reg = _usuarioRegistro_(email);
    if (!reg || reg.status !== 'aprovado') return resposta; // não revela se o e-mail existe

    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    let n = 0;
    for (let i = 0; i < 6; i++) n = (n * 256 + (bytes[i] & 0xff)) % 1000000;
    const codigo = ('000000' + n).slice(-6);

    CacheService.getScriptCache().put('cod_' + email,
      JSON.stringify({ h: _hashSenha_(codigo, email), t: 0 }), 900);

    MailApp.sendEmail({
      to: email,
      subject: '[CRM Dellamed] Seu código de acesso: ' + codigo,
      htmlBody:
        '<div style="font-family:Montserrat,Arial,sans-serif;max-width:460px">' +
        '<h2 style="color:#002f63">Código de acesso ao CRM</h2>' +
        '<p>Use este código para criar ou redefinir sua senha:</p>' +
        '<p style="font-size:28px;font-weight:700;letter-spacing:6px;color:#002f63">' + codigo + '</p>' +
        '<p style="color:#666;font-size:13px">Vale por 15 minutos. Se você não pediu, ignore este e-mail.</p></div>'
    });
    return resposta;
  } catch (err) {
    return { error: err.message };
  }
}

function crmDefinirSenha(email, codigo, novaSenha) {
  try {
    email = _normEmail_(email);
    if (String(novaSenha || '').length < 8) return { error: 'A senha precisa ter pelo menos 8 caracteres.' };
    const cache = CacheService.getScriptCache();
    const bruto = cache.get('cod_' + email);
    if (!bruto) return { error: 'Código expirado ou inválido. Peça um novo código.' };
    const c = JSON.parse(bruto);
    if (c.t >= 5) { cache.remove('cod_' + email); return { error: 'Muitas tentativas. Peça um novo código.' }; }
    if (_hashSenha_(String(codigo || '').trim(), email) !== c.h) {
      c.t++;
      cache.put('cod_' + email, JSON.stringify(c), 900);
      return { error: 'Código incorreto.' };
    }
    const reg = _usuarioRegistro_(email);
    if (!reg || reg.status !== 'aprovado') return { error: 'Usuário sem acesso aprovado.' };
    _gravarSenha_(reg.linha, novaSenha);
    cache.remove('cod_' + email);
    _zeraLimite_('login_' + email);
    return { ok: true, message: 'Senha salva! Agora entre com seu e-mail e a nova senha.' };
  } catch (err) {
    return { error: err.message };
  }
}

// Ponto único de entrada de dados. A identidade vem SEMPRE da sessão.
function crmApi(token, body) {
  try {
    const s = _sessao_(token);
    const u = s.user;
    body = body || {};
    body.userEmail = u.email;
    body.requesterEmail = u.email;
    body.userNome = u.nome;
    body.fromUser = u.nome;

    const soAdmin = function () {
      if (!u.admin) throw new Error('Apenas administradores podem executar esta ação.');
    };

    switch (body.action || '') {
      // Leitura
      case 'getData':            return getData_();
      case 'getWAData':          return SS.getSheetByName('WHATSAPP') ? getData_('WHATSAPP') : { leads: [] };
      case 'getRev':             return getRev_();
      case 'getConfig':          return getConfig_();
      case 'getLocalidades':     return getLocalidades_(body.pais);   // Localidades.gs
      case 'getProspects':       return getProspects_();
      case 'getAuthorized':      return { emails: (_isAuthorizedEmail_(u.email) && !_isSupervisor_(u.email)) ? [u.email] : [], supervisor: _isSupervisor_(u.email) };
      case 'getUsers':           soAdmin(); return getUsers_();
      case 'getTransfers':       soAdmin(); return getTransfers_();
      case 'getNotifications':   return getNotifications_(u.email);
      case 'lookupUser':         return lookupUser_(body.email);
      // Escrita
      case 'addLead':            return addLead_(body);
      case 'editLead':           return editLead_(body);
      case 'deleteLead':         return deleteLead_(body);
      case 'addWA':              return addLead_(body);          // tudo vai para LEADS (base unificada)
      case 'editWA':             return editLead_(body, 'WHATSAPP');
      case 'deleteWA':           return deleteLead_(body, 'WHATSAPP');
      case 'approveUser':        return approveUser_(body);
      case 'rejectUser':         return rejectUser_(body);
      case 'deleteUser':         return deleteUser_(body);
      case 'requestTransfer':    return requestTransfer_(body);
      case 'approveTransfer':    return approveTransfer_(body);
      case 'rejectTransfer':     return rejectTransfer_(body);
      case 'clearNotifications': return clearNotifications_(body);
      case 'notifyAssign':       return notifyAssign_(body);
      case 'notifyActivity':     return notifyActivity_(body);
      case 'uploadFile':         return uploadFile_(body);
      case 'deleteFile':         return deleteFile_(body);
      default:                   return { error: 'Ação inválida' };
    }
  } catch (err) {
    if (err.message === 'SESSAO_EXPIRADA') return { error: 'Sua sessão expirou. Entre de novo.', sessaoExpirada: true };
    return { error: err.message };
  }
}

// ─── SEGURANÇA (interno) ───────────────────────────────────────────

function _normEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function _hashSenha_(senha, salt) {
  let h = String(salt) + '|' + String(senha);
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}

// Garante as colunas de senha (G e H) na aba USUARIOS, ocultas
function _garanteColunasSenha_(ws) {
  const cab = ws.getRange(1, 7, 1, 2).getValues()[0];
  if (cab[0] !== 'SenhaHash' || cab[1] !== 'SenhaSalt') {
    ws.getRange(1, 7, 1, 2).setValues([['SenhaHash', 'SenhaSalt']]);
    ws.hideColumns(7, 2);
  }
}

function _usuarioRegistro_(email) {
  const ws = _sheet_('USUARIOS');
  _garanteColunasSenha_(ws);
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      return {
        linha: i + 1,
        email: email,
        nome: String(data[i][1]).trim(),
        status: String(data[i][2]).trim().toLowerCase(),
        admin: email === ADMIN_EMAIL || String(data[i][4]).trim().toLowerCase() === 'sim',
        hash: String(data[i][6] || ''),
        salt: String(data[i][7] || '')
      };
    }
  }
  return null;
}

function _gravarSenha_(linha, senha) {
  const salt = Utilities.getUuid();
  _sheet_('USUARIOS').getRange(linha, 7, 1, 2).setValues([[_hashSenha_(senha, salt), salt]]);
}

// Valida o token e recarrega o usuário da planilha a cada chamada:
// se ele for bloqueado/excluído ou trocar a senha, a sessão cai na hora.
function _sessao_(token) {
  const cache = CacheService.getScriptCache();
  const bruto = token ? cache.get('sess_' + token) : null;
  if (!bruto) throw new Error('SESSAO_EXPIRADA');
  const s = JSON.parse(bruto);
  const reg = _usuarioRegistro_(s.email);
  if (!reg || reg.status !== 'aprovado' || reg.salt !== s.salt || Date.now() - s.inicio > SESSAO_MAXIMA_MS) {
    cache.remove('sess_' + token);
    throw new Error('SESSAO_EXPIRADA');
  }
  cache.put('sess_' + token, bruto, SESSAO_INATIVIDADE_S);
  return { user: reg };
}

function _checaLimite_(chave, max, msg) {
  if (Number(CacheService.getScriptCache().get('lim_' + chave) || 0) >= max) throw new Error(msg);
}
function _somaLimite_(chave, janelaS) {
  const c = CacheService.getScriptCache();
  c.put('lim_' + chave, String(Number(c.get('lim_' + chave) || 0) + 1), janelaS);
}
function _zeraLimite_(chave) {
  CacheService.getScriptCache().remove('lim_' + chave);
}

// Funções de manutenção só rodam pelo editor do Apps Script (nunca pela tela)
function _soEditor_() {
  const ativo = Session.getActiveUser().getEmail();
  if (!ativo || ativo !== Session.getEffectiveUser().getEmail()) {
    throw new Error('Execute esta função pelo editor do Apps Script.');
  }
}

// ─── USUÁRIOS (admin) ──────────────────────────────────────────────

function approveUser_(body) {
  _requireAdmin_(body.requesterEmail);
  const email = (body.email || '').trim().toLowerCase();
  const ws = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      ws.getRange(i + 1, 3).setValue('aprovado');
      try {
        MailApp.sendEmail({
          to: email,
          subject: '[CRM Dellamed] Acesso aprovado!',
          htmlBody: '<h2>Seu acesso ao CRM Dellamed foi aprovado!</h2><p>No primeiro acesso, clique em <b>"Primeiro acesso ou esqueci a senha"</b> para criar sua senha com um código enviado para este e-mail.</p>'
        });
      } catch (e1) {
        Logger.log('Email aprovação erro: ' + e1.message);
      }
      return { ok: true };
    }
  }
  return { error: 'Usuário não encontrado' };
}

function rejectUser_(body) {
  _requireAdmin_(body.requesterEmail);
  const email = (body.email || '').trim().toLowerCase();
  const ws = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      ws.getRange(i + 1, 3).setValue('rejeitado');
      return { ok: true };
    }
  }
  return { error: 'Usuário não encontrado' };
}

function deleteUser_(body) {
  _requireAdmin_(body.requesterEmail);
  const email = (body.email || '').trim().toLowerCase();
  if (email === ADMIN_EMAIL) return { error: 'Não é possível excluir a conta administradora.' };
  const ws = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      ws.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'Usuário não encontrado' };
}

// ─── PROSPECTS (dropdown de vendedores) ────────────────────────────
// Aba PROSPECTS: coluna A = Nome | coluna B = Email
// Adicione/remova linhas diretamente na planilha — reflete no CRM automaticamente.

function getProspects_() {
  const ws = SS.getSheetByName('PROSPECTS');
  if (!ws) return { prospects: [] };

  const data = ws.getDataRange().getValues();
  const prospects = [];

  for (let i = 1; i < data.length; i++) {
    const nome  = String(data[i][0] || '').trim();
    const email = String(data[i][1] || '').trim().toLowerCase();
    const time  = String(data[i][2] || '').trim(); // coluna C = Time Comercial do vendedor
    if (nome && email) {
      prospects.push({ nome: nome, email: email, time: time });
    }
  }

  return { prospects: prospects };
}

// ─── USUARIOS AUTORIZADOS (atribuição direta de vendedor) ──────────
// Aba: USUARIOS AUTORIZADOS — coluna A = Email
// Adicione/remova e-mails diretamente na planilha.

function getAuthorizedUsers_() {
  var ws = SS.getSheetByName('USUARIOS AUTORIZADOS');
  if (!ws) {
    // Criar aba com cabeçalho se não existir
    ws = SS.insertSheet('USUARIOS AUTORIZADOS');
    ws.appendRow(['Email']);
    ws.appendRow(['camily.soares@dellamed.com.br']);
    ws.appendRow(['fabio.rodrigues@dellamed.com.br']);
    ws.appendRow(['jessica.porto@dellamed.com.br']);
    ws.appendRow(['leonardo.dalpubel@dellamed.com.br']);
    ws.appendRow(['luisa.abreu@dellamed.com.br']);
    ws.appendRow(['marco.ribeiro@dellamed.com.br']);
    ws.appendRow(['patricia.borges@dellamed.com.br']);
    ws.appendRow(['sandriane.maciel@dellamed.com.br']);
    ws.appendRow(['sandro.junior@dellamed.com.br']);
    ws.appendRow(['felipe.fasoli@dellamed.com.br']);
    ws.appendRow(['rodrigo.velho@dellamed.com.br']);
  }

  var data = ws.getDataRange().getValues();
  var emails = [];
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][0] || '').trim().toLowerCase();
    if (email) emails.push(email);
  }
  return { emails: emails };
}

// ─── CRUD LEADS ────────────────────────────────────────────────────

function getData_(sheetName) {
  const ws = _sheet_(sheetName || 'LEADS');
  let data = ws.getDataRange().getValues();
  if (data.length < 2) return { leads: [], headers: data[0] || [] };
  const headers = data[0];

  const hdrClean = headers.map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase());
  let checkCol = hdrClean.indexOf('razão social');
  if (checkCol < 0) checkCol = hdrClean.indexOf('razao social');

  const statusCol = hdrClean.indexOf('status');
  const dpCol     = hdrClean.indexOf('data primeiro contato');
  let   duCol     = hdrClean.indexOf('data último contato');
  if (duCol < 0) duCol = hdrClean.indexOf('data ultimo contato');
  let   daCol     = hdrClean.indexOf('data próxima ação');
  if (daCol < 0) daCol = hdrClean.indexOf('data proxima ação');
  if (daCol < 0) daCol = hdrClean.indexOf('data proxima acao');
  const tempCol   = hdrClean.indexOf('temperatura lead');

  const today = new Date(); today.setHours(0, 0, 0, 0);

  // Auto-Inativo: LICITAÇÃO 90 dias / demais 30 dias.
  // Regra: ação próxima HOJE/FUTURA protege; senão conta a data passada mais recente
  // (inclui data de registro) contra o limite do canal.
  const canalCol = hdrClean.indexOf('canal principal');
  const dataCol  = hdrClean.indexOf('data');
  function _inatToDate(v){
    if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (v) { const p = String(v).split(' ')[0].split('/'); if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0])); }
    return null;
  }
  function _inatEhLicitacao(row){
    let c = canalCol >= 0 ? String(row[canalCol] || '') : '';
    try { c = c.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
    return c.toUpperCase().indexOf('LICITA') >= 0;
  }
  const rowsToUpdate = [];
  for (let i = 1; i < data.length; i++) {
    if (checkCol >= 0 && !data[i][checkCol]) continue;
    const st = statusCol >= 0 ? String(data[i][statusCol]).trim().toLowerCase() : '';
    // Só auto-inativa leads ainda NOVOS (ou sem status). Qualquer status definido
    // manualmente (Contato feito, Negociando, Proposta enviada, Ganho, Perdido, Inativo)
    // é preservado — evita que uma alteração manual volte sozinha para Inativo.
    if (st !== '' && st !== 'novo' && st !== 'aberto') continue;

    // Ação próxima hoje/futura => ativo (há ação planejada)
    const prox = daCol >= 0 ? _inatToDate(data[i][daCol]) : null;
    if (prox && prox.getTime() >= today.getTime()) continue;

    // Data passada mais recente (inclui registro)
    let refPast = null;
    [dataCol, dpCol, duCol, daCol].forEach(col => {
      if (col < 0) return;
      const d = _inatToDate(data[i][col]);
      if (d && d.getTime() <= today.getTime() && (!refPast || d > refPast)) refPast = d;
    });
    if (!refPast) continue;

    const lim = _inatEhLicitacao(data[i]) ? 90 : 30;
    if (Math.floor((today - refPast) / 86400000) >= lim) rowsToUpdate.push(i + 1);
  }

  if (rowsToUpdate.length > 0 && statusCol >= 0) {
    // Uma única escrita por coluna (RangeList) em vez de uma chamada por célula,
    // e atualiza o que já está em memória em vez de reler a planilha inteira.
    const colSt = _colLetra_(statusCol + 1);
    ws.getRangeList(rowsToUpdate.map(r => colSt + r)).setValue('Inativo');
    if (tempCol >= 0) {
      const colTp = _colLetra_(tempCol + 1);
      ws.getRangeList(rowsToUpdate.map(r => colTp + r)).setValue('Frio');
    }
    rowsToUpdate.forEach(r => {
      data[r - 1][statusCol] = 'Inativo';
      if (tempCol >= 0) data[r - 1][tempCol] = 'Frio';
    });
  }

  const hdr = _hdrLimpo_(headers), tz = Session.getScriptTimeZone();
  const leads = [];
  for (let i = 1; i < data.length; i++) {
    if (checkCol >= 0) {
      if (!data[i][checkCol]) continue;
    } else {
      const hasData = data[i].some((v, idx) => idx > 0 && v !== '' && v !== null && v !== undefined);
      if (!hasData) continue;
    }
    leads.push(_leadObj_(hdr, data[i], i + 1, tz));
  }

  return { leads: leads, headers: hdr };
}

// ── helpers de leitura (mesmo formato do getData_) ──
function _hdrLimpo_(headers) {
  return headers.map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim());
}
function _leadObj_(hdr, vals, rowNum, tz) {
  tz = tz || Session.getScriptTimeZone();
  const row = {};
  for (let j = 0; j < hdr.length; j++) {
    let val = vals[j];
    if (val instanceof Date) val = Utilities.formatDate(val, tz, 'dd/MM/yyyy');
    // Valor do Pedido: a planilha pode ter convertido o texto em número — devolve sempre "R$ 1.234,56"
    if (hdr[j] === 'Valor do Pedido' && typeof val === 'number') val = _brlTexto_(val);
    row[hdr[j]] = val || '';
  }
  row._row = rowNum;
  return row;
}
// Número → "R$ 1.234,56" (padrão único do CRM)
function _brlTexto_(n) {
  n = Math.round((Number(n) || 0) * 100) / 100;
  var p = Math.abs(n).toFixed(2).split('.');
  return (n < 0 ? '-' : '') + 'R$ ' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1];
}
// Grava o Valor do Pedido como TEXTO puro, para a planilha não reinterpretar ponto/vírgula
function _gravaValorTexto_(ws, rowNum, headers, lead) {
  if (!lead || !lead.hasOwnProperty('Valor do Pedido')) return;
  var c = headers.indexOf('Valor do Pedido');
  if (c < 0) return;
  var v = lead['Valor do Pedido'];
  if (typeof v === 'number') v = _brlTexto_(v);
  ws.getRange(rowNum, c + 1).setNumberFormat('@').setValue(String(v == null ? '' : v));
}
// Lê de volta UMA linha já gravada, no formato do getData_. Retorna null se o getData_
// não listaria essa linha (sem Razão Social) — aí o front faz o recarregamento completo.
function _linhaComoLead_(ws, rowNum, headers) {
  const hdr = _hdrLimpo_(headers);
  const vals = ws.getRange(rowNum, 1, 1, hdr.length).getValues()[0];
  const low = hdr.map(h => h.toLowerCase());
  let rc = low.indexOf('razão social'); if (rc < 0) rc = low.indexOf('razao social');
  if (rc >= 0 && !vals[rc]) return null;
  return _leadObj_(hdr, vals, rowNum);
}
function _colLetra_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Garante que colunas existam na planilha; cria se faltarem
function _ensureColumns_(ws, colNames) {
  var lastCol = ws.getLastColumn();
  var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h){ return String(h).replace(/\n/g,' ').replace(/\s+/g,' ').trim(); });
  var added = false;
  colNames.forEach(function(name) {
    if (headers.indexOf(name) < 0) {
      lastCol++;
      ws.getRange(1, lastCol).setValue(name);
      headers.push(name);
      added = true;
    }
  });
  return added;
}

// Texto que começa com + = - @ a planilha lê como FÓRMULA → vira #ERROR!
// (ex.: telefone internacional "+54 11 4444-5555"). O apóstrofo força texto;
// ao ler de volta (getValues) o apóstrofo não aparece.
function _txtSeguro_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function addLead_(body, sheetName) {
  const email = (body.userEmail || '').trim().toLowerCase();
  const user  = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  const ws = _sheet_(sheetName || 'LEADS');
  // Trava curta: dois cadastros ao mesmo tempo não pegam o mesmo Nº nem a mesma linha
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { error: 'Muitos cadastros ao mesmo tempo. Tente de novo em alguns segundos.' };
  try {
  // Garantir colunas novas
  _ensureColumns_(ws, ['Status Inicial', 'E-mail Prospect', 'Links', 'Anexos', 'Registro', 'País', 'Telefone Internacional', 'Atividades', 'Pedidos', 'Telefone RF', 'Oculto', 'Origem', 'Porte']);

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
    .map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim());

  const lastRow = ws.getLastRow();
  let nextNum   = 1;
  if (lastRow > 1) {
    const nums = ws.getRange(1, 1, lastRow, 1).getValues();
    for (let i = nums.length - 1; i >= 1; i--) {
      if (nums[i][0] && !isNaN(nums[i][0])) { nextNum = Number(nums[i][0]) + 1; break; }
    }
  }

  const lead = body.lead || {};
  _rsUpper_(lead);

  // Bloqueio de CNPJ duplicado (varre LEADS e WHATSAPP) — impede criar lead repetido
  var _novoCnpj = String(lead['CNPJ'] || '').replace(/\D/g, '');
  if (_novoCnpj.length === 14) {
    var _dupOnde = _cnpjJaExiste_(_novoCnpj);
    if (_dupOnde) return { error: 'CNPJ já cadastrado em ' + _dupOnde + '. Não é possível criar um lead duplicado.' };
  }

  const row  = headers.map(h => {
    if (h === 'Nº')               return nextNum;
    if (h === 'Data')             return lead['Data'] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    if (h === 'Temperatura Lead') return _calcTemp_(lead['Status'] || 'Aberto');
    return _txtSeguro_(lead[h] !== undefined ? lead[h] : '');
  });

  ws.appendRow(row);
  const novaLinha = ws.getLastRow();
  _gravaValorTexto_(ws, novaLinha, headers, lead);
  _bumpRev_();
  return { ok: true, num: nextNum, row: novaLinha, lead: _linhaComoLead_(ws, novaLinha, headers) };
  } finally { lock.releaseLock(); }
}

function editLead_(body, sheetName) {
  const email = (body.userEmail || '').trim().toLowerCase();
  const user  = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  const ws     = _sheet_(sheetName || 'LEADS');
  const rowNum = body.row;
  if (!rowNum || rowNum < 2) return { error: 'Linha inválida' };

  // Garantir colunas novas
  _ensureColumns_(ws, ['Status Inicial', 'E-mail Prospect', 'Links', 'Anexos', 'Registro', 'País', 'Telefone Internacional', 'Atividades', 'Pedidos', 'Telefone RF', 'Oculto', 'Origem', 'Porte']);

  const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
    .map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim());

  // ── Validação centralizada de permissão ──
  // admin e e-mails autorizados: editam qualquer lead.
  // demais: só podem editar o lead onde são o Prospect (dono).
  const _priv = user.admin || _isAuthorizedEmail_(email) || _isSupervisor_(email);
  const _prospectCol = headers.indexOf('Prospect');
  const _emailProspectCol = headers.indexOf('E-mail Prospect');
  const _cur = ws.getRange(rowNum, 1, 1, headers.length).getValues()[0];   // linha atual, 1 leitura só
  const _currentProspect = _prospectCol >= 0
    ? String(_cur[_prospectCol] || '').trim()
    : '';
  const _currentProspectEmail = _emailProspectCol >= 0
    ? String(_cur[_emailProspectCol] || '').trim().toLowerCase()
    : '';
  // Nome do vendedor na aba PROSPECTS correspondente ao e-mail logado (contorna divergência
  // entre o nome de login em USUARIOS e o nome do vendedor em PROSPECTS).
  const _myProspectName = _prospectNameForEmail_(email);
  const _isOwner =
    (_currentProspect.toLowerCase() === String(user.nome || '').trim().toLowerCase()) ||
    (_currentProspectEmail && _currentProspectEmail === email) ||
    (_myProspectName && _currentProspect.toLowerCase() === _myProspectName.toLowerCase());
  if (!_priv && !_isOwner) {
    return { error: 'Sem permissão: este lead pertence a outro vendedor.' };
  }
  // Usuário comum (dono) não pode reatribuir o lead direto — precisa do fluxo de transferência.
  if (!_priv && (body.lead || {}).hasOwnProperty('Prospect')) {
    const _novoPr = String((body.lead || {})['Prospect'] || '').trim();
    if (_novoPr && _novoPr.toLowerCase() !== _currentProspect.toLowerCase()) {
      return { error: 'Para transferir o lead a outro vendedor, use "Solicitar transferência".' };
    }
  }

  const lead = body.lead || {};
  _rsUpper_(lead);
  const row  = headers.map((h, i) => {
    if (h === 'Nº')               return _cur[0];
    if (h === 'Temperatura Lead') return lead.hasOwnProperty('Status') ? _calcTemp_(lead['Status']) : _cur[i];
    if (lead.hasOwnProperty(h))   return _txtSeguro_(lead[h]);
    return _txtSeguro_(_cur[i]);
  });

  ws.getRange(rowNum, 1, 1, row.length).setValues([row]);
  _gravaValorTexto_(ws, rowNum, headers, lead);
  _bumpRev_();
  return { ok: true, row: rowNum, lead: _linhaComoLead_(ws, rowNum, headers) };
}

// Verifica se um CNPJ (só dígitos) já existe em LEADS ou WHATSAPP. Retorna o nome da aba ou ''.
function _cnpjJaExiste_(digits) {
  var alvo = String(digits || '').replace(/\D/g, '');
  if (alvo.length !== 14) return '';
  var abas = [['LEADS','Leads Multicanais'], ['WHATSAPP','WhatsApp']];
  for (var s = 0; s < abas.length; s++) {
    var ws = SS.getSheetByName(abas[s][0]);
    if (!ws || ws.getLastRow() < 2) continue;
    var headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).replace(/\n/g,' ').replace(/\s+/g,' ').trim(); });
    var c = headers.indexOf('CNPJ');
    if (c < 0) continue;
    var col = ws.getRange(2, c + 1, ws.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      if (String(col[i][0] || '').replace(/\D/g, '') === alvo) return abas[s][1];
    }
  }
  return '';
}

// Retorna o Nome do vendedor na aba PROSPECTS cujo e-mail (coluna B) bate com o informado.
function _prospectNameForEmail_(email){
  var e = String(email || '').trim().toLowerCase();
  if (!e) return '';
  var ws = SS.getSheetByName('PROSPECTS');
  if (!ws) return '';
  var data = ws.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim().toLowerCase() === e) {
      return String(data[i][0] || '').trim();
    }
  }
  return '';
}

// Retorna true se o e-mail está na aba USUARIOS AUTORIZADOS (mesma lista do direct-assign)
// ─── SUPERVISÃO: vê tudo, edita o cadastro de qualquer lead e atribui; não exclui e não acessa telas de admin ───
const SUPERVISORES = ['camily.soares@dellamed.com.br'];
function _isSupervisor_(email) { return SUPERVISORES.indexOf(String(email || '').trim().toLowerCase()) >= 0; }

function _isAuthorizedEmail_(email) {
  try {
    var r = getAuthorizedUsers_();
    var list = (r && r.emails) ? r.emails : [];
    return list.indexOf(String(email || '').trim().toLowerCase()) >= 0;
  } catch (_) {
    return false;
  }
}

// Registra uma exclusão de lead no ALTERACOES (aparece no painel do admin) e avisa o admin por e-mail.
function _logExclusao_(sheetName, leadRow, razaoSocial, prospect, actorNome, actorEmail) {
  try {
    var ws = _sheet_('ALTERACOES');
    _ensureColumns_(ws, ['Observação']);
    var obsEx = (actorNome || 'Alguém') + ' excluiu o lead'
      + (razaoSocial ? (' ' + razaoSocial) : '')
      + (sheetName ? (' em ' + (String(sheetName).toUpperCase().indexOf('WHATS')>=0 ? 'WhatsApp' : 'Leads Multicanais')) : '');
    ws.appendRow([
      new Date(),
      ADMIN_EMAIL,      // [1] destinatário da notificação = admin
      actorNome || '',  // [2] quem excluiu (Solicitante)
      'EXCLUSAO',       // [3] tipo
      leadRow || '',    // [4]
      razaoSocial || '',// [5]
      prospect || '',   // [6] prospect do lead excluído (De)
      '',               // [7] Para
      'excluido',       // [8] status
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      '',               // [10]
      obsEx             // [11] Observação
    ]);
  } catch (_) {}
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[CRM Dellamed] Lead excluído',
      htmlBody:
        '<h2>Lead excluído</h2>' +
        '<p><b>' + (actorNome || actorEmail) + '</b> excluiu o lead <b>' +
        (razaoSocial || '(sem razão social)') + '</b>' +
        (prospect ? ' (prospect: ' + prospect + ')' : '') +
        ' na aba <b>' + sheetName + '</b>.</p>'
    });
  } catch (_) {}
}

function deleteLead_(body, sheetName) {
  const email = (body.userEmail || '').trim().toLowerCase();
  const user  = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  const ws     = _sheet_(sheetName || 'LEADS');
  const rowNum = body.row;
  if (!rowNum || rowNum < 2) return { error: 'Linha inválida' };

  // Excluir: somente administradora e usuários autorizados (aba USUARIOS AUTORIZADOS)
  if (!user.admin && (!_isAuthorizedEmail_(email) || _isSupervisor_(email))) {
    return { error: 'Somente a administradora e usuários autorizados podem excluir leads.' };
  }

  // Notificar o admin da exclusão (registro + e-mail) ANTES de apagar a linha
  try {
    var _hdr = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).replace(/\n/g,' ').replace(/\s+/g,' ').trim(); });
    var _rsCol = _hdr.indexOf('Razão Social'); if (_rsCol < 0) _rsCol = _hdr.indexOf('Razao Social');
    var _prCol = _hdr.indexOf('Prospect');
    var _rs = _rsCol >= 0 ? String(ws.getRange(rowNum, _rsCol + 1).getValue() || '') : '';
    var _pr = _prCol >= 0 ? String(ws.getRange(rowNum, _prCol + 1).getValue() || '') : '';
    _logExclusao_(sheetName || 'LEADS', rowNum, _rs, _pr, user.nome, email);
  } catch (_) {}

  ws.deleteRow(rowNum);
  _bumpRev_();
  return { ok: true };
}

// ─── CONFIG ────────────────────────────────────────────────────────

function getConfig_() {
  const ws   = _sheet_('PAINEL');
  const data = ws.getDataRange().getValues();

  const cols   = { A: [], C: [], E: [], G: [], I: [], K: [] };
  const colMap = { A: 0, C: 2, E: 4, G: 6, I: 8, K: 10 };

  for (let i = 1; i < data.length; i++) {
    for (const [key, idx] of Object.entries(colMap)) {
      if (data[i][idx]) cols[key].push(String(data[i][idx]).trim());
    }
  }

  return {
    motivoPerda:    cols.A,
    canalPrincipal: cols.C,
    linhaInteresse: cols.E,
    timeComercial:  cols.G,
    status:         cols.I,
    origem:         cols.K,
    ufs: ['AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT',
          'PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO']
  };
}

function getUsers_() {
  const ws   = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;          // linha vazia na planilha
    // Tudo vira texto: o google.script.run devolve NULL se houver qualquer Date na resposta
    // (a planilha converte sozinha o "dd/MM/yyyy HH:mm" do último acesso em data).
    users.push({
      email:        String(data[i][0]).trim(),
      nome:         String(data[i][1] || '').trim(),
      status:       String(data[i][2] || '').trim().toLowerCase(),
      data:         data[i][3] instanceof Date
                      ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), 'dd/MM/yyyy')
                      : String(data[i][3] || ''),
      admin:        String(data[i][4] || ''),
      ultimoAcesso: _fmtDT_(data[i][5]) || '—',
      principal:    String(data[i][0]).trim().toLowerCase() === ADMIN_EMAIL
    });
  }
  return { users: users };
}

// ─── HELPERS ───────────────────────────────────────────────────────

function _sheet_(name) {
  let ws = SS.getSheetByName(name);
  if (!ws) {
    ws = SS.insertSheet(name);
    if (name === 'USUARIOS') {
      ws.appendRow(['Email', 'Nome', 'Status', 'DataCadastro', 'Admin', 'Último Acesso']);
      ws.appendRow([ADMIN_EMAIL, 'Bianca Lima', 'aprovado', new Date(), 'sim', '']);
    }
    if (name === 'ACESSOS') {
      ws.appendRow(['Nome', 'Email']);
      ws.appendRow(['Bianca Lima', ADMIN_EMAIL]);
    }
    if (name === 'PROSPECTS') {
      ws.appendRow(['Nome', 'Email']);
      ws.appendRow(['Bianca Lima', ADMIN_EMAIL]);
    }
    if (name === 'ALTERACOES') {
      ws.appendRow(['Data','Email','Nome','Aba','LinhaLead','RazãoSocial','ProspectAtual','NovoProspect','Status','ResolvidoEm','Visto','Observação']);
    }
  }
  return ws;
}

function _getUser_(email) {
  const ws   = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      if (String(data[i][2]).trim().toLowerCase() !== 'aprovado') return null;
      return {
        email: email,
        nome:  String(data[i][1]).trim(),
        admin: email === ADMIN_EMAIL || String(data[i][4]).trim().toLowerCase() === 'sim'
      };
    }
  }
  return null;
}

function _requireAdmin_(email) {
  const e    = (email || '').trim().toLowerCase();
  const user = _getUser_(e);
  if (!user || !user.admin) throw new Error('Apenas administradores podem executar esta ação.');
}

function _calcTemp_(status) {
  const s = String(status).trim().toLowerCase();
  // nomes novos (Aberto, Qualificado, Convertido) + antigos, para registros ainda não migrados
  if (['aberto', 'novo', 'inativo'].includes(s))                      return 'Frio';
  if (['qualificado', 'contato feito', 'proposta enviada'].includes(s)) return 'Morno';
  if (['negociando', 'convertido', 'ganho'].includes(s))               return 'Quente';
  if (s === 'perdido')                              return 'Perdido';
  return '';
}

// ─── TRANSFER REQUESTS ─────────────────────────────────────────────

function requestTransfer_(body) {
  const email = (body.userEmail || '').trim().toLowerCase();
  const user  = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  const ws = _sheet_('ALTERACOES');
  _ensureColumns_(ws, ['Observação']);
  const alvo = (body.assignToName && body.assignToName !== user.nome) ? body.assignToName : user.nome;
  const alvoTxt = (alvo === user.nome) ? 'o próprio usuário' : alvo;
  const obsTr = user.nome + ' solicitou autorização para transferir o lead'
    + (body.razaoSocial ? (' ' + body.razaoSocial) : '')
    + ' para ' + alvoTxt
    + (body.sheetName ? (' em ' + (String(body.sheetName).toUpperCase().indexOf('WHATS')>=0 ? 'WhatsApp' : 'Leads Multicanais')) : '');
  ws.appendRow([
    new Date(), email, user.nome,
    body.sheetName || 'LEADS', body.row,
    body.razaoSocial || '', body.prospectAtual || '', alvo,
    'pendente', '', '', obsTr
  ]);

  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '[CRM Dellamed] Solicitação de transferência de lead',
      htmlBody:
        '<h2>Transferência de Lead</h2>' +
        '<p><b>' + user.nome + '</b> (' + email + ') solicita transferência do lead:</p>' +
        '<p><b>Razão Social:</b> '  + _escHtml_(body.razaoSocial  || '') + '</p>' +
        '<p><b>Prospect atual:</b> ' + _escHtml_(body.prospectAtual || '') + '</p>' +
        '<p><b>Novo Prospect:</b> ' + user.nome + '</p>' +
        '<p>Acesse o CRM para aprovar ou recusar.</p>'
    });
  } catch (_) {}

  return { ok: true };
}

// Formata qualquer valor de célula de data/hora para dd/MM/yyyy HH:mm
function _fmtDT_(val) {
  if (!val && val !== 0) return '';
  // Objeto Date nativo (caso mais comum via getValues)
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  // Numero serial do Google Sheets (dias desde 30/12/1899)
  if (typeof val === 'number') {
    var ms = (val - 25569) * 86400000;
    return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  // String ISO 8601 ex: 2026-08-12T01:08:00.000Z - converte antes de formatar
  var s = String(val);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    }
  }
  // Qualquer outra string (ex: ja formatada 12/08/2026 22:08) - retorna como esta
  return s;
}

function getTransfers_() {
  const ws   = _sheet_('ALTERACOES');
  const data = ws.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    items.push({
      _row:          i + 1,
      data:          _fmtDT_(data[i][0]),
      email:         data[i][1],
      nome:          data[i][2],
      sheetName:     data[i][3],
      leadRow:       data[i][4],
      razaoSocial:   data[i][5],
      prospectAtual: data[i][6],
      novoProspect:  data[i][7],
      status:        data[i][8],
      resolvidoEm:   _fmtDT_(data[i][9])
    });
  }
  return { transfers: items };
}

function approveTransfer_(body) {
  _requireAdmin_(body.requesterEmail);
  const ws  = _sheet_('ALTERACOES');
  const row = body.row;
  if (!row || row < 2) return { error: 'Linha inválida' };

  const data        = ws.getRange(row, 1, 1, 10).getValues()[0];
  const sheetName   = data[3] || 'LEADS';
  const leadRow     = data[4];
  const novoProspect = data[7];

  const leadWs = _sheet_(sheetName);
  const headers = leadWs.getRange(1, 1, 1, leadWs.getLastColumn()).getValues()[0]
    .map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim());
  const prospectCol = headers.indexOf('Prospect') + 1;
  if (prospectCol > 0 && leadRow > 1) {
    leadWs.getRange(leadRow, prospectCol).setValue(novoProspect);
    // junto com o Prospect, atualiza o e-mail e o Time do vendedor (aba PROSPECTS)
    try {
      const pv = (getProspects_().prospects || []).find(function (p) { return p.nome === novoProspect; });
      const ce = headers.indexOf('E-mail Prospect') + 1, ct = headers.indexOf('Time Comercial') + 1;
      if (pv && ce > 0) leadWs.getRange(leadRow, ce).setValue(pv.email);
      if (pv && pv.time && ct > 0) leadWs.getRange(leadRow, ct).setValue(pv.time);
    } catch (_) {}
    _bumpRev_();
  }

  ws.getRange(row, 9).setValue('aprovado');
  ws.getRange(row, 10).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  _ensureColumns_(ws, ['Observação']);
  ws.getRange(row, 12).setValue('Transferência APROVADA — lead ' + (data[5] || '') + ' passou para ' + (novoProspect || ''));
  return { ok: true };
}

function rejectTransfer_(body) {
  _requireAdmin_(body.requesterEmail);
  const ws  = _sheet_('ALTERACOES');
  const row = body.row;
  if (!row || row < 2) return { error: 'Linha inválida' };
  ws.getRange(row, 9).setValue('recusado');
  ws.getRange(row, 10).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  _ensureColumns_(ws, ['Observação']);
  ws.getRange(row, 12).setValue('Transferência RECUSADA');
  return { ok: true };
}

function getNotifications_(email) {
  if (!email) return { notifications: [] };
  const e    = email.trim().toLowerCase();
  const ws   = _sheet_('ALTERACOES');
  const data = ws.getDataRange().getValues();
  const notifs = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === e && !data[i][10]) {
      const status = data[i][8];
      if (status === 'aprovado' || status === 'recusado') {
        notifs.push({
          _row: i + 1, razaoSocial: data[i][5], status: status,
          tipo: 'transferencia',
          resolvidoEm: data[i][9] instanceof Date
            ? Utilities.formatDate(data[i][9], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : data[i][9],
          visto: ''
        });
      } else if (status === 'atividade' || status === 'atribuido' || status === 'removido') {
        notifs.push({
          _row: i + 1, tipo: status,
          autor: data[i][2], detalhe: data[i][7], razaoSocial: data[i][5],
          texto: data[i][11] || '',
          resolvidoEm: data[i][9] instanceof Date
            ? Utilities.formatDate(data[i][9], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : data[i][9],
          visto: ''
        });
      }
    }
  }
  return { notifications: notifs };
}

function clearNotifications_(body) {
  const email = (body.userEmail || '').trim().toLowerCase();
  if (!email) return { error: 'E-mail obrigatório' };
  const ws   = _sheet_('ALTERACOES');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][1]).trim().toLowerCase() === email &&
      (data[i][8] === 'aprovado' || data[i][8] === 'recusado' || data[i][8] === 'atividade' || data[i][8] === 'atribuido' || data[i][8] === 'removido') &&
      !data[i][10]
    ) {
      ws.getRange(i + 1, 11).setValue(
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
      );
    }
  }
  return { ok: true };
}

// ─── LOOKUP & ASSIGN ───────────────────────────────────────────────

function lookupUser_(email) {
  if (!email) return { error: 'E-mail obrigatório' };
  const e    = email.trim().toLowerCase();
  const ws   = _sheet_('USUARIOS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === e) {
      if (String(data[i][2]).trim().toLowerCase() === 'aprovado') {
        return { ok: true, nome: data[i][1], email: e };
      }
      return { error: 'Usuário encontrado mas não está aprovado.' };
    }
  }
  return { error: 'Usuário não encontrado. Verifique o e-mail.' };
}

function notifyActivity_(body) {
  var autor = (body.userNome || body.userEmail || '').trim();
  var detalhe = String(body.detalhe || '').trim();
  if (!detalhe) return { ok: true };
  var ws = _sheet_('ALTERACOES');
  _ensureColumns_(ws, ['Observação']);
  var obsAt = autor + ' ' + detalhe
    + (body.razaoSocial ? (' — ' + body.razaoSocial) : '')
    + (body.aba ? (' em ' + body.aba) : '');
  ws.appendRow([
    new Date(),
    ADMIN_EMAIL,                 // [1] destinatário = admin
    autor,                       // [2] quem fez a ação
    body.aba || '',              // [3] aba
    body.leadRow || '',          // [4] linha do lead
    body.razaoSocial || '',      // [5] razão social
    '',                          // [6]
    detalhe,                     // [7] descrição da atividade
    'atividade',                 // [8] status/tipo
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    '',                          // [10] visto
    obsAt                        // [11] Observação (frase pronta)
  ]);
  return { ok: true };
}

function notifyAssign_(body) {
  let assignEmail  = (body.assignEmail || '').trim().toLowerCase();
  let removedEmail = (body.removedEmail || '').trim().toLowerCase();
  // Segurança: só avisa e-mails da equipe (aba PROSPECTS ou usuário aprovado) — ninguém usa o CRM para mandar e-mail para fora
  if (assignEmail && !_emailDaEquipe_(assignEmail)) assignEmail = '';
  if (removedEmail && !_emailDaEquipe_(removedEmail)) removedEmail = '';
  if (!assignEmail && !removedEmail) return { ok: true };

  const ws = _sheet_('ALTERACOES');
  _ensureColumns_(ws, ['Observação']);
  const _quem = body.fromUser || 'Admin';
  const _rs   = body.razaoSocial || '';
  const _nowN = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // Prospect ANTERIOR (perdeu o lead) — notificado, com quem fez
  if (removedEmail && removedEmail !== assignEmail) {
    ws.appendRow([
      new Date(), removedEmail, _quem, 'REMOCAO', '', _rs,
      body.prospectAtual || '', body.assignToName || '', 'removido', _nowN, '',
      _quem + ' transferiu o lead' + (_rs ? (' ' + _rs) : '') + ' do seu nome'
        + (body.assignToName ? (' para ' + body.assignToName) : '')
    ]);
    try {
      MailApp.sendEmail({
        to: removedEmail, cc: ADMIN_EMAIL,
        subject: '[CRM Dellamed] Lead removido do seu nome',
        htmlBody: '<p><b>' + _escHtml_(_quem) + '</b> transferiu o lead <b>' + _escHtml_(_rs) + '</b> do seu nome'
          + (body.assignToName ? (' para <b>' + _escHtml_(body.assignToName) + '</b>') : '') + '.</p>'
      });
    } catch (_) {}
  }

  if (!assignEmail) return { ok: true };
  ws.appendRow([
    new Date(),
    assignEmail,                  // [1] e-mail do destinatário (para notificações)
    _quem,                        // [2] quem atribuiu
    'ATRIBUICAO',                 // [3]
    '',                           // [4] leadRow
    body.razaoSocial || '',       // [5] razão social do lead
    body.prospectAtual || '',     // [6] prospect anterior (coluna De)
    body.assignToName || '',      // [7] prospect novo — quem recebeu (coluna Para)
    'atribuido',                  // [8] status
    _nowN,
    '',
    (_quem + ' atribuiu o lead' + (_rs ? (' ' + _rs) : '') + ' para você')   // [11] Observação
  ]);

  try {
    MailApp.sendEmail({
      to: assignEmail,
      cc: ADMIN_EMAIL,
      subject: '[CRM Dellamed] Lead atribuído a você',
      htmlBody:
        '<h2>Novo lead atribuído</h2>' +
        '<p><b>' + _escHtml_(_quem) + '</b> atribuiu o lead <b>' + _escHtml_(_rs) + '</b> a você.</p>' +
        '<p>Acesse o CRM para ver os detalhes.</p>'
    });
  } catch (_) {}

  return { ok: true };
}

// ─── ONE-TIME FIX: Copiar Data → Data Primeiro Contato onde vazio ──

function fixDates() {
  _soEditor_();
  const sheets = ['LEADS', 'WHATSAPP'];
  sheets.forEach(sheetName => {
    const ws = SS.getSheetByName(sheetName);
    if (!ws) return;
    const headers = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
      .map(h => String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim());
    const dataCol = headers.indexOf('Data') + 1;
    const primCol = headers.indexOf('Data Primeiro Contato') + 1;
    if (dataCol < 1 || primCol < 1) {
      Logger.log(sheetName + ': coluna não encontrada. Data=' + dataCol + ' Prim=' + primCol);
      return;
    }
    const lastRow = ws.getLastRow();
    if (lastRow < 2) return;
    const dataVals = ws.getRange(2, dataCol, lastRow - 1, 1).getValues();
    const primVals = ws.getRange(2, primCol, lastRow - 1, 1).getValues();
    let count = 0;
    primVals.forEach((pv, i) => {
      if (!pv[0] && dataVals[i][0]) {
        let val = dataVals[i][0];
        if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        ws.getRange(i + 2, primCol).setValue(val);
        count++;
      }
    });
    Logger.log(sheetName + ': ' + count + ' datas preenchidas');
  });
  Logger.log('fixDates concluído');
}

// ─── FILE ATTACHMENTS (Google Drive) ───────────────────────────────
// Pasta raiz criada automaticamente no Drive da conta que fez deploy.
// Cada lead ganha uma subpasta: CRM_ANEXOS/Lead_XXX_RazaoSocial

var CRM_FOLDER_NAME = 'CRM_ANEXOS';

function _getCrmFolder_() {
  var folders = DriveApp.getFoldersByName(CRM_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CRM_FOLDER_NAME);
}

function _getLeadFolder_(leadNum, razaoSocial) {
  var root = _getCrmFolder_();
  var name = 'Lead_' + String(leadNum).padStart(4, '0');
  if (razaoSocial) name += '_' + String(razaoSocial).replace(/[^a-zA-Z0-9À-ÿ ]/g, '').substring(0, 40).trim();
  var folders = root.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return root.createFolder(name);
}

function uploadFile_(body) {
  var email = (body.userEmail || '').trim().toLowerCase();
  var user = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  var fileName = body.fileName || 'arquivo';
  var mimeType = body.mimeType || 'application/octet-stream';
  var base64 = body.base64 || '';
  var leadNum = body.leadNum || 0;
  var razaoSocial = body.razaoSocial || '';

  if (!base64) return { error: 'Nenhum arquivo recebido' };

  try {
    // Decodificar base64
    var decoded = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);

    // Salvar no Drive
    var folder = _getLeadFolder_(leadNum, razaoSocial);
    var file = folder.createFile(blob);

    // Visualização só para contas da Dellamed com o link (Google Workspace).
    // Se a conta dona do script não for Workspace, cai para 'qualquer pessoa com o link':
    // o link só aparece para quem está logado no CRM.
    try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (eShare) { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }

    return {
      ok: true,
      file: {
        id: file.getId(),
        name: file.getName(),
        url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
        mimeType: file.getMimeType(),
        size: file.getSize()
      }
    };
  } catch (e) {
    Logger.log('Erro uploadFile: ' + e.message);
    return { error: 'Erro ao salvar arquivo: ' + e.message };
  }
}

function deleteFile_(body) {
  var email = (body.userEmail || '').trim().toLowerCase();
  var user = _getUser_(email);
  if (!user) return { error: 'Não autorizado' };

  var fileId = body.fileId || '';
  if (!fileId) return { error: 'ID do arquivo não informado' };
  if (!_podeRemoverAnexo_(fileId, email, user)) return { error: 'Sem permissão: este anexo é de um lead de outro vendedor.' };

  try {
    var file = DriveApp.getFileById(fileId);
    if (!_dentroDoCrm_(file)) return { error: 'Arquivo não pertence ao CRM.' };
    file.setTrashed(true);
    return { ok: true };
  } catch (e) {
    return { error: 'Arquivo não encontrado: ' + e.message };
  }
}
function _dentroDoCrm_(file) {
  var raizId = _getCrmFolder_().getId();
  var pais = file.getParents();
  while (pais.hasNext()) {
    var p = pais.next();
    if (p.getId() === raizId) return true;
    var avos = p.getParents();
    while (avos.hasNext()) { if (avos.next().getId() === raizId) return true; }
  }
  return false;
}

/**
 * LEMBRETE DE PRÓXIMA AÇÃO — Dellamed CRM
 * Cole este trecho no final do seu Codigo.gs (não substitui nada).
 * Envia um e-mail ao vendedor (E-mail Prospect) no dia da "Data Próxima Ação".
 *
 * DEPOIS DE COLAR:
 *   1) Rode a função setupLembreteTrigger() UMA vez (menu ▶ Executar) para criar
 *      o gatilho diário. Autorize quando pedir.
 *   2) Pronto. Todo dia ~08:00 (horário de Brasília) ele avisa as ações do dia.
 *
 * Observação: usa MailApp.sendEmail (funciona onde GmailApp é bloqueado).
 */

function setupLembreteTrigger() {
  _soEditor_();
  // Remove gatilhos antigos desta função para não duplicar
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'remindProximaAcao') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('remindProximaAcao')
    .timeBased().everyDays(1).atHour(8) // 08h no fuso do projeto
    .create();
  Logger.log('Gatilho diário de lembrete criado (08h).');
}

function remindProximaAcao(e) {
  // Segurança: só roda pelo gatilho diário deste projeto ou pelo editor
  // (ninguém consegue disparar os lembretes chamando esta função pelo navegador).
  var doGatilho = !!(e && e.triggerUid) && ScriptApp.getProjectTriggers().some(function (t) {
    return t.getUniqueId() === String(e.triggerUid);
  });
  if (!doGatilho) _soEditor_();
  var TZ = 'America/Sao_Paulo';
  var hoje = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ['LEADS', 'WHATSAPP'];
  var props = PropertiesService.getScriptProperties();
  var enviados = 0;

  abas.forEach(function (nomeAba) {
    var sh = ss.getSheetByName(nomeAba);
    if (!sh) return;
    var dados = sh.getDataRange().getValues();
    if (dados.length < 2) return;
    var head = dados[0].map(function (h) { return String(h).trim(); });
    var ix = function (nome) { return head.indexOf(nome); };

    var cDataAcao = ix('Data Próxima Ação');
    var cProxAcao = ix('Próxima Ação');
    var cEmailPr  = ix('E-mail Prospect');
    var cProspect = ix('Prospect');
    var cRazao    = ix('Razão Social');
    var cLead     = ix('Lead');
    if (cDataAcao < 0 || cEmailPr < 0) return; // planilha sem as colunas necessárias

    for (var r = 1; r < dados.length; r++) {
      var linha = dados[r];
      var dataAcao = _lembFmtData_(linha[cDataAcao], TZ);
      if (dataAcao !== hoje) continue;

      var email = String(linha[cEmailPr] || '').trim();
      if (!email || email.indexOf('@') < 0) continue;

      // Idempotência: não repete no mesmo dia/linha
      var chave = 'lemb_' + nomeAba + '_' + (r + 1) + '_' + hoje;
      if (props.getProperty(chave)) continue;

      var razao   = String(linha[cRazao] || linha[cLead] || 'Lead').trim();
      var acao    = String(linha[cProxAcao] || '').trim();
      var prospect= String(linha[cProspect] || '').trim();

      var assunto = 'Lembrete CRM: próxima ação hoje — ' + razao;
      var corpo =
        'Olá' + (prospect ? ' ' + prospect : '') + ',\n\n' +
        'Você tem uma próxima ação agendada para HOJE (' + hoje + ') no CRM Dellamed:\n\n' +
        '• Lead: ' + razao + '\n' +
        (acao ? '• Ação: ' + acao + '\n' : '') +
        '\nAbra o CRM para atualizar o andamento.\n\n— CRM Dellamed';

      try {
        MailApp.sendEmail(email, assunto, corpo);
        props.setProperty(chave, '1');
        enviados++;
      } catch (e) {
        Logger.log('Falha ao enviar lembrete (' + nomeAba + ' linha ' + (r + 1) + '): ' + e);
      }
    }
  });

  Logger.log('Lembretes enviados: ' + enviados);
  return enviados;
}

// Aceita Date (serial do Sheets) ou texto "dd/mm/aaaa" e devolve "dd/MM/yyyy"
function _lembFmtData_(v, TZ) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'dd/MM/yyyy');
  var s = String(v || '').trim();
  var m = s.split(' ')[0].split('/');
  if (m.length === 3) {
    var dd = ('0' + m[0]).slice(-2), mm = ('0' + m[1]).slice(-2), yy = m[2];
    return dd + '/' + mm + '/' + yy;
  }
  return s;
}


// ═══════════════════ CARGA EM LOTE: dados cadastrais via CNPJ ═══════════════════
// ▶ Rode preencherDadosCadastraisPorCNPJ() no editor (botão Executar).
//   Percorre a aba LEADS, consulta a Receita (open.cnpja.com) por CNPJ e preenche:
//   Razão Social, Cidade, UF, Endereço Fiscal, CNAE, Porte, SUFRAMA, Regime, Telefone RF.
//   Sobrescreve os campos cadastrais (conforme solicitado). NÃO toca em Status/Prospect/Time.
//   Respeita ~1,2s entre chamadas para não estourar o limite da API pública.
function preencherDadosCadastraisPorCNPJ() {
  _soEditor_();
  return _cargaCNPJ_('LEADS');
}
function preencherDadosCadastraisPorCNPJ_WhatsApp() {
  _soEditor_();
  return _cargaCNPJ_('WHATSAPP');
}

function _cargaCNPJ_(sheetName) {
  var ws = SS.getSheetByName(sheetName);
  if (!ws) { Logger.log('Aba ' + sheetName + ' não encontrada.'); return; }

  // Garante colunas cadastrais
  _ensureColumns_(ws, ['Razão Social','Cidade','UF','Endereço Fiscal','CNAE','Porte','SUFRAMA','Regime','Telefone RF','CNPJ']);

  var lastCol = ws.getLastColumn(), lastRow = ws.getLastRow();
  if (lastRow < 2) { Logger.log('Sem linhas de dados.'); return; }

  var headers = ws.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h){ return String(h).replace(/\n/g,' ').replace(/\s+/g,' ').trim(); });
  function col(name){ return headers.indexOf(name); } // índice 0-based
  var cCNPJ = col('CNPJ');
  if (cCNPJ < 0) { Logger.log('Coluna CNPJ não encontrada.'); return; }

  var map = {
    rs:  col('Razão Social'),
    ci:  col('Cidade'),
    uf:  col('UF'),
    end: col('Endereço Fiscal'),
    cnae:col('CNAE'),
    porte:col('Porte'),
    suf: col('SUFRAMA'),
    reg: col('Regime'),
    telrf:col('Telefone RF')
  };

  var data = ws.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var okc = 0, err = 0, skip = 0;

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var cnpjDigits = String(data[i][cCNPJ] || '').replace(/\D/g, '');
    if (cnpjDigits.length !== 14) { skip++; continue; }

    var d;
    try {
      var resp = UrlFetchApp.fetch('https://open.cnpja.com/office/' + cnpjDigits + '?suframa=true', {
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) { err++; Logger.log('CNPJ ' + cnpjDigits + ' → HTTP ' + resp.getResponseCode()); Utilities.sleep(1200); continue; }
      d = JSON.parse(resp.getContentText());
    } catch (e) {
      err++; Logger.log('CNPJ ' + cnpjDigits + ' → erro: ' + e.message); Utilities.sleep(1200); continue;
    }

    var a = d.address || {};
    function put(idx, val){ if (idx >= 0 && val !== undefined && val !== null && val !== '') ws.getRange(rowNum, idx + 1).setValue(val); }

    // Razão Social
    put(map.rs, (d.company && d.company.name) ? d.company.name : (d.alias || ''));
    // Endereço fiscal
    put(map.end, [a.street, a.number, a.district].filter(Boolean).join(', '));
    // Cidade / UF
    put(map.ci, a.city || '');
    put(map.uf, a.state || '');
    // Telefone(s) RF
    var tels = (d.phones || []).map(function(p){
      var ddd = String(p.area || '').replace(/\D/g,'');
      var num = String(p.number || '').replace(/\D/g,'');
      if (!ddd && !num) return '';
      var fmt = num.length >= 9 ? num.replace(/^(\d{5})(\d{4})$/,'$1-$2') : num.replace(/^(\d{4})(\d{4})$/,'$1-$2');
      return (ddd ? '(' + ddd + ') ' : '') + fmt;
    }).filter(Boolean);
    if (tels.length) put(map.telrf, tels.join(' | '));
    // Porte
    var size = d.company && d.company.size ? d.company.size : null;
    if (size) put(map.porte, size.text || size.acronym || '');
    // SUFRAMA
    var suf = (d.suframa || []).find(function(s){ return s.approved && s.status && s.status.text === 'Ativa'; });
    put(map.suf, suf ? ('OK — ativa (nº ' + suf.number + ')') : 'Sem SUFRAMA');
    // Regime (só Simples/MEI são resolvíveis pela API)
    if (map.reg >= 0) {
      if (d.company && d.company.simei && d.company.simei.optant) ws.getRange(rowNum, map.reg + 1).setValue('MEI');
      else if (d.company && d.company.simples && d.company.simples.optant) ws.getRange(rowNum, map.reg + 1).setValue('Simples Nacional');
      // não-optante: deixa como está (Presumido/Real é manual)
    }
    // CNAE — principal + secundárias
    var cn = [];
    if (d.mainActivity) cn.push(d.mainActivity.id + ' — ' + d.mainActivity.text + ' (principal)');
    (d.sideActivities || []).forEach(function(c){ cn.push(c.id + ' — ' + c.text); });
    if (cn.length) put(map.cnae, cn.join(' | '));

    okc++;
    SpreadsheetApp.flush();
    Utilities.sleep(1200); // respeita limite da API pública (~gentil)
  }

  Logger.log('Carga concluída em ' + sheetName + ': ' + okc + ' preenchidos, ' + err + ' com erro, ' + skip + ' sem CNPJ válido.');
  return { ok: okc, erro: err, semCNPJ: skip };
}


// ═══════════════════ SINCRONIZAÇÃO COM A PLANILHA ═══════════════════
// A tela pergunta a cada ~45s se algo mudou. "rev" muda quando o CRM grava
// (addLead/editLead/deleteLead/transferência) OU quando alguém edita a planilha
// direto (data de modificação do arquivo no Drive). Só então a tela recarrega.
// ═══════════════════ RAZÃO SOCIAL SEMPRE EM MAIÚSCULAS ═══════════════════
function _rsUpper_(lead) {
  ['Razão Social', 'Razao Social'].forEach(function (k) {
    if (lead && lead[k] != null) lead[k] = String(lead[k]).toUpperCase();
  });
}
// ▶ Rode padronizarRazaoSocialMaiusculas() UMA vez pelo editor para converter o que já existe
//   (LEADS e, se ainda tiver dados, WHATSAPP). Pode rodar de novo sem problema: só mexe no que estiver diferente.
function padronizarRazaoSocialMaiusculas() {
  _soEditor_();
  var total = 0;
  ['LEADS', 'WHATSAPP'].forEach(function (nome) {
    var ws = SS.getSheetByName(nome);
    if (!ws || ws.getLastRow() < 2) return;
    var hdr = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim(); });
    var c = hdr.indexOf('Razão Social'); if (c < 0) c = hdr.indexOf('Razao Social');
    if (c < 0) return;
    var rg = ws.getRange(2, c + 1, ws.getLastRow() - 1, 1);
    var v = rg.getValues(), mudou = 0;
    for (var i = 0; i < v.length; i++) {
      var up = String(v[i][0] == null ? '' : v[i][0]).toUpperCase();
      if (up !== String(v[i][0])) { v[i][0] = up; mudou++; }
    }
    if (mudou) { rg.setValues(v); total += mudou; }
    Logger.log(nome + ': ' + mudou + ' razão(ões) social(is) convertida(s).');
  });
  _bumpRev_();
  Logger.log('Pronto. Total convertido: ' + total);
  return total;
}

// ─── Segurança: helpers ───
function _escHtml_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function _emailDaEquipe_(email) {
  email = _normEmail_(email);
  if (!email) return false;
  if (email === _normEmail_(ADMIN_EMAIL)) return true;
  if ((getProspects_().prospects || []).some(function (p) { return p.email === email; })) return true;
  var reg = _usuarioRegistro_(email);
  return !!(reg && reg.status === 'aprovado');
}
// Anexo só pode ser removido por quem pode alterar o lead onde ele está (admin, autorizados ou o dono).
// Arquivo que ainda não foi salvo em nenhum lead (acabou de subir) pode ser removido por quem subiu.
function _podeRemoverAnexo_(fileId, email, user) {
  if (user.admin || _isAuthorizedEmail_(email) || _isSupervisor_(email)) return true;
  var ws = SS.getSheetByName('LEADS');
  if (!ws || ws.getLastRow() < 2) return true;
  var hdr = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0].map(function (h) { return String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim(); });
  var cA = hdr.indexOf('Anexos'), cP = hdr.indexOf('Prospect'), cE = hdr.indexOf('E-mail Prospect');
  if (cA < 0) return true;
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn()).getValues();
  var meuNome = _norm_(_prospectNameForEmail_(email) || user.nome || '');
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][cA] || '').indexOf(fileId) < 0) continue;
    var em = cE >= 0 ? _normEmail_(data[i][cE]) : '';
    var pr = cP >= 0 ? _norm_(data[i][cP]) : '';
    return (em && em === email) || (!!pr && pr === meuNome);
  }
  return true;
}
function _norm_(v) { return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase(); }

function _bumpRev_() {
  try { PropertiesService.getScriptProperties().setProperty('crm_rev', String(Date.now())); } catch (_) {}
}
function getRev_() {
  var p = PropertiesService.getScriptProperties().getProperty('crm_rev') || '0';
  var f = '0';
  try { f = String(DriveApp.getFileById(SS.getId()).getLastUpdated().getTime()); } catch (_) {}
  return { rev: p + '_' + f };
}

// ═══════════════════ MIGRAÇÃO ÚNICA — base unificada + status novos ═══════════════════
// ▶ Rode migrarParaNovoPadrao() UMA vez pelo editor (botão Executar), ANTES de publicar a nova versão.
//   1. Faz cópia de segurança das abas LEADS e WHATSAPP (BKP_LEADS_... e BKP_WHATSAPP_...).
//   2. Junta WHATSAPP dentro de LEADS (sem perder colunas; Origem vazia vira "WhatsApp").
//   3. Renomeia status: Novo → Aberto, Contato feito → Qualificado, Ganho → Convertido
//      (também no histórico de Atividades, para as datas do dashboard continuarem certas).
//   4. Ordena por data de criação e renumera o Nº (1, 2, 3...) na ordem de criação.
//   5. Deixa a aba WHATSAPP só com o cabeçalho (não apaga a aba).
//   6. Corrige as solicitações de transferência PENDENTES para apontar para a nova linha.
function migrarParaNovoPadrao() {
  _soEditor_();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('migracao_unificada_ok')) {
    throw new Error('A migração já foi feita em ' + props.getProperty('migracao_unificada_ok') + '. Para rodar de novo, apague a propriedade migracao_unificada_ok (Configurações do projeto > Propriedades do script).');
  }
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var TZ = Session.getScriptTimeZone();
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
    var wsL = SS.getSheetByName('LEADS'), wsW = SS.getSheetByName('WHATSAPP');
    if (!wsL) throw new Error('Aba LEADS não encontrada.');
    wsL.copyTo(SS).setName('BKP_LEADS_' + stamp);
    if (wsW) wsW.copyTo(SS).setName('BKP_WHATSAPP_' + stamp);

    function limpa(h) { return String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim(); }
    var dL = wsL.getDataRange().getValues();
    var headers = dL[0].map(limpa);
    var dW = wsW && wsW.getLastRow() > 0 ? wsW.getDataRange().getValues() : [[]];
    var hW = (dW[0] || []).map(limpa);
    hW.forEach(function (h) { if (h && headers.indexOf(h) < 0) headers.push(h); });
    if (headers.indexOf('Origem') < 0) headers.push('Origem');

    var MAPA = { 'novo': 'Aberto', 'contato feito': 'Qualificado', 'ganho': 'Convertido' };
    function novoStatus(v) { var k = String(v || '').trim().toLowerCase(); return MAPA[k] || v; }
    function trocaNomes(txt) {
      return String(txt).replace(/(:\s*|→\s*)(Novo|Contato feito|Ganho)(?=\s*(→|$))/gi, function (m, pre, nome) { return pre + MAPA[nome.toLowerCase()]; });
    }
    function toDate(v) {
      if (v instanceof Date) return v.getTime();
      var p = String(v || '').split(' ')[0].split('/');
      if (p.length === 3) { var d = new Date(+p[2], +p[1] - 1, +p[0]); if (!isNaN(d)) return d.getTime(); }
      return 9e15; // sem data vai para o fim
    }
    function vazia(row) { return row.every(function (c) { return c === '' || c === null; }); }

    var itens = [];
    function junta(data, hdr, aba) {
      for (var i = 1; i < data.length; i++) {
        if (vazia(data[i])) continue;
        var o = {};
        hdr.forEach(function (h, j) { if (h) o[h] = data[i][j]; });
        if (aba === 'WHATSAPP' && !String(o['Origem'] || '').trim()) o['Origem'] = 'WhatsApp';
        o['Status'] = novoStatus(o['Status']);
        if (o['Razão Social'] != null) o['Razão Social'] = String(o['Razão Social']).toUpperCase();
        o['Temperatura Lead'] = _calcTemp_(o['Status']);
        if (headers.indexOf('Temperatura Lead') < 0) headers.push('Temperatura Lead');
        if (o['Atividades']) {
          try {
            var acts = JSON.parse(o['Atividades']);
            if (Array.isArray(acts)) {
              acts.forEach(function (a) { if (a && a.type === 'status' && a.title) a.title = trocaNomes(a.title); });
              o['Atividades'] = JSON.stringify(acts);
            }
          } catch (_) {}
        }
        itens.push({ o: o, aba: aba, linha: i + 1, t: toDate(o['Data']), n: Number(o['Nº']) || 0 });
      }
    }
    junta(dL, dL[0].map(limpa), 'LEADS');
    if (wsW) junta(dW, hW, 'WHATSAPP');

    // ordem de criação: data; empate → LEADS antes de WHATSAPP → Nº antigo
    itens.sort(function (a, b) { return a.t - b.t || (a.aba === b.aba ? 0 : (a.aba === 'LEADS' ? -1 : 1)) || a.n - b.n; });
    var nova = {};
    var rows = itens.map(function (it, idx) {
      it.o['Nº'] = idx + 1;
      nova[it.aba + ':' + it.linha] = idx + 2;
      return headers.map(function (h) { return it.o[h] !== undefined && it.o[h] !== null ? it.o[h] : ''; });
    });

    wsL.clearContents();
    wsL.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) wsL.getRange(2, 1, rows.length, headers.length).setValues(rows);
    if (wsW && wsW.getLastRow() > 1) wsW.getRange(2, 1, wsW.getLastRow() - 1, Math.max(1, wsW.getLastColumn())).clearContent();

    // solicitações pendentes passam a apontar para a nova linha em LEADS
    var wsA = SS.getSheetByName('ALTERACOES'), corrigidas = 0;
    if (wsA && wsA.getLastRow() > 1) {
      var dA = wsA.getDataRange().getValues();
      for (var r = 1; r < dA.length; r++) {
        var aba = String(dA[r][3] || '').toUpperCase();
        if (String(dA[r][8]) === 'pendente' && (aba === 'LEADS' || aba === 'WHATSAPP')) {
          var nl = nova[aba + ':' + dA[r][4]];
          if (nl) { wsA.getRange(r + 1, 4, 1, 2).setValues([['LEADS', nl]]); corrigidas++; }
        }
      }
    }
    _bumpRev_();
    props.setProperty('migracao_unificada_ok', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    var msg = 'Migração concluída: ' + rows.length + ' leads em LEADS (' + (dW.length > 1 ? (itens.filter(function(i){return i.aba==='WHATSAPP';}).length) : 0) +
      ' vieram do WHATSAPP). Nº renumerado de 1 a ' + rows.length + ' pela data de criação. ' + corrigidas + ' solicitação(ões) pendente(s) ajustada(s). Backups: BKP_LEADS_' + stamp + (wsW ? ' e BKP_WHATSAPP_' + stamp : '') + '.';
    Logger.log(msg);
    return msg;
  } finally { lock.releaseLock(); }
}
