// ============================================================
//  MODULO_GLOSAS.gs — v6.0
//  Cadastro manual de glosas + resumo para o dashboard.
//  status possíveis: Aberta, Recorrida, Mantida, Revertida
// ============================================================

const GLOSAS_HEADERS = ['id','guia_id','convenio_nome','data_glosa','valor_glosado','motivo','status','observacao','lancado_por','criado_em'];

function getGlosas(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    let rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.GLOSAS).map(r => _toObj(GLOSAS_HEADERS, r));
    if (filtros && filtros.status) rows = rows.filter(r => r.status === filtros.status);
    if (filtros && filtros.convenio_nome) rows = rows.filter(r => r.convenio_nome === filtros.convenio_nome);
    return rows;
  } catch(e) { return []; }
}

function salvarGlosa(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (!dados.guia_id) throw new Error('Informe a guia relacionada à glosa.');
    if (!_sanNum(dados.valor_glosado)) throw new Error('Informe o valor glosado (maior que zero).');
    if (!dados.motivo) throw new Error('Informe o motivo da glosa (poka-yoke: motivo em branco impede análise futura).');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
    const id = 'GLO' + new Date().getTime();
    sh.appendRow([id, dados.guia_id, _san(dados.convenio_nome), dados.data_glosa||_dateStr(new Date()), _sanNum(dados.valor_glosado), _san(dados.motivo), dados.status||'Aberta', _san(dados.observacao), usuario.nome, new Date()]);
    _log(usuario.nome, 'NOVA_GLOSA', `${dados.guia_id} | R$ ${dados.valor_glosado}`);
    return { ok:true, id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function atualizarStatusGlosa(id, status, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (!['Aberta','Recorrida','Mantida','Revertida'].includes(status)) throw new Error('status inválido.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
    const allData = sh.getDataRange().getValues();
    const colStatus = GLOSAS_HEADERS.indexOf('status') + 1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) { sh.getRange(i+1,colStatus).setValue(status); _log(usuario.nome,'STATUS_GLOSA',`${id} → ${status}`); return {ok:true}; }
    }
    return { ok:false, msg:'Glosa não encontrada' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Chamado automaticamente por atualizarStatusGuia() quando alguém marca
// valor_glosado > 0 direto na guia, sem passar pelo formulário de Glosas —
// evita ter valor glosado "solto" na guia sem nenhum registro rastreável.
function _garantirRegistroGlosa(guia_id, convenio_nome, valor, usuario) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
  if (!sh) return;
  const allData = sh.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) if (allData[i][1] === guia_id) return; // já existe
  sh.appendRow(['GLO'+new Date().getTime(), guia_id, convenio_nome, _dateStr(new Date()), valor, 'PREENCHER MOTIVO — gerado automaticamente ao marcar status da guia', 'Aberta', '', usuario ? usuario.nome : 'sistema', new Date()]);
}

function getResumoGlosas(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const glosas = getGlosas(filtros, usuario);
    const totalGlosado = glosas.reduce((s,g)=>s+_sanNum(g.valor_glosado),0);
    const porConvenio = {};
    glosas.forEach(g => { porConvenio[g.convenio_nome] = (porConvenio[g.convenio_nome]||0) + _sanNum(g.valor_glosado); });
    const porStatus = {};
    glosas.forEach(g => { porStatus[g.status] = (porStatus[g.status]||0) + 1; });
    return { ok:true, totalGlosado, porConvenio, porStatus, qtd: glosas.length };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}
