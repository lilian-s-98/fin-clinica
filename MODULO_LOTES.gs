// ============================================================
//  MODULO_LOTES.gs — v6.0
//  Rastreamento de lotes/protocolos enviados aos convênios.
//  Pipeline de status sugerido (editável — é só texto na planilha):
//   Aberto -> Enviado -> Protocolado -> Pago Parcial / Pago Total -> Glosado
// ============================================================

const LOTES_HEADERS = ['id','lote','convenio_id','convenio_nome','data_fechamento','data_envio','data_prev_recebimento','valor_total_lote','valor_recebido','valor_glosado','status','guias_ids','observacao','criado_em','atualizado_em'];
const STATUS_LOTE_VALIDOS = ['Aberto','Enviado','Protocolado','Pago Parcial','Pago Total','Glosado'];

function getLotes(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    let rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.LOTES).map(r => _toObj(LOTES_HEADERS, r));
    if (filtros && filtros.status) rows = rows.filter(r => r.status === filtros.status);
    if (filtros && filtros.convenio_id) rows = rows.filter(r => r.convenio_id === filtros.convenio_id);
    return rows;
  } catch(e) { return []; }
}

function criarLote(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    if (!dados.convenio_id) throw new Error('Selecione o convênio do lote.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
    const id = 'LOTE' + new Date().getTime();
    const agora = new Date();
    sh.appendRow([id, dados.lote||id, dados.convenio_id, _san(dados.convenio_nome), dados.data_fechamento||'', dados.data_envio||'', dados.data_prev_recebimento||'', 0, 0, 0, 'Aberto', '', _san(dados.observacao), agora, agora]);
    _log(usuario.nome, 'NOVO_LOTE', id);
    return { ok:true, id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Vincula uma guia recém-criada a um lote existente (chamado por salvarGuia
// em Modulo_Guias.gs quando dados.lote_id vem preenchido).
function _vincularGuiaAoLote(lote_id, guia_id, valorGuia) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
  if (!sh) return;
  const allData = sh.getDataRange().getValues();
  const colGuias = LOTES_HEADERS.indexOf('guias_ids');
  const colValor = LOTES_HEADERS.indexOf('valor_total_lote');
  const colAtualizado = LOTES_HEADERS.indexOf('atualizado_em');
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === lote_id) {
      const guiasAtuais = allData[i][colGuias] ? String(allData[i][colGuias]).split(',') : [];
      guiasAtuais.push(guia_id);
      sh.getRange(i+1, colGuias+1).setValue(guiasAtuais.join(','));
      sh.getRange(i+1, colValor+1).setValue(_sanNum(allData[i][colValor]) + valorGuia);
      sh.getRange(i+1, colAtualizado+1).setValue(new Date());
      return;
    }
  }
}

function atualizarStatusLote(lote_id, status, valor_recebido, valor_glosado, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (STATUS_LOTE_VALIDOS.indexOf(status) === -1) throw new Error('status inválido. Use um de: ' + STATUS_LOTE_VALIDOS.join(', '));
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === lote_id) {
        sh.getRange(i+1, LOTES_HEADERS.indexOf('status')+1).setValue(status);
        if (valor_recebido !== undefined) sh.getRange(i+1, LOTES_HEADERS.indexOf('valor_recebido')+1).setValue(_sanNum(valor_recebido));
        if (valor_glosado !== undefined) sh.getRange(i+1, LOTES_HEADERS.indexOf('valor_glosado')+1).setValue(_sanNum(valor_glosado));
        sh.getRange(i+1, LOTES_HEADERS.indexOf('atualizado_em')+1).setValue(new Date());
        _log(usuario.nome, 'STATUS_LOTE', `${lote_id} → ${status}`);
        return { ok:true };
      }
    }
    return { ok:false, msg:'Lote não encontrado' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Visão do "status dos lotes no mês" para o dashboard/gráfico.
function getStatusLotesMes(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const lotes = getLotes(filtros, usuario);
    const porStatus = {};
    STATUS_LOTE_VALIDOS.forEach(s => porStatus[s] = { qtd:0, valor:0 });
    lotes.forEach(l => {
      if (!porStatus[l.status]) porStatus[l.status] = { qtd:0, valor:0 };
      porStatus[l.status].qtd++;
      porStatus[l.status].valor += _sanNum(l.valor_total_lote);
    });
    return { ok:true, porStatus, totalLotes: lotes.length };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}
