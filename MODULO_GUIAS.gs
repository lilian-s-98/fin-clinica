// ============================================================
//  MODULO_GUIAS.gs  — v6.0
//  Cadastro de guias de convênio + itens, agora com:
//   - profissional_id/nome POR ITEM (permite guia com mais de
//     um profissional dentro do mesmo protocolo)
//   - campo "concluido" POR ITEM (SIM/NÃO) — é a base do cálculo
//     proporcional de comissão (ver Modulo_Comissionamento.gs)
//   - vínculo opcional com um Lote (ver Modulo_Lotes.gs)
//
//  ATENÇÃO (poka-yoke): se você não informar itens por profissional,
//  o sistema assume automaticamente que TODOS os itens pertencem
//  ao profissional_id principal da guia e estão 100% concluídos —
//  ou seja, o comportamento antigo (v5.0) continua funcionando sem
//  quebrar nada. A granularidade por item é opcional.
// ============================================================

const ITENS_GUIA_HEADERS = ['id','guia_id','convenio_nome','codigo','descricao','quantidade','valor_unitario','valor_total','profissional_id','profissional_nome','concluido'];
const GUIAS_HEADERS = ['id','data','mes','convenio_id','convenio_nome','paciente_id','paciente_nome','profissional_id','profissional_nome','lote','protocolo','num_nf','valor_total','prazo_dias','data_envio','data_prev_pgto','data_pgto_real','status','valor_glosado','observacao','lancado_por','criado_em','lote_id'];

function salvarGuia(dados, itens, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    if (!itens || itens.length === 0) throw new Error('A guia precisa de ao menos 1 código/procedimento.');
    if (itens.length > 20) throw new Error('Máximo de 20 itens por guia (poka-yoke: confira se não duplicou lançamento).');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shG = ss.getSheetByName(CONFIG.SHEETS.GUIAS);
    const shI = ss.getSheetByName(CONFIG.SHEETS.ITENS_GUIA);
    const id = 'GUIA' + new Date().getTime();
    const mes = _mesDoDate(new Date(dados.data));
    const valorTotal = itens.reduce((s,it) => s + (_sanNum(it.valor_unitario)*parseInt(it.quantidade||1)),0);

    // Cálculo de datas usa as Regras_Convenio quando existirem (Modulo_Convenios_Faturamento.gs).
    // Se não houver regra cadastrada para o convênio, cai no prazo_dias informado manualmente
    // (comportamento antigo), para nunca travar o lançamento por falta de configuração.
    let dataPrev = '';
    const prazoDias = dados.prazo_dias ? parseInt(dados.prazo_dias) : null;
    if (dados.data_envio && prazoDias) {
      const d = new Date(dados.data_envio);
      d.setDate(d.getDate() + prazoDias);
      dataPrev = Utilities.formatDate(d,'America/Sao_Paulo','dd/MM/yyyy');
    }

    shG.appendRow([id,dados.data,mes,dados.convenio_id,dados.convenio_nome,dados.paciente_id,dados.paciente_nome,dados.profissional_id,dados.profissional_nome,dados.lote||'',dados.protocolo||'',dados.num_nf||'',valorTotal,prazoDias||60,dados.data_envio||'',dataPrev,'','Pendente',0,dados.observacao||'',usuario.nome,new Date(),dados.lote_id||'']);

    itens.forEach((it,idx) => {
      const vt = _sanNum(it.valor_unitario)*parseInt(it.quantidade||1);
      // poka-yoke: se o item não vier com profissional próprio, herda o da guia.
      const profId = it.profissional_id || dados.profissional_id;
      const profNome = it.profissional_nome || dados.profissional_nome;
      // poka-yoke: se "concluido" não for informado, assume SIM (não trava lançamento
      // simples do dia-a-dia por causa de um campo novo que a recepção não conhece ainda).
      const concluido = (it.concluido === false || it.concluido === 'NAO' || it.concluido === 'NÃO') ? 'NÃO' : 'SIM';
      shI.appendRow([id+'_'+(idx+1),id,dados.convenio_nome,it.codigo,it.descricao,parseInt(it.quantidade||1),_sanNum(it.valor_unitario),vt,profId,profNome,concluido]);
    });

    // se a guia já nasce vinculada a um lote, atualiza o lote (soma valor + lista de guias)
    if (dados.lote_id) _vincularGuiaAoLote(dados.lote_id, id, valorTotal);

    _log(usuario.nome,'NOVA_GUIA',`${id} | ${dados.convenio_nome} | R$ ${valorTotal.toFixed(2)}`);
    return {ok:true,id,valorTotal};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function getGuias(filtros) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.GUIAS);
  let result = rows.map(r => _toObj(GUIAS_HEADERS,r));
  if (filtros && filtros.mes && filtros.mes !== 'todos') result = result.filter(r => r.mes === filtros.mes);
  if (filtros && filtros.convenio_id) result = result.filter(r => r.convenio_id === filtros.convenio_id);
  if (filtros && filtros.profissional_id) result = result.filter(r => r.profissional_id === filtros.profissional_id);
  if (filtros && filtros.status) result = result.filter(r => r.status === filtros.status);
  return result;
}

function getItensGuia(guia_id) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.ITENS_GUIA);
  return rows.filter(r => r[1]===guia_id).map(r => _toObj(ITENS_GUIA_HEADERS,r));
}

// Acompanhamento de guias por profissional — "quais guias pertencem a quais profissionais"
function getGuiasPorProfissional(profissional_id, filtros) {
  const guias = getGuias(filtros).filter(g => g.profissional_id === profissional_id);
  return guias.map(g => ({ ...g, itens: getItensGuia(g.id) }));
}

// Marca um item específico da guia como concluído/não concluído — usado quando
// mais de um profissional participa da mesma guia (ex.: paciente trocou de terapeuta
// no meio do protocolo, ou parte dos procedimentos ainda não foi realizada).
function marcarItemGuiaConcluido(item_id, concluido, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','fisioterapeuta']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.ITENS_GUIA);
    const allData = sh.getDataRange().getValues();
    const colConcluido = ITENS_GUIA_HEADERS.indexOf('concluido') + 1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === item_id) {
        sh.getRange(i+1, colConcluido).setValue(concluido ? 'SIM' : 'NÃO');
        _log(usuario.nome, 'ITEM_GUIA_STATUS', `${item_id} → ${concluido?'SIM':'NÃO'}`);
        return {ok:true};
      }
    }
    return {ok:false, msg:'Item não encontrado'};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function atualizarStatusGuia(guia_id, status, data_pgto, valor_glosado, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GUIAS);
    const allData = sh.getDataRange().getValues();
    const colDataPgto = GUIAS_HEADERS.indexOf('data_pgto_real') + 1;
    const colStatus = GUIAS_HEADERS.indexOf('status') + 1;
    const colGlosa = GUIAS_HEADERS.indexOf('valor_glosado') + 1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === guia_id) {
        sh.getRange(i+1,colDataPgto).setValue(data_pgto||'');
        sh.getRange(i+1,colStatus).setValue(status);
        sh.getRange(i+1,colGlosa).setValue(_sanNum(valor_glosado)||0);
        _log(usuario.nome,'STATUS_GUIA',`${guia_id} → ${status}`);
        // poka-yoke: se marcou glosa mas não abriu registro em Glosas, cria o rascunho
        // automaticamente para não perder o controle (o usuário completa o motivo depois).
        if (_sanNum(valor_glosado) > 0) {
          _garantirRegistroGlosa(guia_id, allData[i][4], _sanNum(valor_glosado), usuario);
        }
        return {ok:true};
      }
    }
    return {ok:false,msg:'Guia não encontrada'};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}
