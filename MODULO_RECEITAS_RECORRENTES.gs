// ============================================================
//  MODULO_RECEITAS_RECORRENTES.gs — v1.0
//
//  Primeira fonte de receita recorrente do sistema fora de
//  convênio/particular: aluguel de espaço/equipamento (ex: piscina),
//  pago por um profissional à clínica todo mês. Arquitetura espelha
//  Modulo_Despesas_Recorrentes.gs de propósito, para manter o mesmo
//  padrão mental de "recorrente" em toda a planilha.
//
//  Colunas de Receitas_Recorrentes:
//   id, descricao, pagador_tipo ('profissional' ou 'outro'),
//   pagador_id, pagador_nome, categoria, valor_padrao,
//   regra_vencimento (mesmo formato de Despesas_Recorrentes: DIA_N,
//   ATE_N_DU ou MANUAL), forma_pgto, data_inicio, data_fim, ativo,
//   observacao, criado_por, criado_em
//
//  Cadastro inicial (migrarV7): Aluguel de Espaço — Piscina, Yasmin
//  Silva, R$425. Ver _popularReceitasRecorrentesPadrao() em
//  Código_Principal.gs.
// ============================================================

const RECEITAS_RECORRENTES_HEADERS = ['id','descricao','pagador_tipo','pagador_id','pagador_nome','categoria','valor_padrao','regra_vencimento','forma_pgto','data_inicio','data_fim','ativo','observacao','criado_por','criado_em'];
const RECEBIMENTOS_HEADERS_V7 = ['id','data','tipo','referencia_id','convenio_nome','paciente_nome','valor','forma_pgto','mes','observacao','criado_em'];

function getReceitasRecorrentes(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    return _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.RECEITAS_RECORRENTES).map(r => _toObj(RECEITAS_RECORRENTES_HEADERS, r));
  } catch(e) { return []; }
}

function salvarReceitaRecorrente(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _validarReceitaRecorrente(dados);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.RECEITAS_RECORRENTES);
    const allData = sh.getDataRange().getValues();
    const linha = [
      dados.id || ('REC'+Utilities.getUuid().slice(0,8).toUpperCase()), _san(dados.descricao),
      dados.pagador_tipo||'profissional', dados.pagador_id||'', _san(dados.pagador_nome),
      _san(dados.categoria)||'Aluguel de Espaço', _sanNum(dados.valor_padrao),
      dados.regra_vencimento||'MANUAL', _san(dados.forma_pgto), dados.data_inicio||'', dados.data_fim||'',
      dados.ativo||'SIM', _san(dados.observacao), usuario.nome,
      dados.id ? (allData.find(r=>r[0]===dados.id) || [,,,,,,,,,,,,,,new Date()])[14] : new Date()
    ];
    if (dados.id) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === dados.id) { sh.getRange(i+1,1,1,linha.length).setValues([linha]); _log(usuario.nome,'EDIT_RECEITA_RECORRENTE',dados.descricao); return {ok:true}; }
      }
    }
    sh.appendRow(linha);
    _log(usuario.nome, 'NOVA_RECEITA_RECORRENTE', dados.descricao);
    return { ok:true, id: linha[0] };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function excluirReceitaRecorrente(id, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.RECEITAS_RECORRENTES);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) {
        // poka-yoke: não apaga de verdade, só marca como inativa — preserva histórico
        sh.getRange(i+1, RECEITAS_RECORRENTES_HEADERS.indexOf('ativo')+1).setValue('NÃO');
        _log(usuario.nome, 'INATIVAR_RECEITA_RECORRENTE', id);
        return { ok:true };
      }
    }
    return { ok:false, msg:'Receita recorrente não encontrada.' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function _validarReceitaRecorrente(dados) {
  if (!dados.descricao) throw new Error('Descreva a receita (ex.: "Aluguel de Espaço — Piscina").');
  if (!_sanNum(dados.valor_padrao)) throw new Error('Informe um valor_padrao > 0.');
  const regra = String(dados.regra_vencimento||'');
  if (regra !== 'MANUAL' && !/^DIA_\d{1,2}$/.test(regra) && !/^ATE_\d{1,2}_DU$/.test(regra)) {
    throw new Error('regra_vencimento inválida. Use DIA_N (ex.: DIA_5), ATE_N_DU (ex.: ATE_5_DU) ou MANUAL.');
  }
}

// ------------------------------------------------------------
//  GERAÇÃO DO MÊS — idempotente, mesmo padrão de
//  gerarDespesasDoMes(). Lança em Recebimentos com tipo
//  'Receita Recorrente', para não misturar com recebimento de
//  guia/particular (que já usam essa aba com outro tipo).
// ------------------------------------------------------------
function gerarReceitasRecorrentesDoMes(ano, mes, usuario) { // mes: 0-11
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recorrentes = getReceitasRecorrentes(usuario).filter(r => r.ativo === 'SIM');
    const shRec = ss.getSheetByName(CONFIG.SHEETS.RECEBIMENTOS);
    const recebimentosExistentes = _getSheet(ss, CONFIG.SHEETS.RECEBIMENTOS).map(r => _toObj(RECEBIMENTOS_HEADERS_V7, r));
    const mesTexto = _mesDoDate(new Date(ano, mes, 1));
    const refInicioMes = new Date(ano, mes, 1);
    const refFimMes = new Date(ano, mes+1, 0);

    const geradas = [], puladas = [];
    recorrentes.forEach(r => {
      if (r.data_inicio && new Date(r.data_inicio) > refFimMes) { puladas.push(r.descricao + ' (ainda não começou)'); return; }
      if (r.data_fim && new Date(r.data_fim) < refInicioMes) { puladas.push(r.descricao + ' (já encerrada)'); return; }
      const jaExiste = recebimentosExistentes.some(rec => rec.referencia_id === r.id && rec.mes === mesTexto);
      if (jaExiste) { puladas.push(r.descricao + ' (já gerada este mês)'); return; }

      const vencimento = _calcularVencimentoRecorrente(r.regra_vencimento, ano, mes);
      const id = 'RECB' + new Date().getTime() + Math.floor(Math.random()*1000);
      shRec.appendRow([id, vencimento?_dateStr(vencimento):'', 'Receita Recorrente', r.id, '', r.pagador_nome, _sanNum(r.valor_padrao), r.forma_pgto||'', mesTexto, r.descricao, new Date()]);
      geradas.push({ descricao:r.descricao, valor:_sanNum(r.valor_padrao), vencimento: vencimento?_dateStr(vencimento):'(sem data automática)' });
    });

    _log(usuario.nome, 'GERAR_RECEITAS_RECORRENTES_MES', `${mesTexto}/${ano}: ${geradas.length} geradas, ${puladas.length} puladas`);
    return { ok:true, mes:mesTexto, ano, geradas, puladas };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Resumo simples do mês, usado no dashboard e na tela de Faturamento
// (aba "Outras Receitas") para mostrar o total já confirmado.
function getResumoReceitasRecorrentes(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const mes = (filtros && filtros.mes) || 'todos';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let recebimentos = _getSheet(ss, CONFIG.SHEETS.RECEBIMENTOS).map(r => _toObj(RECEBIMENTOS_HEADERS_V7, r))
      .filter(r => r.tipo === 'Receita Recorrente');
    if (mes !== 'todos') recebimentos = recebimentos.filter(r => r.mes === mes);
    const total = recebimentos.reduce((s,r)=>s+_sanNum(r.valor),0);
    return { ok:true, total, qtd: recebimentos.length, itens: recebimentos };
  } catch(e) { return { ok:false, total:0, qtd:0, itens:[], msg:e.toString() }; }
}
