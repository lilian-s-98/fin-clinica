// ============================================================
//  MODULO_CONVENIOS_FATURAMENTO.gs — v6.0
//
//  Regras de cada convênio (dia de faturamento, prazo de envio,
//  prazo de recebimento) ficam na aba "Regras_Convenio" — nada
//  fica preso no código. Para mudar o prazo de um convênio,
//  edite a linha dele na planilha.
//
//  Colunas de Regras_Convenio:
//   convenio_id, convenio_nome,
//   tipo_faturamento: 'FIXO' | 'CONDICIONAL' | 'ULTIMO_DIA_UTIL' | 'IMEDIATO'
//   dia_fixo: dia do mês (usado quando tipo_faturamento=FIXO)
//   dia_corte: dia limite usado quando tipo_faturamento=CONDICIONAL
//              (ex.: CASSÍ — se o atendimento ocorreu até o dia 10,
//              fatura no dia 10; senão fatura no dia 2 do mês seguinte)
//   prazo_envio_dias_uteis: dias úteis entre faturamento e envio do lote
//   prazo_recebimento_dias: dias corridos entre envio e recebimento previsto
//   atraso_frequente: SIM/NÃO — convênio conhecido por atrasar (usado só
//                     para não gerar alerta prematuro na conferência)
//   ativo, observacao
// ============================================================

const REGRAS_CONVENIO_HEADERS = ['convenio_id','convenio_nome','tipo_faturamento','dia_fixo','dia_corte','prazo_envio_dias_uteis','prazo_recebimento_dias','atraso_frequente','ativo','observacao'];

// Dados de partida (v6.0) extraídos das regras informadas pelo cliente.
// Usados SOMENTE por _popularRegrasConvenioPadrao() na migração — depois
// disso, a fonte da verdade passa a ser a planilha, não este objeto.
const _REGRAS_CONVENIO_PADRAO = {
  'GEAP':      { tipo_faturamento:'FIXO', dia_fixo:5,  dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:90, atraso_frequente:'NÃO' },
  'AMIL':      { tipo_faturamento:'FIXO', dia_fixo:15, dia_corte:'', prazo_envio_dias_uteis:0,  prazo_recebimento_dias:30, atraso_frequente:'NÃO', observacao:'Sessão única' },
  'CASSÍ':     { tipo_faturamento:'CONDICIONAL', dia_fixo:'', dia_corte:10, prazo_envio_dias_uteis:0, prazo_recebimento_dias:30, atraso_frequente:'NÃO' },
  'CASSIND':   { tipo_faturamento:'ULTIMO_DIA_UTIL', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'CAPESESP':  { tipo_faturamento:'ULTIMO_DIA_UTIL', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'BLUE':      { tipo_faturamento:'ULTIMO_DIA_UTIL', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'PETROBRAS': { tipo_faturamento:'FIXO', dia_fixo:20, dia_corte:'', prazo_envio_dias_uteis:0, prazo_recebimento_dias:30, atraso_frequente:'NÃO' },
  'ASSEC':     { tipo_faturamento:'ULTIMO_DIA_UTIL', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'CASSE':     { tipo_faturamento:'ULTIMO_DIA_UTIL', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'CASEC':     { tipo_faturamento:'FIXO', dia_fixo:5, dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'CASSIND2':  { tipo_faturamento:'FIXO', dia_fixo:5, dia_corte:'', prazo_envio_dias_uteis:10, prazo_recebimento_dias:60, atraso_frequente:'SIM' },
  'PARTICULAR':{ tipo_faturamento:'IMEDIATO', dia_fixo:'', dia_corte:'', prazo_envio_dias_uteis:0, prazo_recebimento_dias:0, atraso_frequente:'NÃO', observacao:'Recebimento imediato, fora das regras de atraso de convênio' }
};

function _popularRegrasConvenioPadrao(ss, log) {
  const shConv = ss.getSheetByName(CONFIG.SHEETS.CONVENIOS);
  const shRegras = ss.getSheetByName(CONFIG.SHEETS.REGRAS_CONVENIO);
  if (!shConv || !shRegras) return;
  const convenios = shConv.getLastRow() > 1 ? shConv.getRange(2,1,shConv.getLastRow()-1,shConv.getLastColumn()).getValues() : [];
  const existentes = shRegras.getLastRow() > 1 ? shRegras.getRange(2,1,shRegras.getLastRow()-1,1).getValues().map(r=>r[0]) : [];
  convenios.forEach(c => {
    const id = c[0], nome = c[1];
    if (existentes.indexOf(id) !== -1) { log.push('Regra de convênio já existia (mantida): ' + nome); return; }
    const padrao = _REGRAS_CONVENIO_PADRAO[String(nome).toUpperCase()];
    if (padrao) {
      shRegras.appendRow([id, nome, padrao.tipo_faturamento, padrao.dia_fixo, padrao.dia_corte, padrao.prazo_envio_dias_uteis, padrao.prazo_recebimento_dias, padrao.atraso_frequente, 'SIM', padrao.observacao||'']);
      log.push('Regra de convênio criada a partir do padrão: ' + nome);
    } else {
      // convênio sem correspondência no padrão: cria linha em branco para preenchimento manual
      // (poka-yoke: melhor aparecer vazio e óbvio do que assumir um prazo errado silenciosamente)
      shRegras.appendRow([id, nome, '', '', '', '', c[2]||'', 'NÃO', 'NÃO', 'PREENCHER MANUALMENTE — convênio novo, sem regra padrão conhecida']);
      log.push('AVISO: convênio "' + nome + '" não tem regra padrão. Linha em branco criada — preencha manualmente.');
    }
  });
}

function getRegrasConvenio(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    return _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.REGRAS_CONVENIO).map(r => _toObj(REGRAS_CONVENIO_HEADERS, r));
  } catch(e) { return []; }
}

function _getRegraConvenio(convenio_id) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.REGRAS_CONVENIO);
  const row = rows.find(r => r[0] === convenio_id);
  return row ? _toObj(REGRAS_CONVENIO_HEADERS, row) : null;
}

function salvarRegraConvenio(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (!['FIXO','CONDICIONAL','ULTIMO_DIA_UTIL','IMEDIATO'].includes(dados.tipo_faturamento)) {
      throw new Error('tipo_faturamento inválido. Use FIXO, CONDICIONAL, ULTIMO_DIA_UTIL ou IMEDIATO.');
    }
    if (dados.tipo_faturamento === 'FIXO' && !dados.dia_fixo) throw new Error('Faturamento FIXO precisa do dia_fixo (1 a 31).');
    if (dados.tipo_faturamento === 'CONDICIONAL' && !dados.dia_corte) throw new Error('Faturamento CONDICIONAL precisa do dia_corte.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGRAS_CONVENIO);
    const allData = sh.getDataRange().getValues();
    const linha = [dados.convenio_id, _san(dados.convenio_nome), dados.tipo_faturamento, dados.dia_fixo||'', dados.dia_corte||'', parseInt(dados.prazo_envio_dias_uteis)||0, parseInt(dados.prazo_recebimento_dias)||0, dados.atraso_frequente||'NÃO', dados.ativo||'SIM', _san(dados.observacao)];
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === dados.convenio_id) { sh.getRange(i+1,1,1,linha.length).setValues([linha]); _log(usuario.nome,'EDIT_REGRA_CONVENIO',dados.convenio_id); return {ok:true}; }
    }
    sh.appendRow(linha);
    _log(usuario.nome,'NOVA_REGRA_CONVENIO',dados.convenio_id);
    return {ok:true};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ------------------------------------------------------------
//  CALENDÁRIO — dias úteis (considera só sáb/dom; se quiser
//  feriados nacionais, adicione uma lista e cheque aqui)
// ------------------------------------------------------------
function _ehDiaUtil(date) {
  const dia = date.getDay();
  return dia !== 0 && dia !== 6;
}

function _somarDiasUteis(dataBase, quantidade) {
  const d = new Date(dataBase);
  let restante = quantidade;
  while (restante > 0) {
    d.setDate(d.getDate() + 1);
    if (_ehDiaUtil(d)) restante--;
  }
  return d;
}

function _ultimoDiaUtilDoMes(ano, mes) { // mes: 0-11
  const d = new Date(ano, mes + 1, 0); // último dia do mês
  while (!_ehDiaUtil(d)) d.setDate(d.getDate() - 1);
  return d;
}

// ------------------------------------------------------------
//  Data de faturamento a partir da regra do convênio + data de
//  referência (normalmente a data do atendimento).
// ------------------------------------------------------------
function calcularDataFaturamento(convenio_id, dataReferencia) {
  const regra = _getRegraConvenio(convenio_id);
  if (!regra) return { ok:false, msg:'Convênio sem regra cadastrada em Regras_Convenio.' };
  const ref = new Date(dataReferencia);
  let dataFat;
  switch (regra.tipo_faturamento) {
    case 'IMEDIATO':
      dataFat = ref;
      break;
    case 'FIXO': {
      const dia = parseInt(regra.dia_fixo);
      dataFat = new Date(ref.getFullYear(), ref.getMonth(), dia);
      if (dataFat < ref) dataFat = new Date(ref.getFullYear(), ref.getMonth()+1, dia);
      break;
    }
    case 'CONDICIONAL': {
      const corte = parseInt(regra.dia_corte);
      // se o atendimento ocorreu até o dia de corte, fatura no próprio dia de corte
      // deste mês; senão, fatura no dia 2 do mês seguinte.
      if (ref.getDate() <= corte) dataFat = new Date(ref.getFullYear(), ref.getMonth(), corte);
      else dataFat = new Date(ref.getFullYear(), ref.getMonth()+1, 2);
      break;
    }
    case 'ULTIMO_DIA_UTIL':
      dataFat = _ultimoDiaUtilDoMes(ref.getFullYear(), ref.getMonth());
      break;
    default:
      return { ok:false, msg:'tipo_faturamento desconhecido: ' + regra.tipo_faturamento };
  }
  return { ok:true, data_faturamento: _dateStr(dataFat), regra };
}

// Data limite de envio do lote e data prevista de recebimento,
// a partir da data de faturamento já calculada.
function calcularProjecaoGuia(convenio_id, dataReferencia) {
  const fat = calcularDataFaturamento(convenio_id, dataReferencia);
  if (!fat.ok) return fat;
  const regra = fat.regra;
  const dataEnvio = regra.prazo_envio_dias_uteis > 0
    ? _somarDiasUteis(new Date(fat.data_faturamento), parseInt(regra.prazo_envio_dias_uteis))
    : new Date(fat.data_faturamento);
  const dataPrevRecebimento = new Date(dataEnvio);
  dataPrevRecebimento.setDate(dataPrevRecebimento.getDate() + parseInt(regra.prazo_recebimento_dias||0));
  return {
    ok:true,
    data_faturamento: fat.data_faturamento,
    data_envio_limite: _dateStr(dataEnvio),
    data_prevista_recebimento: _dateStr(dataPrevRecebimento),
    atraso_frequente: regra.atraso_frequente === 'SIM'
  };
}

// ------------------------------------------------------------
//  PROJEÇÃO DE RECEBIMENTO — todas as guias pendentes com a data
//  prevista e o valor, para alimentar o fluxo de caixa futuro.
// ------------------------------------------------------------
function getProjecaoRecebimentos(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const guias = getGuias(filtros).filter(g => g.status === 'Pendente' || g.status === 'Enviada' || g.status === '');
    const hoje = new Date();
    const projecao = guias.map(g => {
      let dataPrev = g.data_prev_pgto ? new Date(g.data_prev_pgto) : null;
      if (!dataPrev && g.data_envio) {
        const calc = calcularProjecaoGuia(g.convenio_id, g.data);
        if (calc.ok) dataPrev = new Date(calc.data_prevista_recebimento);
      }
      const dias_restantes = dataPrev ? Math.round((dataPrev - hoje)/86400000) : null;
      return {
        guia_id: g.id, convenio_nome: g.convenio_nome, paciente_nome: g.paciente_nome,
        valor_total: _sanNum(g.valor_total), data_prevista: dataPrev ? _dateStr(dataPrev) : '(sem data de envio informada)',
        dias_restantes, situacao: dias_restantes === null ? 'SEM PREVISÃO' : dias_restantes < 0 ? 'ATRASADO' : dias_restantes <= 7 ? 'PRÓXIMO' : 'NO PRAZO'
      };
    }).sort((a,b) => (a.dias_restantes ?? 999) - (b.dias_restantes ?? 999));

    const totalProjetado = projecao.reduce((s,p)=>s+p.valor_total,0);
    const totalAtrasado = projecao.filter(p=>p.situacao==='ATRASADO').reduce((s,p)=>s+p.valor_total,0);
    return { ok:true, projecao, totalProjetado, totalAtrasado };
  } catch(e) { return { ok:false, projecao:[], msg:e.toString() }; }
}

// ------------------------------------------------------------
//  CONFERÊNCIA DE PRAZOS — cruza o que foi recebido com o que
//  estava previsto, sinalizando atraso real (não apenas projetado).
//  Convênios com atraso_frequente=SIM ganham uma tolerância extra
//  de 10 dias antes de aparecerem como "EM ATRASO CRÍTICO", para não
//  gerar alarme falso todo mês com quem sempre atrasa um pouco.
// ------------------------------------------------------------
function conferirPrazosRecebimento(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const guias = getGuias(filtros);
    const hoje = new Date();
    const TOLERANCIA_ATRASO_FREQUENTE_DIAS = 10;

    const resultado = guias.filter(g => g.status !== 'Pago' && g.status !== 'Cancelada').map(g => {
      const regra = _getRegraConvenio(g.convenio_id);
      const dataPrev = g.data_prev_pgto ? new Date(g.data_prev_pgto) : null;
      if (!dataPrev) return { ...g, conferencia: 'SEM DATA PREVISTA — configure a Regra de Convênio ou informe data_envio.' };
      const diasAtraso = Math.round((hoje - dataPrev)/86400000);
      const tolerancia = (regra && regra.atraso_frequente === 'SIM') ? TOLERANCIA_ATRASO_FREQUENTE_DIAS : 0;
      let conferencia;
      if (diasAtraso <= tolerancia) conferencia = 'NO PRAZO';
      else if (diasAtraso <= tolerancia + 15) conferencia = 'EM ATRASO';
      else conferencia = 'EM ATRASO CRÍTICO';
      return { guia_id:g.id, convenio_nome:g.convenio_nome, valor_total:_sanNum(g.valor_total), data_prevista:_dateStr(dataPrev), dias_atraso: Math.max(diasAtraso,0), conferencia };
    });

    return { ok:true, resultado };
  } catch(e) { return { ok:false, resultado:[], msg:e.toString() }; }
}

// ------------------------------------------------------------
//  RESUMO POR CONVÊNIO — faturado, recebido e pendente separados
// ------------------------------------------------------------
function getResumoPorConvenio(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const guias = getGuias(filtros);
    const mapa = {};
    guias.forEach(g => {
      if (!mapa[g.convenio_nome]) mapa[g.convenio_nome] = { convenio_nome:g.convenio_nome, faturado:0, recebido:0, pendente:0, glosado:0, qtd_guias:0 };
      const m = mapa[g.convenio_nome];
      const valor = _sanNum(g.valor_total);
      m.faturado += valor;
      m.glosado += _sanNum(g.valor_glosado);
      m.qtd_guias++;
      if (g.status === 'Pago') m.recebido += valor - _sanNum(g.valor_glosado);
      else m.pendente += valor - _sanNum(g.valor_glosado);
    });
    return { ok:true, dados: Object.values(mapa).sort((a,b)=>b.faturado-a.faturado) };
  } catch(e) { return { ok:false, dados:[], msg:e.toString() }; }
}
