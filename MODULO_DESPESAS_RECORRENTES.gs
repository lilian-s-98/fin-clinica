// ============================================================
//  MODULO_DESPESAS_RECORRENTES.gs — v6.0
//
//  Cadastro de despesas que se repetem todo mês (aluguel, folha,
//  impostos, empréstimos etc.) na aba "Despesas_Recorrentes".
//  Uma função gera automaticamente o lançamento do mês em "Despesas"
//  — hoje e também para meses futuros, pra você já ver a projeção
//  de fluxo de caixa e lucro antes mesmo do mês chegar.
//
//  IMPORTANTE (o ponto que você levantou): recorrente ≠ fixo.
//  Por isso o campo "tipo" separa os dois:
//   tipo = 'FIXO'     -> mesmo valor todo mês (ex.: aluguel, contabilidade).
//          O sistema já lança o valor certo automaticamente.
//   tipo = 'VARIAVEL' -> se repete todo mês mas o valor muda (ex.: água,
//          luz). O sistema lança a despesa do mês com valor R$ 0,00 e
//          status "A Confirmar" — você só edita o valor real na aba
//          Despesas quando a conta chegar. Isso evita esquecer de
//          lançar, mas também evita "chutar" um valor errado na projeção.
//
//  Colunas de Despesas_Recorrentes:
//   id, descricao, fornecedor, categoria, tipo (FIXO/VARIAVEL),
//   valor_padrao (usado quando tipo=FIXO; ignorado quando VARIAVEL),
//   regra_vencimento: 'DIA_N' (dia fixo do mês, ex. DIA_20) ou
//                     'ATE_N_DU' (até o Nº dia útil do mês, ex. ATE_5_DU)
//                     ou 'MANUAL' (sem data automática — ex. contas que
//                     variam de dia, você preenche na hora)
//   forma_pgto, data_inicio (opcional), data_fim (opcional — despesa
//     que tem prazo pra acabar, ex. empréstimo com parcelas contadas),
//   ativo, observacao, criado_por, criado_em
// ============================================================

const DESPESAS_RECORRENTES_HEADERS = ['id','descricao','fornecedor','categoria','tipo','valor_padrao','regra_vencimento','forma_pgto','data_inicio','data_fim','ativo','observacao','criado_por','criado_em'];
const DESPESAS_HEADERS_V6 = ['id','data','mes','categoria','descricao','fornecedor','valor','forma_pgto','tipo','status','data_vencimento','data_pgto','comprovante_url','observacao','lancado_por','criado_em','origem_recorrente_id'];

// ------------------------------------------------------------
//  SEED — sua lista de despesas fixas/recorrentes, já pronta.
//  Rodada automaticamente por migrarV6() (só na primeira vez —
//  se a aba já tiver alguma linha, ela pula e não duplica nada).
//  ATENÇÃO — revise estas linhas antes de gerar o primeiro mês:
//   • A primeira despesa da sua lista veio SEM DESCRIÇÃO no texto que
//     você mandou (só o valor R$ 9.141,00). Criei como
//     "DESPESA_SEM_NOME — CONFERIR" — edite o nome correto na planilha.
//   • IPTU: você disse "até novembro" — assumi novembro DESTE ano
//     (2026). Se for outro ano, ajuste data_fim.
//   • PRONAMPE e Empréstimo Banco Nordeste: você mencionou parcelas
//     restantes (36 meses / reajuste anual) mas não a data exata de
//     início/fim — deixei data_fim em branco. Preencha a data que a
//     última parcela vence, pra deixar de gerar sozinho depois.
//   • Título Caixa: "22 meses restantes" a partir de agora (achei a
//     data aproximada). Título Banese: "até dez/2027" (usei essa data).
//   • Água e Luz: cadastradas como VARIAVEL/MANUAL — todo mês o sistema
//     cria a despesa com valor R$0,00 esperando você preencher.
// ------------------------------------------------------------
function _popularDespesasRecorrentesPadrao(ss, log) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.DESPESAS_RECORRENTES);
  if (!sh) { log.push('AVISO: aba Despesas_Recorrentes não encontrada.'); return; }
  if (sh.getLastRow() > 1) { log.push('Despesas_Recorrentes já tinha dados — seed não foi rodado de novo (nada foi duplicado).'); return; }

  const agora = new Date();
  const SEED = [
    // descricao, fornecedor, categoria, tipo, valor_padrao, regra_vencimento, forma_pgto, data_inicio, data_fim, observacao
    ['DESPESA_SEM_NOME — CONFERIR', '', 'A definir', 'FIXO', 9141,   'ATE_5_DU', '', '', '', 'Veio sem descrição na lista original — identifique e renomeie.'],
    ['CEMED',                        '', 'Serviço',   'FIXO', 130,    'DIA_1',   '', '', '', ''],
    ['SASE',                         '', 'Serviço',   'FIXO', 180,    'DIA_3',   '', '', '', ''],
    ['FGTS',                         '', 'Encargo',   'FIXO', 625,    'DIA_20',  '', '', '', ''],
    ['INSS',                         '', 'Encargo',   'FIXO', 738.64, 'DIA_20',  '', '', '', ''],
    ['ALUGUEL CASA',                 '', 'Aluguel',   'FIXO', 2000,   'ATE_5_DU','', '', '', ''],
    ['TISS',                         '', 'Serviço',   'FIXO', 55,     'DIA_30',  '', '', '', ''],
    ['IPTU',                         '', 'Imposto',   'FIXO', 638,    'DIA_8',   '', '2026-11-30', 'Você disse "até novembro" — confirme o ano/mês final.'],
    ['SIMPLES NACIONAL',             '', 'Imposto',   'FIXO', 1200,   'DIA_30',  '', '', '', ''],
    ['CONTABILIDADE',                '', 'Serviço',   'FIXO', 1621,   'DIA_15',  '', '', '', ''],
    ['SIMPLES NACIONAL PARCELA',     '', 'Imposto',   'FIXO', 155,    'DIA_30',  '', '', '', ''],
    ['ALUGUEL STUDIO',               '', 'Aluguel',   'FIXO', 3000,   'DIA_20',  '', '', '', ''],
    ['CONDOMINIO',                   '', 'Aluguel',   'FIXO', 730,    'DIA_10',  '', '', '', ''],
    ['PRONAMPE (EMPRÉSTIMO)',        '', 'Empréstimo','FIXO', 2805.50,'DIA_14',  '', '', '', '36 parcelas com reajuste no próximo ano — preencha data_fim com a data da última parcela.'],
    ['EMPRESTIMO BANCO NORDESTE',    '', 'Empréstimo','FIXO', 662.06, 'DIA_16',  '', '', '', '36 parcelas, reajuste anual — preencha data_fim.'],
    ['TITULO CAPITALIZAÇÃO CAIXA',   '', 'Outros',    'FIXO', 200,    'DIA_10',  '', '', '2028-06-30', '22 parcelas restantes informadas em ago/2026 — confirme a data final exata.'],
    ['TITULO CAPITALIZAÇÃO BANESE',  '', 'Outros',    'FIXO', 200,    'DIA_15',  '', '', '2027-12-31', ''],
    ['VIVO TELEFONE E INTERNET',     '', 'Serviço',   'FIXO', 295,    'DIA_25',  '', '', '', ''],
    ['ÁGUA',                         '', 'Utilidade', 'VARIAVEL', 0,  'MANUAL',  '', '', '', 'Valor muda todo mês — preencha na hora que a conta chegar.'],
    ['LUZ',                          '', 'Utilidade', 'VARIAVEL', 0,  'MANUAL',  '', '', '', 'Valor muda todo mês — preencha na hora que a conta chegar.'],
    ['SOCIAL MEDIA',                 '', 'Serviço',   'FIXO', 950,    'ATE_5_DU','', '', '', '']
  ];

  SEED.forEach(row => {
    const id = 'DR' + Utilities.getUuid().slice(0,8).toUpperCase();
    sh.appendRow([id, row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]||'', row[8]||'', 'SIM', row[9]||'', 'SISTEMA (migração v6)', agora]);
  });
  log.push(SEED.length + ' despesas recorrentes cadastradas a partir da sua lista. REVISE a aba Despesas_Recorrentes antes de gerar o primeiro mês (veja os comentários no topo do arquivo Modulo_Despesas_Recorrentes.gs).');
}

// ------------------------------------------------------------
//  CRUD
// ------------------------------------------------------------
function getDespesasRecorrentes(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    return _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.DESPESAS_RECORRENTES).map(r => _toObj(DESPESAS_RECORRENTES_HEADERS, r));
  } catch(e) { return []; }
}

function salvarDespesaRecorrente(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _validarDespesaRecorrente(dados);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.DESPESAS_RECORRENTES);
    const allData = sh.getDataRange().getValues();
    const linha = [
      dados.id || ('DR'+Utilities.getUuid().slice(0,8).toUpperCase()), _san(dados.descricao), _san(dados.fornecedor),
      _san(dados.categoria)||'Outros', dados.tipo, dados.tipo==='FIXO' ? _sanNum(dados.valor_padrao) : 0,
      dados.regra_vencimento, _san(dados.forma_pgto), dados.data_inicio||'', dados.data_fim||'',
      dados.ativo||'SIM', _san(dados.observacao), usuario.nome, dados.id ? allData.find(r=>r[0]===dados.id)[13] : new Date()
    ];
    if (dados.id) {
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === dados.id) { sh.getRange(i+1,1,1,linha.length).setValues([linha]); _log(usuario.nome,'EDIT_DESPESA_RECORRENTE',dados.descricao); return {ok:true}; }
      }
    }
    sh.appendRow(linha);
    _log(usuario.nome, 'NOVA_DESPESA_RECORRENTE', dados.descricao);
    return { ok:true, id: linha[0] };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function _validarDespesaRecorrente(dados) {
  if (!dados.descricao) throw new Error('Descreva a despesa (ex.: "Aluguel Studio").');
  if (!['FIXO','VARIAVEL'].includes(dados.tipo)) throw new Error('tipo precisa ser FIXO ou VARIAVEL.');
  if (dados.tipo === 'FIXO' && !_sanNum(dados.valor_padrao)) throw new Error('Despesa FIXA precisa de um valor_padrao > 0.');
  const regra = String(dados.regra_vencimento||'');
  if (regra !== 'MANUAL' && !/^DIA_\d{1,2}$/.test(regra) && !/^ATE_\d{1,2}_DU$/.test(regra)) {
    throw new Error('regra_vencimento inválida. Use DIA_N (ex.: DIA_20), ATE_N_DU (ex.: ATE_5_DU) ou MANUAL.');
  }
}

// ------------------------------------------------------------
//  CALENDÁRIO — Nº dia útil do mês (reaproveita _ehDiaUtil do
//  Modulo_Convenios_Faturamento.gs)
// ------------------------------------------------------------
function _diaUtilNDoMes(ano, mes, n) { // mes: 0-11
  let d = new Date(ano, mes, 1);
  let contador = _ehDiaUtil(d) ? 1 : 0;
  while (contador < n) { d.setDate(d.getDate()+1); if (_ehDiaUtil(d)) contador++; }
  return d;
}

function _calcularVencimentoRecorrente(regra, ano, mes) {
  if (regra === 'MANUAL' || !regra) return null;
  const mDia = regra.match(/^DIA_(\d{1,2})$/);
  if (mDia) {
    const dia = parseInt(mDia[1]);
    const ultimoDiaMes = new Date(ano, mes+1, 0).getDate();
    return new Date(ano, mes, Math.min(dia, ultimoDiaMes)); // poka-yoke: "DIA_30" em fevereiro não estoura o mês
  }
  const mDu = regra.match(/^ATE_(\d{1,2})_DU$/);
  if (mDu) return _diaUtilNDoMes(ano, mes, parseInt(mDu[1]));
  return null;
}

// ------------------------------------------------------------
//  GERAÇÃO DO MÊS — idempotente: não duplica se você rodar de novo.
// ------------------------------------------------------------
function gerarDespesasDoMes(ano, mes, usuario) { // mes: 0-11
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recorrentes = getDespesasRecorrentes(usuario).filter(r => r.ativo === 'SIM');
    const shDesp = ss.getSheetByName(CONFIG.SHEETS.DESPESAS);
    const despesasExistentes = _getSheet(ss, CONFIG.SHEETS.DESPESAS).map(r => _toObj(DESPESAS_HEADERS_V6, r));
    const mesTexto = _mesDoDate(new Date(ano, mes, 1));
    const refInicioMes = new Date(ano, mes, 1);
    const refFimMes = new Date(ano, mes+1, 0);

    const geradas = [], puladas = [];
    recorrentes.forEach(r => {
      if (r.data_inicio && new Date(r.data_inicio) > refFimMes) { puladas.push(r.descricao + ' (ainda não começou)'); return; }
      if (r.data_fim && new Date(r.data_fim) < refInicioMes) { puladas.push(r.descricao + ' (já encerrada)'); return; }
      // poka-yoke: já existe lançamento desta recorrente para este mês? não duplica.
      const jaExiste = despesasExistentes.some(d => d.origem_recorrente_id === r.id && d.mes === mesTexto);
      if (jaExiste) { puladas.push(r.descricao + ' (já gerada este mês)'); return; }

      const vencimento = _calcularVencimentoRecorrente(r.regra_vencimento, ano, mes);
      const id = 'DESP' + new Date().getTime() + Math.floor(Math.random()*1000);
      const valor = r.tipo === 'FIXO' ? _sanNum(r.valor_padrao) : 0;
      const status = r.tipo === 'FIXO' ? 'Pendente' : 'A Confirmar';
      shDesp.appendRow([id, vencimento?_dateStr(vencimento):'', mesTexto, r.categoria, r.descricao, r.fornecedor, valor, r.forma_pgto||'', r.tipo, status, vencimento?_dateStr(vencimento):'', '', '', r.observacao||'', usuario.nome, new Date(), r.id]);
      geradas.push({ descricao:r.descricao, valor, vencimento: vencimento?_dateStr(vencimento):'(preencher manualmente)', status });
    });

    _log(usuario.nome, 'GERAR_DESPESAS_MES', `${mesTexto}/${ano}: ${geradas.length} geradas, ${puladas.length} puladas`);
    return { ok:true, mes:mesTexto, ano, geradas, puladas };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Gera o mês atual + N meses à frente de uma vez (ex.: qtdMeses=6 gera até
// 6 meses no futuro) — é isso que alimenta a projeção de fluxo de caixa.
function gerarProjecaoDespesasFuturas(qtdMeses, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const hoje = new Date();
    const resultado = [];
    for (let i = 0; i <= qtdMeses; i++) {
      const d = new Date(hoje.getFullYear(), hoje.getMonth()+i, 1);
      resultado.push(gerarDespesasDoMes(d.getFullYear(), d.getMonth(), usuario));
    }
    return { ok:true, resultado };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ------------------------------------------------------------
//  PROJEÇÃO DE FLUXO DE CAIXA — junta despesas (já geradas, incluindo
//  futuras) + recebimentos previstos de convênio (Modulo_Convenios_
//  Faturamento.gs) + uma ESTIMATIVA de particulares (média dos últimos
//  3 meses fechados — deixo claro que é estimativa, não garantia).
// ------------------------------------------------------------
function getProjecaoFluxoCaixa(qtdMeses, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const hoje = new Date();

    // média de particulares dos últimos 3 meses FECHADOS (não conta o mês atual, que está incompleto)
    const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
    let somaParticular = 0;
    for (let i = 1; i <= 3; i++) {
      const d = new Date(hoje.getFullYear(), hoje.getMonth()-i, 1);
      const mesTx = MESES[d.getMonth()];
      somaParticular += getParticulares({mes:mesTx}).reduce((s,p)=>s+_sanNum(p.valor),0);
    }
    const mediaParticularMensal = somaParticular / 3;

    const meses = [];
    for (let i = 0; i <= qtdMeses; i++) {
      const d = new Date(hoje.getFullYear(), hoje.getMonth()+i, 1);
      const mesTx = MESES[d.getMonth()];
      const despesasMes = getDespesas({mes:mesTx}).filter(x => new Date(x.data||x.data_vencimento).getFullYear() === d.getFullYear());
      const totalDespesas = despesasMes.reduce((s,x)=>s+_sanNum(x.valor),0);
      const qtdVariavelSemValor = despesasMes.filter(x => x.status === 'A Confirmar').length;

      const guiasMes = getGuias({mes:mesTx}).filter(g => g.status !== 'Pago');
      const receitaConvenioPrevista = guiasMes.reduce((s,g)=>s+_sanNum(g.valor_total)-_sanNum(g.valor_glosado),0);
      const receitaConvenioJaPaga = getGuias({mes:mesTx}).filter(g=>g.status==='Pago').reduce((s,g)=>s+_sanNum(g.valor_total)-_sanNum(g.valor_glosado),0);

      const receitaTotal = receitaConvenioPrevista + receitaConvenioJaPaga + mediaParticularMensal;
      const lucroProjetado = receitaTotal - totalDespesas;

      meses.push({
        mes: mesTx, ano: d.getFullYear(),
        despesasPrevistas: totalDespesas,
        qtdDespesasSemValorDefinido: qtdVariavelSemValor,
        receitaConvenioPrevista: receitaConvenioPrevista + receitaConvenioJaPaga,
        receitaParticularEstimada: mediaParticularMensal,
        receitaTotalEstimada: receitaTotal,
        lucroProjetado
      });
    }

    return { ok:true, meses, mediaParticularMensal, aviso: 'Receita de particulares é uma ESTIMATIVA baseada na média dos últimos 3 meses fechados — não é garantida. Despesas do tipo VARIAVEL sem valor confirmado entram como R$0,00 até você preencher.' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}
