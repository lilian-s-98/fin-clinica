// ============================================================
//  MODULO_DASHBOARD_KPI.gs — v7.0
//  Substitui a antiga getDashboardData (mesma assinatura/uso no
//  front-end), agora incluindo: Total Faturado, Total Recebido,
//  Valor em Glosa, Valores a Receber, meta mensal + quanto falta
//  para bater a meta do mês/ano, projeção de mês em que a meta
//  anual será atingida no ritmo atual, e os dados para os gráficos
//  de faturamento por convênio, status de lotes e desempenho por
//  profissional. NÃO inclui mais nada de IA/Anthropic.
//
//  META MENSAL — DEFAULT ASSUMIDO (confirmar com a gestora):
//  CONFIG.META_ANUAL / 12, linear, sem sazonalidade. Se no futuro
//  a clínica quiser meta diferente por mês (ex.: dezembro mais
//  fraco), trocar a constante META_MENSAL_LINEAR abaixo por uma
//  tabela MES -> valor, sem precisar mexer em mais nada — o resto
//  do cálculo já usa a função getMetaDoMes() como ponto único.
// ============================================================

const MESES_ORDEM = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];

// Ponto único de definição da meta mensal. Hoje: linear (meta anual / 12).
// Troque aqui se a gestora confirmar sazonalidade — ex:
// const METAS_MENSAIS_CUSTOM = {JANEIRO:100000, ..., DEZEMBRO:80000};
// e troque o corpo da função para usar essa tabela em vez do cálculo linear.
function getMetaDoMes(mesTexto) {
  return CONFIG.META_ANUAL / 12;
}

function getDashboardData(filtros, usuario) {
  try {
    const particulares = getParticulares(filtros);
    const guias = getGuias(filtros);
    const despesas = getDespesas(filtros);

    const recParticular = particulares.reduce((s,r) => s+_sanNum(r.valor),0);
    const totalFaturadoConvenio = guias.reduce((s,r) => s+_sanNum(r.valor_total),0);
    const totalGlosado = guias.reduce((s,r) => s+_sanNum(r.valor_glosado),0);
    const totalRecebidoConvenio = guias.filter(g=>g.status==='Pago').reduce((s,r) => s+_sanNum(r.valor_total)-_sanNum(r.valor_glosado),0);
    const totalAReceber = guias.filter(g=>g.status!=='Pago').reduce((s,r)=>s+_sanNum(r.valor_total)-_sanNum(r.valor_glosado),0);
    const totalDespesas = despesas.reduce((s,r) => s+_sanNum(r.valor),0);

    const totalFaturado = recParticular + totalFaturadoConvenio;
    const totalRecebido = recParticular + totalRecebidoConvenio; // particular é sempre "recebido na hora"
    const resultado = totalRecebido - totalDespesas;

    const porMes = {};
    MESES_ORDEM.forEach(m => porMes[m] = {particular:0,convenio:0,despesa:0});
    particulares.forEach(r => { if(porMes[r.mes]) porMes[r.mes].particular += _sanNum(r.valor); });
    guias.forEach(r => { if(porMes[r.mes]) porMes[r.mes].convenio += _sanNum(r.valor_total); });
    despesas.forEach(r => { if(porMes[r.mes]) porMes[r.mes].despesa += _sanNum(r.valor); });

    // gráfico: faturamento por convênio (usa o módulo dedicado)
    const resumoConvenio = getResumoPorConvenio(filtros, usuario);

    const porProfissional = {};
    particulares.forEach(r => { porProfissional[r.profissional_nome] = (porProfissional[r.profissional_nome]||0) + _sanNum(r.valor); });

    const hoje = new Date();
    const aVencer = guias.filter(g => g.status==='Pendente' && g.data_prev_pgto)
      .map(g => ({...g, dias_restantes: Math.round((new Date(g.data_prev_pgto)-hoje)/86400000)}))
      .sort((a,b) => a.dias_restantes-b.dias_restantes).slice(0,8);

    const agendaHoje = getAgendamentos({data:_dateStr(hoje)});
    const conf = getRelatorioConformidade(7);

    // gráfico: status dos lotes no mês
    const statusLotes = getStatusLotesMes(filtros, usuario);

    // resumo de glosas
    const resumoGlosas = getResumoGlosas(filtros, usuario);

    // --- META MENSAL/ANUAL — usa o faturamento do ANO INTEIRO até agora,
    //     não só do mês filtrado, porque a meta é sempre um acumulado anual.
    const mesAtualTexto = MESES_ORDEM[hoje.getMonth()];
    const anoFiltros = { mes: 'todos' }; // sempre olha o ano todo pra meta, independente do filtro de mês da tela
    const particularesAno = getParticulares(anoFiltros).filter(r => new Date(r.data).getFullYear() === hoje.getFullYear());
    const guiasAno = getGuias(anoFiltros).filter(g => new Date(g.data).getFullYear() === hoje.getFullYear());
    const faturadoAnoAteAgora = particularesAno.reduce((s,r)=>s+_sanNum(r.valor),0) + guiasAno.reduce((s,r)=>s+_sanNum(r.valor_total),0);

    const faturadoMesAtual = porMes[mesAtualTexto] ? (porMes[mesAtualTexto].particular + porMes[mesAtualTexto].convenio) : 0;
    const metaMesAtual = getMetaDoMes(mesAtualTexto);
    const faltaParaMetaMes = Math.max(metaMesAtual - faturadoMesAtual, 0);
    const faltaParaMetaAno = Math.max(CONFIG.META_ANUAL - faturadoAnoAteAgora, 0);

    // projeção: no ritmo médio mensal do ano até agora, em que mês a meta anual seria atingida
    const mesesDecorridos = hoje.getMonth() + 1; // 1-12
    const mediaMensalAno = mesesDecorridos > 0 ? faturadoAnoAteAgora / mesesDecorridos : 0;
    let mesProjecaoMeta = null;
    if (mediaMensalAno > 0) {
      const mesesNecessarios = Math.ceil(CONFIG.META_ANUAL / mediaMensalAno);
      mesProjecaoMeta = mesesNecessarios <= 12 ? MESES_ORDEM[mesesNecessarios-1] + '/' + hoje.getFullYear()
        : 'não atinge a meta em ' + hoje.getFullYear() + ' no ritmo atual (precisaria de ' + mesesNecessarios + ' meses)';
    }

    return {
      ok: true,
      kpis: {
        totalFaturado, totalRecebido, totalGlosado, totalAReceber,
        recParticular, recConvenio: totalFaturadoConvenio, totalDespesas, resultado,
        qtdGuiasPendentes: guias.filter(g=>g.status==='Pendente').length,
        ticketMedio: particulares.length > 0 ? recParticular/particulares.length : 0,
        qtdHoje: agendaHoje.length,
        pctMeta: Math.min((faturadoAnoAteAgora/CONFIG.META_ANUAL)*100,999).toFixed(1),
        metaAnual: CONFIG.META_ANUAL,
        metaMesAtual, faturadoMesAtual, faltaParaMetaMes,
        faturadoAnoAteAgora, faltaParaMetaAno,
        mediaMensalAno, mesProjecaoMeta
      },
      porMes, porConvenio: resumoConvenio.dados||[], porProfissional, aVencer,
      agendaHoje, conformidade: conf,
      statusLotes: statusLotes.porStatus||{}, resumoGlosas
    };
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  DESEMPENHO POR PROFISSIONAL — agora incluindo comissão real
//  calculada pelo Modulo_Comissionamento.gs (não só o faturamento bruto)
// ============================================================
function getDesempenhoProfissionais(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mes = (filtros && filtros.mes) || 'todos';

    const profs = _getSheet(ss, CONFIG.SHEETS.PROFISSIONAIS).filter(r=>r[5]==='SIM')
      .map(r=>({id:r[0],nome:r[1],especialidade:r[2],cor:r[7]||'#0049AF'}));

    const particulares = getParticulares({mes});
    const guias = getGuias({mes});
    const comissoes = getResumoComissoesPorProfissional({mes}, usuario).dados || [];

    const agendaRows = _getSheet(ss, CONFIG.SHEETS.AGENDA);
    const headersAg = ['id','data','hora','hora_fim','profissional_id','profissional_nome','paciente_id','paciente_nome','servico_id','servico_nome','tipo','status','observacao','criado_por','criado_em','atualizado_em','cor_profissional','duracao_minutos'];
    const agenda = agendaRows.map(r=>_toObj(headersAg,r))
      .filter(r => mes === 'todos' || (r.data && _mesDoDate(new Date(r.data)) === mes));

    const dados = profs.map(p => {
      const pp = particulares.filter(x=>x.profissional_id===p.id);
      const gg = guias.filter(x=>x.profissional_id===p.id);
      const ag = agenda.filter(x=>x.profissional_id===p.id && x.status!=='cancelado');
      const com = comissoes.find(c => c.profissional_id === p.id);
      const valorParticular = pp.reduce((s,x)=>s+_sanNum(x.valor),0);
      const valorConvenio = gg.reduce((s,x)=>s+_sanNum(x.valor_total),0);
      const compareceu = ag.filter(x=>x.status==='compareceu').length;
      const faltou = ag.filter(x=>x.status==='faltou').length;
      const taxaComparecimento = ag.length>0 ? Math.round((compareceu/ag.length)*100) : null;
      return {
        id: p.id, nome: p.nome, especialidade: p.especialidade, cor: p.cor,
        qtd_particular: pp.length, valor_particular: valorParticular,
        qtd_convenio: gg.length, valor_convenio: valorConvenio,
        valor_total: valorParticular + valorConvenio,
        comissao_total: com ? com.comissao_total : 0,
        avisos_comissao: com ? com.avisos : [],
        ticket_medio: pp.length>0 ? valorParticular/pp.length : 0,
        qtd_atendimentos: ag.length, compareceu, faltou,
        taxa_comparecimento: taxaComparecimento
      };
    }).sort((a,b) => b.valor_total - a.valor_total);

    const totalGeral = dados.reduce((s,r)=>s+r.valor_total,0);
    dados.forEach(r => r.participacao = totalGeral>0 ? Math.round((r.valor_total/totalGeral)*1000)/10 : 0);

    return { ok: true, dados, totalGeral };
  } catch(e) { return { ok:false, dados:[], msg: e.toString() }; }
}
