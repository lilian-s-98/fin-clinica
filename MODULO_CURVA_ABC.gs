// ============================================================
//  MODULO_CURVA_ABC.gs — v1.0
//
//  Curva ABC (Pareto) de faturamento em três dimensões:
//   - por código de procedimento (usa Itens_Guia.codigo — os códigos
//     lançados nas guias de convênio)
//   - por serviço (usa o nome do serviço em Guias + Particulares)
//   - por profissional (reaproveita getDesempenhoProfissionais, só
//     adiciona a classificação A/B/C)
//
//  Cortes padrão (Pareto clássico), CONFIRMAR COM A GESTORA:
//   A = itens que somam até 80% do faturamento acumulado
//   B = de 80% até 95%
//   C = de 95% até 100%
//   Os cortes são configuráveis via parâmetro `cortes` (opcional);
//   se não informado, usa os padrões acima.
//
//  Este módulo é independente — só LÊ dados de Itens_Guia, Guias e
//  Particulares (já existentes), não escreve nada novo.
// ============================================================

const CORTES_ABC_PADRAO = { A: 80, B: 95 };

function _classificarABC(itensOrdenados, cortes) {
  const c = cortes || CORTES_ABC_PADRAO;
  const total = itensOrdenados.reduce((s, it) => s + it.valor, 0);
  let acumulado = 0;
  return itensOrdenados.map(it => {
    acumulado += it.valor;
    const pctAcumulado = total > 0 ? (acumulado / total) * 100 : 0;
    const classe = pctAcumulado <= c.A ? 'A' : (pctAcumulado <= c.B ? 'B' : 'C');
    return { ...it, pct_participacao: total > 0 ? (it.valor / total) * 100 : 0, pct_acumulado: pctAcumulado, classe };
  });
}

// ------------------------------------------------------------
//  CURVA ABC POR CÓDIGO DE PROCEDIMENTO
//  (código lançado em cada item de guia — ex: FISIO.M, RPG, DTM)
// ------------------------------------------------------------
function getCurvaABCPorCodigo(filtros, usuario, cortes) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itensRows = _getSheet(ss, CONFIG.SHEETS.ITENS_GUIA);
    const guiasHeaders = ['id','data','mes','convenio_id','convenio_nome','paciente_id','paciente_nome','profissional_id','profissional_nome','lote','protocolo','num_nf','valor_total','prazo_dias','data_envio','data_prev_pgto','data_pgto_real','status','valor_glosado','observacao','lancado_por','criado_em','lote_id'];
    const guias = _getSheet(ss, CONFIG.SHEETS.GUIAS).map(r => _toObj(guiasHeaders, r));
    const guiaPorId = {};
    guias.forEach(g => { guiaPorId[g.id] = g; });

    const mes = (filtros && filtros.mes) || 'todos';
    const itensHeaders = ['id','guia_id','convenio_nome','codigo','descricao','quantidade','valor_unitario','valor_total','profissional_id','profissional_nome','concluido'];
    const itens = itensRows.map(r => _toObj(itensHeaders, r)).filter(it => {
      if (mes === 'todos') return true;
      const g = guiaPorId[it.guia_id];
      return g && g.mes === mes;
    });

    const porCodigo = {};
    itens.forEach(it => {
      const cod = it.codigo || '(sem código)';
      if (!porCodigo[cod]) porCodigo[cod] = { codigo: cod, descricao: it.descricao || '', valor: 0, qtd_lancamentos: 0 };
      porCodigo[cod].valor += _sanNum(it.valor_total);
      porCodigo[cod].qtd_lancamentos += 1;
    });

    const lista = Object.values(porCodigo).sort((a,b) => b.valor - a.valor);
    return { ok:true, dados: _classificarABC(lista, cortes) };
  } catch(e) { return { ok:false, dados:[], msg:e.toString() }; }
}

// ------------------------------------------------------------
//  CURVA ABC POR SERVIÇO
//  (soma Guias + Particulares por nome de serviço/especialidade)
// ------------------------------------------------------------
function getCurvaABCPorServico(filtros, usuario, cortes) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const mes = (filtros && filtros.mes) || 'todos';
    const particulares = getParticulares({mes});
    const guias = getGuias({mes});

    const porServico = {};
    particulares.forEach(p => {
      const nome = p.servico_nome || '(sem serviço)';
      if (!porServico[nome]) porServico[nome] = { servico: nome, valor: 0, qtd_lancamentos: 0, origem_particular: 0, origem_convenio: 0 };
      porServico[nome].valor += _sanNum(p.valor);
      porServico[nome].origem_particular += _sanNum(p.valor);
      porServico[nome].qtd_lancamentos += 1;
    });
    // guias não têm "servico_nome" direto (o serviço vem do código do item),
    // então soma por convenio_nome como aproximação de "linha de serviço"
    // quando o item não tiver descrição própria — mantém granularidade da
    // curva por código (função acima) para o detalhe fino.
    guias.forEach(g => {
      const nome = 'Convênio — ' + (g.convenio_nome || '(sem convênio)');
      if (!porServico[nome]) porServico[nome] = { servico: nome, valor: 0, qtd_lancamentos: 0, origem_particular: 0, origem_convenio: 0 };
      porServico[nome].valor += _sanNum(g.valor_total);
      porServico[nome].origem_convenio += _sanNum(g.valor_total);
      porServico[nome].qtd_lancamentos += 1;
    });

    const lista = Object.values(porServico).sort((a,b) => b.valor - a.valor);
    return { ok:true, dados: _classificarABC(lista, cortes) };
  } catch(e) { return { ok:false, dados:[], msg:e.toString() }; }
}

// ------------------------------------------------------------
//  CURVA ABC POR PROFISSIONAL
//  (reaproveita getDesempenhoProfissionais, adiciona classe A/B/C)
// ------------------------------------------------------------
function getCurvaABCPorProfissional(filtros, usuario, cortes) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const desemp = getDesempenhoProfissionais(filtros, usuario);
    if (!desemp.ok) return { ok:false, dados:[], msg: desemp.msg };
    const lista = desemp.dados.map(p => ({ profissional_id: p.id, profissional_nome: p.nome, valor: p.valor_total, qtd_lancamentos: p.qtd_particular + p.qtd_convenio }))
      .sort((a,b) => b.valor - a.valor);
    return { ok:true, dados: _classificarABC(lista, cortes) };
  } catch(e) { return { ok:false, dados:[], msg:e.toString() }; }
}

// ------------------------------------------------------------
//  RESUMO CONSOLIDADO — as 3 curvas de uma vez, para a tela dedicada
//  de Curva ABC não precisar fazer 3 chamadas separadas.
// ------------------------------------------------------------
function getCurvaABCConsolidada(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const porCodigo = getCurvaABCPorCodigo(filtros, usuario);
    const porServico = getCurvaABCPorServico(filtros, usuario);
    const porProfissional = getCurvaABCPorProfissional(filtros, usuario);
    return {
      ok: true,
      porCodigo: porCodigo.ok ? porCodigo.dados : [],
      porServico: porServico.ok ? porServico.dados : [],
      porProfissional: porProfissional.ok ? porProfissional.dados : [],
      cortes: CORTES_ABC_PADRAO
    };
  } catch(e) { return { ok:false, porCodigo:[], porServico:[], porProfissional:[], msg:e.toString() }; }
}
