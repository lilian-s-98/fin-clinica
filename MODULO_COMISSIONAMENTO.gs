// ============================================================
//  MODULO_COMISSIONAMENTO.gs  — v6.0
//
//  Toda regra de comissão fica na aba "Regras_Comissao" (uma linha
//  por profissional). Nada de percentual fixo dentro do código —
//  para mudar uma taxa, edite a célula na planilha e pronto.
//
//  Colunas de Regras_Comissao:
//   profissional_id, profissional_nome, tipo_vinculo (PJ/CLT),
//   tipo_remuneracao (FIXO/VARIAVEL),
//   valor_fixo_mensal          -> só usado se tipo_remuneracao=FIXO
//   fixo_recebe_comissao_guia  -> SIM/NÃO: profissional FIXO também
//                                 recebe % por guia/particular além do fixo?
//   variacao_particular_convenio -> SIM/NÃO: usa percentual_particular
//                                    diferente do percentual de convênio?
//   percentual_particular      -> % aplicado em atendimentos particulares
//                                 quando variacao_particular_convenio=SIM
//   variacao_por_servico       -> SIM/NÃO: separa em categoria A / B
//   categoria_a_nome, categoria_a_codigos (códigos de serviço separados
//     por vírgula, ex: "FISIO.M,DTM"), percentual_a, unidade_a
//   categoria_b_nome, categoria_b_codigos, percentual_b, unidade_b
//   ativo
//
//   unidade_a / unidade_b: '%' (percentual sobre o valor do item, padrão)
//     ou 'R$' (valor FIXO por atendimento/sessão, multiplicado pela
//     quantidade — usado por alguns profissionais do PDF que recebem,
//     por exemplo, "R$ 40,00 por sessão de Fisio Motora" em vez de %).
//
//  Regra de repasse (PJ variável), quando unidade='%':
//    comissao_item = valor_item * (% de conclusão do item) * percentual/100
//  quando unidade='R$':
//    comissao_item = percentual_a (aqui usado como R$/unidade) * quantidade * pctConclusao
//  Exemplo do enunciado (%): guia de R$1000, profissional concluiu 50%
//  dos procedimentos, comissão acordada 70% => 1000*0.5*0.7 = R$350.
// ============================================================

const REGRAS_COMISSAO_HEADERS = ['profissional_id','profissional_nome','tipo_vinculo','tipo_remuneracao','valor_fixo_mensal','fixo_recebe_comissao_guia','variacao_particular_convenio','percentual_particular','variacao_por_servico','categoria_a_nome','categoria_a_codigos','percentual_a','unidade_a','categoria_b_nome','categoria_b_codigos','percentual_b','unidade_b','ativo'];

function _shRegrasComissao() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGRAS_COMISSAO);
}

function getRegrasComissao(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    return _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.REGRAS_COMISSAO).map(r => _toObj(REGRAS_COMISSAO_HEADERS, r));
  } catch(e) { return []; }
}

// Busca a regra de UM profissional (uso interno, sem checagem de perfil —
// é chamada durante o cálculo de comissão de qualquer lançamento).
function _getRegraComissao(profissional_id) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.REGRAS_COMISSAO);
  const row = rows.find(r => r[0] === profissional_id);
  return row ? _toObj(REGRAS_COMISSAO_HEADERS, row) : null;
}

function salvarRegraComissao(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _validarRegraComissao(dados); // poka-yoke — lança erro claro se algo estiver inconsistente
    const sh = _shRegrasComissao();
    const allData = sh.getDataRange().getValues();
    const linha = [
      dados.profissional_id, _san(dados.profissional_nome), dados.tipo_vinculo||'PJ',
      dados.tipo_remuneracao||'VARIAVEL', _sanNum(dados.valor_fixo_mensal), dados.fixo_recebe_comissao_guia||'NÃO',
      dados.variacao_particular_convenio||'NÃO', _sanNum(dados.percentual_particular),
      dados.variacao_por_servico||'NÃO',
      _san(dados.categoria_a_nome)||'Todos', _san(dados.categoria_a_codigos)||'Todos', _sanNum(dados.percentual_a), dados.unidade_a==='R$'?'R$':'%',
      _san(dados.categoria_b_nome), _san(dados.categoria_b_codigos), _sanNum(dados.percentual_b), dados.unidade_b==='R$'?'R$':'%',
      dados.ativo||'SIM'
    ];
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === dados.profissional_id) {
        sh.getRange(i+1,1,1,linha.length).setValues([linha]);
        _log(usuario.nome, 'EDIT_REGRA_COMISSAO', dados.profissional_id);
        return {ok:true};
      }
    }
    sh.appendRow(linha);
    _log(usuario.nome, 'NOVA_REGRA_COMISSAO', dados.profissional_id);
    return {ok:true};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// Poka-yoke de cadastro: impede regra inconsistente ANTES de gravar.
function _validarRegraComissao(dados) {
  if (!dados.profissional_id) throw new Error('Selecione o profissional.');
  const pct = v => v === '' || v === undefined || v === null ? 0 : parseFloat(v);
  ['percentual_particular','percentual_a','percentual_b'].forEach(campo => {
    const v = pct(dados[campo]);
    if (v < 0 || v > 100) throw new Error(`Campo ${campo} precisa estar entre 0 e 100 (recebido: ${dados[campo]}).`);
  });
  if (dados.tipo_remuneracao === 'FIXO' && dados.tipo_vinculo === 'PJ' && !_sanNum(dados.valor_fixo_mensal)) {
    throw new Error('Profissional PJ com remuneração FIXO precisa de um valor_fixo_mensal > 0. (CLT não precisa — salário é tratado fora deste sistema.)');
  }
  if (dados.variacao_por_servico === 'SIM' && (!dados.categoria_a_codigos || !dados.categoria_b_codigos)) {
    throw new Error('Marcou "variação por serviço" mas faltou preencher os códigos da categoria A e/ou B.');
  }
}

// Cria uma linha-rascunho de 0% quando um profissional novo é cadastrado, para
// que ele NUNCA fique sem regra (comissão calculada como zero sem ninguém perceber).
function _garantirRascunhoRegraComissao(id, nome, tipoVinculo) {
  const sh = _shRegrasComissao();
  if (!sh) return;
  const allData = sh.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) if (allData[i][0] === id) return; // já existe
  sh.appendRow([id, nome, tipoVinculo, 'VARIAVEL', 0, 'NÃO', 'NÃO', 0, 'NÃO', 'Todos', 'Todos', 0, '%', '', '', 0, '%', 'SIM']);
  _criarNotificacao('config_pendente', `Configurar comissão de ${nome}`,
    `Profissional ${nome} foi cadastrado com regra de comissão em 0%. Acesse Regras de Comissão e configure o percentual correto antes de fechar o próximo faturamento.`,
    'gestor', 'sistema');
}

// Decide se um código de serviço pertence à categoria A ou B da regra.
// "Todos" na lista de códigos = coringa (cobre qualquer código).
function _categoriaDoCodigo(regra, codigo) {
  const emLista = (lista, cod) => {
    if (!lista) return false;
    const itens = String(lista).split(',').map(s => s.trim().toUpperCase());
    return itens.includes('TODOS') || itens.includes(String(cod).toUpperCase());
  };
  if (regra.variacao_por_servico !== 'SIM') return { percentual: regra.percentual_a || 0, categoria: regra.categoria_a_nome, unidade: regra.unidade_a || '%' };
  if (emLista(regra.categoria_b_codigos, codigo)) return { percentual: regra.percentual_b || 0, categoria: regra.categoria_b_nome, unidade: regra.unidade_b || '%' };
  // padrão cai na categoria A (que normalmente é "Todos")
  return { percentual: regra.percentual_a || 0, categoria: regra.categoria_a_nome, unidade: regra.unidade_a || '%' };
}

// Calcula a comissão de UM item já com a categoria resolvida — trata tanto
// percentual (%) quanto valor fixo por atendimento (R$/unidade * quantidade).
function _comissaoPorUnidade(cat, valorItem, quantidade, pctConclusao) {
  if (cat.unidade === 'R$') return cat.percentual * (quantidade||1) * pctConclusao; // aqui "percentual" guarda o R$/unidade
  return valorItem * pctConclusao * (cat.percentual/100);
}

// % de conclusão de um item de guia: SIM=100%, NÃO=0%.
// (Se no futuro quiser conclusão parcial em %, é só trocar este helper
// para ler uma coluna numérica em vez do SIM/NÃO.)
function _pctConclusaoItem(item) {
  return (item.concluido === 'SIM' || item.concluido === true) ? 1 : 0;
}

// ------------------------------------------------------------
//  CÁLCULO PRINCIPAL — Comissão de uma guia de convênio
//  Regra de repasse (PJ, exceto FIXO sem "recebe comissão de guia"):
//    comissao_item = valor_item * pctConclusao(item) * percentual/100
//  Exemplo do enunciado: guia de R$1000, profissional concluiu 50%
//  dos procedimentos, comissão acordada 70% => 1000*0.5*0.7 = R$350.
// ------------------------------------------------------------
function calcularComissaoGuia(guia_id) {
  const itens = getItensGuia(guia_id);
  if (itens.length === 0) return { ok:false, msg:'Guia sem itens', total:0, detalhe:[] };

  // agrupa por profissional (normalmente 1, mas suporta mais de 1 por guia)
  const porProfissional = {};
  itens.forEach(it => {
    if (!porProfissional[it.profissional_id]) porProfissional[it.profissional_id] = [];
    porProfissional[it.profissional_id].push(it);
  });

  const detalhe = [];
  let total = 0;
  Object.keys(porProfissional).forEach(profId => {
    const regra = _getRegraComissao(profId);
    if (!regra || regra.ativo !== 'SIM') {
      detalhe.push({ profissional_id: profId, comissao: 0, aviso: 'SEM REGRA DE COMISSÃO CADASTRADA — configure em Regras_Comissao.' });
      return;
    }
    if (regra.tipo_remuneracao === 'FIXO' && regra.fixo_recebe_comissao_guia !== 'SIM') {
      detalhe.push({ profissional_id: profId, comissao: 0, aviso: 'Profissional com remuneração FIXA (não recebe % por guia).' });
      return;
    }
    let comissaoProf = 0;
    const itensDoProf = porProfissional[profId].map(it => {
      const cat = _categoriaDoCodigo(regra, it.codigo);
      const pctConclusao = _pctConclusaoItem(it);
      const valorItem = _sanNum(it.valor_total);
      const comissaoItem = _comissaoPorUnidade(cat, valorItem, _sanNum(it.quantidade,1), pctConclusao);
      comissaoProf += comissaoItem;
      return { item_id: it.id, codigo: it.codigo, valor_item: valorItem, pct_conclusao: pctConclusao, categoria: cat.categoria, percentual: cat.percentual, unidade: cat.unidade, comissao_item: comissaoItem };
    });
    total += comissaoProf;
    detalhe.push({ profissional_id: profId, comissao: comissaoProf, itens: itensDoProf });
  });

  return { ok:true, guia_id, total, detalhe };
}

// Comissão de um atendimento particular (Particulares) — usa
// percentual_particular se variacao_particular_convenio=SIM,
// senão cai na mesma lógica de categoria A/B usada nas guias.
function calcularComissaoParticular(particular) {
  const regra = _getRegraComissao(particular.profissional_id);
  if (!regra || regra.ativo !== 'SIM') return { ok:false, comissao:0, aviso:'Sem regra de comissão cadastrada.' };
  if (regra.tipo_remuneracao === 'FIXO' && regra.fixo_recebe_comissao_guia !== 'SIM') {
    return { ok:true, comissao:0, aviso:'Remuneração fixa — não recebe % por atendimento.' };
  }
  const valor = _sanNum(particular.valor);
  const qtd = _sanNum(particular.quantidade,1);
  if (regra.variacao_particular_convenio === 'SIM') {
    // percentual_particular é sempre % (não há variante R$ fixo para particular no PDF)
    return { ok:true, comissao: valor * (regra.percentual_particular/100), percentual: regra.percentual_particular };
  }
  const cat = _categoriaDoCodigo(regra, particular.servico_id);
  const comissao = _comissaoPorUnidade(cat, valor, qtd, 1);
  return { ok:true, comissao, percentual: cat.percentual, unidade: cat.unidade };
}

// ============================================================
//  POPULAÇÃO INICIAL (chamada por migrarV6 em Codigo_Principal.gs)
//  Extraída do PDF "REGRAS_POR_PROFISSIONAL". Só grava profissionais
//  que AINDA NÃO têm linha em Regras_Comissao — nunca sobrescreve o
//  que você já editou manualmente.
//
//  ATENÇÃO — revise depois de rodar, principalmente:
//   • PROF001 (Bruno) e PROF019 (Viviane): assumi que "ambulatório" e
//     "Fisio. Motora" correspondem ao código FISIO.M, e "especializado"/
//     "Home Care" aos códigos LIB/HOME. Confirme se está certo.
//   • PROF014 (Juliana Archimino): o PDF não trouxe um código de serviço
//     para ela. Deixei "Todos" — troque pelo código real assim que
//     cadastrar o serviço "Coluna" em Serviços/Códigos.
//   • Linhas com unidade_a/unidade_b = "R$" (Viviane, Luciana, Brena,
//     Vinicius) pagam valor FIXO por atendimento/sessão, não percentual
//     — a fórmula já trata isso, mas confira os valores na planilha.
// ============================================================
function _popularRegrasComissaoPadrao(ss, log) {
  const shProf = ss.getSheetByName(CONFIG.SHEETS.PROFISSIONAIS);
  const shRegras = ss.getSheetByName(CONFIG.SHEETS.REGRAS_COMISSAO);
  if (!shProf || !shRegras) return;

  const existentes = shRegras.getLastRow() > 1 ? shRegras.getRange(2,1,shRegras.getLastRow()-1,1).getValues().map(r=>r[0]) : [];

  // dados do PDF — chave = id do profissional (mesmo id usado na aba Profissionais)
  const PADRAO = {
    PROF001: ['BRUNO NASCIMENTO','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Ambulatório','FISIO.M',40,'%','Especializado','LIB',60,'%','SIM'],
    PROF002: ['LETICIA HELEN','CLT','FIXO',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',0,'%','','',0,'%','SIM'],
    PROF003: ['LUCAS REZENDE','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',80,'%','','',0,'%','SIM'],
    PROF004: ['MARIANA MENDONÇA','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',80,'%','','',0,'%','SIM'],
    PROF005: ['MILA PIRES','PJ','VARIAVEL',0,'NÃO','NÃO',50,'NÃO','Todos','Todos',60,'%','','',0,'%','SIM'],
    PROF006: ['MILENA MORAES','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Acup. Clínica','ACUP.IN',50,'%','Acup. Socorro','ACUP.EX',60,'%','SIM'],
    PROF007: ['OSMALÍ SILVA','PJ','FIXO',2200,'NÃO','NÃO',0,'NÃO','Todos','Todos',0,'%','','',0,'%','SIM'],
    PROF008: ['YANNA MENEZES','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Clínica','Todos',50,'%','','',0,'%','SIM'],
    PROF009: ['YASMIN SILVA','PJ','VARIAVEL',0,'SIM','NÃO',0,'NÃO','Todos','Todos',63,'%','','',0,'%','SIM'],
    PROF010: ['LUIZA GABRIELA','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Pilates','PILATES,PILATESCASA',40,'%','Quiro/RPG/DTM','QUIRO,RPG,DTM',50,'%','SIM'],
    PROF011: ['JAMILLE GONÇALVES','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',60,'%','','',0,'%','SIM'],
    PROF012: ['JULIANA LINHARES','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',60,'%','','',0,'%','SIM'],
    PROF013: ['JOSI NASCIMENTO','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Fisio. Motora','FISIO.M',40,'%','DTM','DTM',50,'%','SIM'],
    PROF014: ['JULIANA ARCHIMINO','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos (REVISAR CÓDIGO)','Todos',50,'%','','',0,'%','SIM'],
    PROF015: ['LAIS MARINHO','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','Todos',50,'%','','',0,'%','SIM'],
    PROF016: ['MARIA APARECIDA','PJ','VARIAVEL',0,'SIM','NÃO',0,'NÃO','Todos','Todos',60,'%','','',0,'%','SIM'],
    PROF017: ['GRACE KELLY','PJ','VARIAVEL',0,'SIM','NÃO',0,'NÃO','Todos','Todos',63,'%','','',0,'%','SIM'],
    PROF018: ['DANIELA','PJ','VARIAVEL',0,'SIM','NÃO',0,'NÃO','Todos','Todos',63,'%','','',0,'%','SIM'],
    PROF019: ['VIVIANE COSTA','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Fisio. Motora','FISIO.M',40,'R$','Home Care','HOME',60,'%','SIM'],
    PROF020: ['LUCIANA','PJ','VARIAVEL',0,'NÃO','NÃO',0,'NÃO','Todos','FISIO.M,HOME',50,'R$','','',0,'%','SIM'],
    PROF021: ['BRENA MIRELI','PJ','VARIAVEL',0,'NÃO','SIM',50,'NÃO','Todos','NUTRI',60,'R$','','',0,'%','SIM'],
    PROF022: ['VINICIUS SOBRAL','PJ','VARIAVEL',0,'NÃO','NÃO',0,'SIM','Consulta','ORTO',50,'R$','Infiltração','INF.CORT,INF.ACIDO',60,'%','SIM']
  };

  Object.keys(PADRAO).forEach(id => {
    if (existentes.indexOf(id) !== -1) { log.push('Regra de comissão já existia (mantida): ' + id); return; }
    const p = PADRAO[id];
    const nome = p[0];
    shRegras.appendRow([id, nome, p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], p[10], p[11], p[12], p[13], p[14], p[15], p[16]]);
    log.push('Regra de comissão criada para ' + nome + ' (' + id + ') — CONFIRA os valores na planilha.');
  });

  // profissionais que existem na aba Profissionais mas não estão no PADRAO acima
  // (cadastrados depois do PDF, por exemplo) recebem o rascunho de 0% de sempre.
  const profs = shProf.getLastRow() > 1 ? shProf.getRange(2,1,shProf.getLastRow()-1,shProf.getLastColumn()).getValues() : [];
  profs.forEach(r => {
    const id = r[0];
    if (PADRAO[id] || existentes.indexOf(id) !== -1) return;
    _garantirRascunhoRegraComissao(id, r[1], r[3]);
    log.push('AVISO: profissional ' + r[1] + ' (' + id + ') sem regra conhecida — rascunho 0% criado, configure manualmente.');
  });
}

// Resumo agregado de comissões por profissional num período — alimenta o
// dashboard de "desempenho financeiro por profissional".
function getResumoComissoesPorProfissional(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const guias = getGuias(filtros);
    const particulares = getParticulares(filtros);
    const resumo = {}; // profissional_id -> {nome, comissao_guias, comissao_particulares, avisos:[]}

    guias.forEach(g => {
      const calc = calcularComissaoGuia(g.id);
      (calc.detalhe||[]).forEach(d => {
        if (!resumo[d.profissional_id]) resumo[d.profissional_id] = { profissional_id:d.profissional_id, comissao_guias:0, comissao_particulares:0, avisos:[] };
        resumo[d.profissional_id].comissao_guias += (d.comissao||0);
        if (d.aviso) resumo[d.profissional_id].avisos.push(d.aviso);
      });
    });

    particulares.forEach(p => {
      const calc = calcularComissaoParticular(p);
      if (!resumo[p.profissional_id]) resumo[p.profissional_id] = { profissional_id:p.profissional_id, comissao_guias:0, comissao_particulares:0, avisos:[] };
      resumo[p.profissional_id].comissao_particulares += (calc.comissao||0);
      if (calc.aviso) resumo[p.profissional_id].avisos.push(calc.aviso);
    });

    const profs = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.PROFISSIONAIS);
    const dados = Object.values(resumo).map(r => {
      const p = profs.find(x => x[0] === r.profissional_id);
      return { ...r, profissional_nome: p ? p[1] : r.profissional_id, comissao_total: r.comissao_guias + r.comissao_particulares, avisos: [...new Set(r.avisos)] };
    }).sort((a,b) => b.comissao_total - a.comissao_total);

    return { ok:true, dados };
  } catch(e) { return { ok:false, dados:[], msg:e.toString() }; }
}
