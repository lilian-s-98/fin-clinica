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
//     por vírgula, ex: "FISIO.M,DTM"), percentual_a
//   categoria_b_nome, categoria_b_codigos, percentual_b
//   ativo
//
//  Regra de repasse (PJ variável):
//    comissao_item = valor_item * (% de conclusão do item) * percentual/100
//  onde percentual vem de categoria_a ou categoria_b, dependendo do
//  código do serviço do item (categoria_a_codigos = "Todos" cobre tudo).
// ============================================================

const REGRAS_COMISSAO_HEADERS = ['profissional_id','profissional_nome','tipo_vinculo','tipo_remuneracao','valor_fixo_mensal','fixo_recebe_comissao_guia','variacao_particular_convenio','percentual_particular','variacao_por_servico','categoria_a_nome','categoria_a_codigos','percentual_a','categoria_b_nome','categoria_b_codigos','percentual_b','ativo'];

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
      _san(dados.categoria_a_nome)||'Todos', _san(dados.categoria_a_codigos)||'Todos', _sanNum(dados.percentual_a),
      _san(dados.categoria_b_nome), _san(dados.categoria_b_codigos), _sanNum(dados.percentual_b),
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
  if (dados.tipo_remuneracao === 'FIXO' && !_sanNum(dados.valor_fixo_mensal)) {
    throw new Error('Profissional com remuneração FIXO precisa de um valor_fixo_mensal > 0.');
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
  sh.appendRow([id, nome, tipoVinculo, 'VARIAVEL', 0, 'NÃO', 'NÃO', 0, 'NÃO', 'Todos', 'Todos', 0, '', '', 0, 'SIM']);
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
  if (regra.variacao_por_servico !== 'SIM') return { percentual: regra.percentual_a || 0, categoria: regra.categoria_a_nome };
  if (emLista(regra.categoria_b_codigos, codigo)) return { percentual: regra.percentual_b || 0, categoria: regra.categoria_b_nome };
  // padrão cai na categoria A (que normalmente é "Todos")
  return { percentual: regra.percentual_a || 0, categoria: regra.categoria_a_nome };
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
      const comissaoItem = valorItem * pctConclusao * (cat.percentual/100);
      comissaoProf += comissaoItem;
      return { item_id: it.id, codigo: it.codigo, valor_item: valorItem, pct_conclusao: pctConclusao, categoria: cat.categoria, percentual: cat.percentual, comissao_item: comissaoItem };
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
  let percentual;
  if (regra.variacao_particular_convenio === 'SIM') {
    percentual = regra.percentual_particular;
  } else {
    const cat = _categoriaDoCodigo(regra, particular.servico_id);
    percentual = cat.percentual;
  }
  return { ok:true, comissao: valor * (percentual/100), percentual };
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
