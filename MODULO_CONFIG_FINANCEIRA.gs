// ============================================================
//  MODULO_CONFIG_FINANCEIRA.gs — v2.0
//
//  Armazena, na aba "Config" já existente (colunas: chave, valor),
//  configurações editáveis pela tela (sem precisar mexer em código):
//
//   chave = 'precos_pacote'      → JSON com preço bruto por serviço x pacote
//                                  (serviços "simples": sessão avulsa,
//                                  pacote de 5, pacote de 10 etc.)
//   chave = 'matriz_preco_srv'   → JSON com preços de serviços que têm mais
//                                  de 2 dimensões (hoje: Pilates, que varia
//                                  por plano × frequência semanal). Ver
//                                  getMatrizPrecoServico() abaixo.
//   chave = 'taxas_cartao'       → JSON com taxas de cada maquininha
//                                  (débito + crédito por parcela + taxa de
//                                  antecipação), usado para calcular quanto
//                                  realmente cai na conta depois do desconto
//                                  da maquininha.
//   chave = 'taxa_guia_por_faixa'→ JSON com a tabela de taxa por faixa de
//                                  sessões (1-5, 6-15, 16-25, 26-35) e a
//                                  lista de profissionais aos quais ela se
//                                  aplica. Ver seção dedicada abaixo —
//                                  IMPORTANTE: esta é só a CONFIGURAÇÃO.
//                                  O efeito financeiro real (a taxa desconta
//                                  da comissão da profissional, ou é retida
//                                  à parte pela clínica antes de calcular a
//                                  comissão sobre o restante?) AINDA NÃO
//                                  está implementado em
//                                  calcularComissaoGuia() — depende de
//                                  confirmação da gestora. Ver aviso na
//                                  função calcularEfeitoTaxaGuiaPorFaixa().
//
//  Este módulo é INDEPENDENTE dos módulos existentes — não altera
//  nenhuma função, aba ou comportamento já em produção. Se a chave
//  ainda não existir na planilha, os getters devolvem um objeto vazio
//  (o front-end trata isso mostrando os campos zerados para preencher).
// ============================================================

function _getConfigValor(chave) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!sh || sh.getLastRow() < 2) return null;
  const dados = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0] === chave) return dados[i][1];
  }
  return null;
}

function _setConfigValor(chave, valorStr) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
  const dados = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,2).getValues() : [];
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0] === chave) { sh.getRange(i+2,2).setValue(valorStr); return; }
  }
  sh.appendRow([chave, valorStr]);
}

// ------------------------------------------------------------
//  PREÇOS POR PACOTE — Serviço x Tipo de Pacote = valor bruto
// ------------------------------------------------------------
function getPrecosPacote(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao','fisioterapeuta']);
    const raw = _getConfigValor('precos_pacote');
    return { ok:true, dados: raw ? JSON.parse(raw) : {} };
  } catch(e) { return { ok:false, dados:{}, msg:e.toString() }; }
}

function salvarPrecosPacote(precos, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _setConfigValor('precos_pacote', JSON.stringify(precos||{}));
    _log(usuario.nome, 'CONFIG_PRECOS_PACOTE', 'Tabela de preços atualizada');
    return { ok:true };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ------------------------------------------------------------
//  TAXAS DE CARTÃO — por maquininha: débito + crédito por parcela
//  + taxa extra de antecipação
// ------------------------------------------------------------
function getTaxasCartao(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao','fisioterapeuta']);
    const raw = _getConfigValor('taxas_cartao');
    return { ok:true, dados: raw ? JSON.parse(raw) : {} };
  } catch(e) { return { ok:false, dados:{}, msg:e.toString() }; }
}

function salvarTaxasCartao(taxas, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _setConfigValor('taxas_cartao', JSON.stringify(taxas||{}));
    _log(usuario.nome, 'CONFIG_TAXAS_CARTAO', 'Taxas de cartão atualizadas');
    return { ok:true };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ------------------------------------------------------------
//  MATRIZ DE PREÇO POR SERVIÇO — para serviços com mais de uma
//  dimensão de preço (hoje: Pilates, que varia por
//  plano [mensal/trimestral/semestral] × frequência semanal [1x/2x/3x],
//  separado por Pilates Clínica vs. Pilates Casa/Home).
//
//  Formato salvo: { "PILATES": { "MENSAL": {"1x":220,"2x":300,"3x":420},
//                                 "TRIMESTRAL": {...}, "SEMESTRAL": {...} },
//                    "PILATES-CASA": { ... } }
//
//  Isso é separado de 'precos_pacote' de propósito: forçar Pilates na
//  grade simples de Serviço×Pacote (PACOTES_PADRAO) exigiria criar
//  pacotes artificiais tipo "MENSAL-1X", "MENSAL-2X" etc., o que
//  rapidamente fica confuso na tela. Aqui cada serviço "matricial"
//  declara suas próprias dimensões.
// ------------------------------------------------------------
const VALOR_PADRAO_MATRIZ_PILATES = {
  'PILATES': {
    MENSAL:     {'1x':220,'2x':300,'3x':420},
    TRIMESTRAL: {'1x':210,'2x':285,'3x':400},
    SEMESTRAL:  {'1x':200,'2x':265,'3x':380}
  },
  'PILATES-CASA': {
    MENSAL:     {'1x':210,'2x':264,'3x':310},
    TRIMESTRAL: {'1x':200,'2x':253,'3x':242},
    SEMESTRAL:  {'1x':null,'2x':242,'3x':232}
  }
};

function getMatrizPrecoServico(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao','fisioterapeuta']);
    const raw = _getConfigValor('matriz_preco_srv');
    // se ainda não foi configurado manualmente, devolve os valores reais
    // da tabela "VALORES_SERVIÇOS — PARA VALIDAÇÃO" como ponto de partida,
    // mas SEM gravar nada ainda — só grava quando a gestora salvar pela tela.
    return { ok:true, dados: raw ? JSON.parse(raw) : VALOR_PADRAO_MATRIZ_PILATES };
  } catch(e) { return { ok:false, dados:{}, msg:e.toString() }; }
}

function salvarMatrizPrecoServico(matriz, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _setConfigValor('matriz_preco_srv', JSON.stringify(matriz||{}));
    _log(usuario.nome, 'CONFIG_MATRIZ_PRECO_SERVICO', 'Matriz de preço (Pilates) atualizada');
    return { ok:true };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ------------------------------------------------------------
//  TAXA DE GUIA POR FAIXA DE SESSÕES — aplicável a profissionais com
//  remuneração fixa que também atendem guia de convênio (hoje, pelo
//  PDF de regras: Yasmin Silva, Maria Aparecida, Grace Kelly, Daniela
//  — todas com "Fixo paga guia = SIM" em Regras_Comissao).
//
//  Formato salvo: { faixas: [{de:1,ate:5,valor:90}, {de:6,ate:15,valor:140},
//                             {de:16,ate:25,valor:250}, {de:26,ate:35,valor:360}],
//                    profissionais_ids: ['PROF009','PROF016','PROF017','PROF018'] }
// ------------------------------------------------------------
const FAIXAS_PADRAO_TAXA_GUIA = [
  {de:1, ate:5,  valor:90},
  {de:6, ate:15, valor:140},
  {de:16,ate:25, valor:250},
  {de:26,ate:35, valor:360}
];

function getTaxaGuiaPorFaixa(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao','fisioterapeuta']);
    const raw = _getConfigValor('taxa_guia_por_faixa');
    if (raw) return { ok:true, dados: JSON.parse(raw) };
    // ponto de partida com a tabela informada pela gestora + as 4
    // profissionais já marcadas com "Fixo paga guia = SIM" no backend.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shRegras = ss.getSheetByName(CONFIG.SHEETS.REGRAS_COMISSAO);
    let profsSugeridos = [];
    if (shRegras && shRegras.getLastRow() > 1) {
      const dados = shRegras.getRange(2,1,shRegras.getLastRow()-1,shRegras.getLastColumn()).getValues();
      const colFixoPagaGuia = 5; // índice 5 = 'fixo_recebe_comissao_guia' na ordem de REGRAS_COMISSAO
      // ordem real (ver Código_Principal.gs, migrarV6): profissional_id(0),
      // profissional_nome(1), tipo_vinculo(2), tipo_remuneracao(3),
      // valor_fixo_mensal(4), fixo_recebe_comissao_guia(5), ...
      profsSugeridos = dados.filter(r => String(r[colFixoPagaGuia]).toUpperCase() === 'SIM').map(r => r[0]);
    }
    return { ok:true, dados: { faixas: FAIXAS_PADRAO_TAXA_GUIA, profissionais_ids: profsSugeridos } };
  } catch(e) { return { ok:false, dados:{faixas:[],profissionais_ids:[]}, msg:e.toString() }; }
}

function salvarTaxaGuiaPorFaixa(config, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    _setConfigValor('taxa_guia_por_faixa', JSON.stringify(config||{faixas:[],profissionais_ids:[]}));
    _log(usuario.nome, 'CONFIG_TAXA_GUIA_FAIXA', 'Configuração de taxa de guia por faixa atualizada');
    return { ok:true };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Calcula, para uma dada guia (já com itens carregados), qual seria a
// taxa aplicável pela tabela de faixas — SEM aplicar o efeito financeiro
// ainda. Usado hoje só para exibir "referência" na tela de configuração
// e na guia individual, não para debitar de ninguém.
//
// ATENÇÃO — PENDENTE DE CONFIRMAÇÃO DA GESTORA:
// esta função retorna o valor da taxa pela quantidade de sessões da guia,
// mas NÃO decide se esse valor desconta da comissão da profissional ou
// se é retido à parte pela clínica antes de calcular a comissão sobre o
// restante. Enquanto isso não for confirmado, NENHUMA outra função do
// sistema (calcularComissaoGuia, getResumoComissoesPorProfissional etc.)
// deve chamar esta função para alterar valores reais — ela existe apenas
// para a tela de configuração poder mostrar "quanto seria a taxa" como
// referência.
function calcularTaxaGuiaPorFaixaReferencia(totalSessoes, usuario) {
  try {
    const cfgRes = getTaxaGuiaPorFaixa(usuario);
    if (!cfgRes.ok) return { ok:false, valor:0, msg:cfgRes.msg };
    const faixas = cfgRes.dados.faixas || [];
    const qtd = parseInt(totalSessoes) || 0;
    const faixa = faixas.find(f => qtd >= f.de && qtd <= f.ate);
    return { ok:true, valor: faixa ? faixa.valor : 0, faixa_aplicada: faixa || null };
  } catch(e) { return { ok:false, valor:0, msg:e.toString() }; }
}
