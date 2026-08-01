// ============================================================
//  MODULO_CONFIG_FINANCEIRA.gs — v1.0
//
//  Armazena, na aba "Config" já existente (colunas: chave, valor),
//  duas configurações editáveis pela tela (sem precisar mexer em
//  código para atualizar preço ou taxa):
//
//   chave = 'precos_pacote'  → JSON com preço bruto por serviço x pacote
//   chave = 'taxas_cartao'   → JSON com taxas de cada maquininha
//                              (débito + crédito por parcela + taxa de
//                              antecipação), usado para calcular quanto
//                              realmente cai na conta depois do desconto
//                              da maquininha.
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
