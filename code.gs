
// ===== Código_Principal.gs =====
// ============================================================
//  INSTITUTO DA DOR — Codigo_Principal.gs  v7.0
//  Sistema de Gestão Clínica e Financeira
//  Gabillaud Consultoria | 2026
//
//  NOVIDADES v7.0:
//  - Meta anual atualizada para R$ 1.400.000 (CONFIG.META_ANUAL).
//  - Novo serviço Pilates (Solo e Casa/Home) com matriz de preços
//    própria (plano × frequência semanal), ver
//    MODULO_CONFIG_FINANCEIRA.gs (getMatrizPrecoServico).
//  - Nova aba Receitas_Recorrentes — primeira fonte de receita
//    recorrente fora de convênio/particular (ex: aluguel de espaço),
//    ver MODULO_RECEITAS_RECORRENTES.gs.
//  - Novo painel de configuração de taxa de guia por faixa de
//    sessões, para profissionais com remuneração fixa que também
//    recebem guia de convênio (ver MODULO_CONFIG_FINANCEIRA.gs).
//    ATENÇÃO: por ora esta é só a TELA DE CONFIGURAÇÃO — o efeito
//    financeiro (se desconta da comissão ou é retido à parte antes
//    dela) ainda não está plugado em calcularComissaoGuia() porque
//    depende de confirmação da gestora. Ver comentário no próprio
//    MODULO_CONFIG_FINANCEIRA.gs.
//  - Nova Curva ABC (por código de procedimento, por serviço e por
//    profissional), ver MODULO_CURVA_ABC.gs.
//  - Nova função importarHistoricoReal() — importa os lançamentos
//    particulares e despesas reais de Jan-Jul/2026 fornecidos pela
//    gestora, sem apagar nada existente e sem duplicar se rodada
//    mais de uma vez. Ver MODULO_IMPORTACAO_HISTORICO.gs.
//
//  NOVIDADES v6.0 (mantidas):
//  - Removido: Assistente de IA (consultarIA e toda a integração
//    com a API da Anthropic). Não há mais chave de API armazenada.
//  - Módulo de Comissionamento configurável por profissional
//    (Modulo_Comissionamento.gs)
//  - Módulo de Regras de Convênio / Faturamento configurável
//    (Modulo_Convenios_Faturamento.gs)
//  - Módulo de Glosas (Modulo_Glosas.gs)
//  - Módulo de Lotes/Protocolos (Modulo_Lotes.gs)
//  - Dashboard e KPIs ampliados (Modulo_Dashboard_KPI.gs)
//  - Mantido 100% do que já existia: checklist por perfil,
//    histórico, notificações, guias, particulares, despesas,
//    agenda, autenticação, log de auditoria.
//
//  ATENÇÃO — LEIA ANTES DE RODAR QUALQUER FUNÇÃO:
//  setupInicial() CONTINUA existindo apenas para uma planilha NOVA
//  e vazia (ele limpa o conteúdo das abas). Se sua planilha já tem
//  dados de produção, NÃO rode setupInicial(). Rode migrarV6() e,
//  em seguida, migrarV7() — cada uma só ADICIONA as abas/colunas
//  novas da sua versão, sem apagar nada do que já existe. Se você
//  já rodou migrarV6() antes, rode só migrarV7() agora.
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  META_ANUAL: 1400000,
  VERSION: '7.0',
  PERFIS: ['admin', 'gestor', 'recepcao', 'fisioterapeuta'],
  MAX_LOGIN_ATTEMPTS: 5,
  LOCKOUT_MINUTES: 5,
  SHEETS: {
    CONFIG:            'Config',
    USUARIOS:          'Usuários',
    PACIENTES:         'Pacientes',
    PROFISSIONAIS:     'Profissionais',
    SERVICOS:          'Serviços',
    CONVENIOS:         'Convênios',
    CODIGOS:           'Códigos',
    PARTICULARES:      'Particulares',
    GUIAS:             'Guias',
    ITENS_GUIA:        'Itens_Guia',
    DESPESAS:          'Despesas',
    RECEBIMENTOS:      'Recebimentos',
    AGENDA:            'Agenda',
    CHECKLIST_DEF:     'Checklist_Definicoes',
    CHECKLIST_HIST:    'Checklist_Historico',
    NOTIFICACOES:      'Notificações',
    LOG:               'Log',
    LOGIN_ATTEMPTS:    'Login_Attempts',
    // --- novas abas v6.0 ---
    REGRAS_CONVENIO:   'Regras_Convenio',
    REGRAS_COMISSAO:   'Regras_Comissao',
    GLOSAS:            'Glosas',
    LOTES:             'Lotes',
    DESPESAS_RECORRENTES: 'Despesas_Recorrentes',
    // --- novas abas v7.0 ---
    RECEITAS_RECORRENTES: 'Receitas_Recorrentes'
    ,PLANO_CONTAS: 'Plano_Contas'
  }
};

// ============================================================
//  PONTO DE ENTRADA WEB
// ============================================================
function doGet(e) {
  // O index não usa tags de template do Apps Script. Servi-lo diretamente
  // evita a reescrita adicional do HtmlService (document.write), que podia
  // corromper o bloco JavaScript grande e gerar "missing )" no VM do iframe.
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Instituto da Dor — Gestão v6.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  SETUP INICIAL — SOMENTE PARA PLANILHA NOVA (apaga conteúdo!)
// ============================================================
function setupInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Iniciando setup v6.0 (planilha nova)...');

  _criarAba(ss, CONFIG.SHEETS.CONFIG,         ['chave','valor']);
  _criarAba(ss, CONFIG.SHEETS.USUARIOS,       ['id','nome','email','perfil','ativo','senha','criado_em','ultimo_login','cor_avatar','cargo']);
  _criarAba(ss, CONFIG.SHEETS.PACIENTES,      ['id','nome','cpf','rg','data_nascimento','sexo','telefone','email','cep','logradouro','numero','complemento','bairro','cidade','estado','convenio_principal','num_carteirinha','observacoes','ativo','criado_em','atualizado_em']);
  _criarAba(ss, CONFIG.SHEETS.PROFISSIONAIS,  ['id','nome','especialidade','tipo_vinculo','percentual','ativo','criado_em','cor','cro_crf_crm','telefone','email']);
  _criarAba(ss, CONFIG.SHEETS.SERVICOS,       ['id','nome','categoria','valor_particular','ativo','duracao_minutos']);
  _criarAba(ss, CONFIG.SHEETS.CONVENIOS,      ['id','nome','prazo_pgto_dias','contato','ativo','email_contato','observacao']);
  _criarAba(ss, CONFIG.SHEETS.CODIGOS,        ['id','convenio_id','convenio_nome','codigo','descricao','valor','ativo']);
  _criarAba(ss, CONFIG.SHEETS.AGENDA,         ['id','data','hora','hora_fim','profissional_id','profissional_nome','paciente_id','paciente_nome','servico_id','servico_nome','tipo','status','observacao','criado_por','criado_em','atualizado_em','cor_profissional','duracao_minutos']);
  _criarAba(ss, CONFIG.SHEETS.PARTICULARES,   ['id','data','mes','paciente_id','paciente_nome','profissional_id','profissional_nome','servico_id','servico_nome','valor','forma_pgto','quantidade','tipo_qtd','observacao','status','lancado_por','criado_em','agenda_id']);
  _criarAba(ss, CONFIG.SHEETS.GUIAS,          ['id','data','mes','convenio_id','convenio_nome','paciente_id','paciente_nome','profissional_id','profissional_nome','lote','protocolo','num_nf','valor_total','prazo_dias','data_envio','data_prev_pgto','data_pgto_real','status','valor_glosado','observacao','lancado_por','criado_em','lote_id']);
  _criarAba(ss, CONFIG.SHEETS.ITENS_GUIA,     ['id','guia_id','convenio_nome','codigo','descricao','quantidade','valor_unitario','valor_total','profissional_id','profissional_nome','concluido']);
  _criarAba(ss, CONFIG.SHEETS.DESPESAS,       ['id','data','mes','categoria','descricao','fornecedor','valor','forma_pgto','tipo','status','data_vencimento','data_pgto','comprovante_url','observacao','lancado_por','criado_em','origem_recorrente_id']);
  _criarAba(ss, CONFIG.SHEETS.DESPESAS_RECORRENTES, ['id','descricao','fornecedor','categoria','tipo','valor_padrao','regra_vencimento','forma_pgto','data_inicio','data_fim','ativo','observacao','criado_por','criado_em']);
  _criarAba(ss, CONFIG.SHEETS.RECEBIMENTOS,   ['id','data','tipo','referencia_id','convenio_nome','paciente_nome','valor','forma_pgto','mes','observacao','criado_em']);
  _criarAba(ss, CONFIG.SHEETS.PLANO_CONTAS, ['id','codigo','nome','tipo','parent_id','ativo','criado_em','atualizado_em']);
  _criarAba(ss, CONFIG.SHEETS.CHECKLIST_DEF,  ['id','perfil','titulo','descricao','categoria','obrigatorio','ordem','ativo','icone']);
  _criarAba(ss, CONFIG.SHEETS.CHECKLIST_HIST, ['id','data','usuario_id','usuario_nome','perfil','checklist_id','titulo','concluido','hora_conclusao','observacao','criado_em']);
  _criarAba(ss, CONFIG.SHEETS.NOTIFICACOES,   ['id','tipo','titulo','mensagem','para_perfil','lida','criado_em','criado_por']);
  _criarAba(ss, CONFIG.SHEETS.LOG,            ['timestamp','usuario','acao','detalhes','ip']);
  _criarAba(ss, CONFIG.SHEETS.LOGIN_ATTEMPTS, ['email','tentativas','ultimo_at','bloqueado_ate']);
  // --- novas abas v6.0 ---
  _criarAba(ss, CONFIG.SHEETS.REGRAS_CONVENIO,['convenio_id','convenio_nome','tipo_faturamento','dia_fixo','dia_corte','prazo_envio_dias_uteis','prazo_recebimento_dias','atraso_frequente','ativo','observacao']);
  _criarAba(ss, CONFIG.SHEETS.REGRAS_COMISSAO,['profissional_id','profissional_nome','tipo_vinculo','tipo_remuneracao','valor_fixo_mensal','fixo_recebe_comissao_guia','variacao_particular_convenio','percentual_particular','variacao_por_servico','categoria_a_nome','categoria_a_codigos','percentual_a','unidade_a','categoria_b_nome','categoria_b_codigos','percentual_b','unidade_b','ativo']);
  _criarAba(ss, CONFIG.SHEETS.GLOSAS,         ['id','guia_id','convenio_nome','data_glosa','valor_glosado','motivo','status','observacao','lancado_por','criado_em']);
  _criarAba(ss, CONFIG.SHEETS.LOTES,          ['id','lote','convenio_id','convenio_nome','data_fechamento','data_envio','data_prev_recebimento','valor_total_lote','valor_recebido','valor_glosado','status','guias_ids','observacao','criado_em','atualizado_em']);

  _popularDadosIniciais(ss);
  Logger.log('Setup v6.0 concluído!');
  return { ok: true, msg: 'Setup v6.0 realizado com sucesso!' };
}

// ============================================================
//  MIGRAÇÃO SEGURA v5 -> v6  (NÃO apaga dados existentes)
//  Rode esta função UMA VEZ na sua planilha atual de produção.
// ============================================================
function migrarV6() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // 1) cria as abas novas, se não existirem (sem mexer nas que já existem)
  const novasAbas = {
    [CONFIG.SHEETS.REGRAS_CONVENIO]: ['convenio_id','convenio_nome','tipo_faturamento','dia_fixo','dia_corte','prazo_envio_dias_uteis','prazo_recebimento_dias','atraso_frequente','ativo','observacao'],
    [CONFIG.SHEETS.REGRAS_COMISSAO]: ['profissional_id','profissional_nome','tipo_vinculo','tipo_remuneracao','valor_fixo_mensal','fixo_recebe_comissao_guia','variacao_particular_convenio','percentual_particular','variacao_por_servico','categoria_a_nome','categoria_a_codigos','percentual_a','unidade_a','categoria_b_nome','categoria_b_codigos','percentual_b','unidade_b','ativo'],
    [CONFIG.SHEETS.GLOSAS]: ['id','guia_id','convenio_nome','data_glosa','valor_glosado','motivo','status','observacao','lancado_por','criado_em'],
    [CONFIG.SHEETS.LOTES]: ['id','lote','convenio_id','convenio_nome','data_fechamento','data_envio','data_prev_recebimento','valor_total_lote','valor_recebido','valor_glosado','status','guias_ids','observacao','criado_em','atualizado_em'],
    [CONFIG.SHEETS.DESPESAS_RECORRENTES]: ['id','descricao','fornecedor','categoria','tipo','valor_padrao','regra_vencimento','forma_pgto','data_inicio','data_fim','ativo','observacao','criado_por','criado_em']
  };
  Object.keys(novasAbas).forEach(nome => {
    let sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      const cab = novasAbas[nome];
      sh.getRange(1,1,1,cab.length).setValues([cab]).setFontWeight('bold').setBackground('#0049AF').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
      log.push('Aba criada: ' + nome);
    } else {
      log.push('Aba já existia (mantida): ' + nome);
    }
  });

  // 2) adiciona colunas novas em abas existentes, SEM apagar as atuais
  log.push(..._garantirColunas(ss, CONFIG.SHEETS.ITENS_GUIA, ['profissional_id','profissional_nome','concluido']));
  log.push(..._garantirColunas(ss, CONFIG.SHEETS.GUIAS, ['lote_id']));
  // se você já rodou uma versão anterior desta migração (que deu erro), a aba
  // Regras_Comissao pode existir com cabeçalho antigo — garante as colunas novas.
  log.push(..._garantirColunas(ss, CONFIG.SHEETS.REGRAS_COMISSAO, ['unidade_a','unidade_b']));
  log.push(..._garantirColunas(ss, CONFIG.SHEETS.DESPESAS, ['origem_recorrente_id']));

  // 3) popula Regras_Convenio a partir da aba Convênios + regras do PDF,
  //    SOMENTE para convênios que ainda não têm regra cadastrada.
  _popularRegrasConvenioPadrao(ss, log);

  // 4) popula Regras_Comissao a partir da aba Profissionais,
  //    SOMENTE para profissionais que ainda não têm regra cadastrada.
  _popularRegrasComissaoPadrao(ss, log);

  // 5) popula Despesas_Recorrentes com a lista de despesas fixas/recorrentes
  //    que você passou, SOMENTE se a aba ainda estiver vazia (não roda de novo
  //    se você já editou/completou manualmente).
  _popularDespesasRecorrentesPadrao(ss, log);

  Logger.log(log.join('\n'));
  return { ok: true, log };
}

// ============================================================
//  MIGRAÇÃO SEGURA v6 -> v7  (NÃO apaga dados existentes)
//  Rode esta função UMA VEZ na sua planilha atual de produção,
//  depois de já ter rodado migrarV6() em algum momento anterior.
//  Adiciona: aba Receitas_Recorrentes (nova fonte de receita, ex:
//  aluguel de espaço/piscina) e atualiza a meta anual para 1.400.000.
// ============================================================
function migrarV7() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  const novasAbas = {
    [CONFIG.SHEETS.RECEITAS_RECORRENTES]: ['id','descricao','pagador_tipo','pagador_id','pagador_nome','categoria','valor_padrao','regra_vencimento','forma_pgto','data_inicio','data_fim','ativo','observacao','criado_por','criado_em'],
    [CONFIG.SHEETS.PLANO_CONTAS]: PLANO_CONTAS_HEADERS
  };
  Object.keys(novasAbas).forEach(nome => {
    let sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      const cab = novasAbas[nome];
      sh.getRange(1,1,1,cab.length).setValues([cab]).setFontWeight('bold').setBackground('#0049AF').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
      log.push('Aba criada: ' + nome);
    } else {
      log.push('Aba já existia (mantida): ' + nome);
    }
  });

  // popula com o cadastro inicial conhecido (aluguel de piscina da Yasmin),
  // SOMENTE se a aba ainda estiver vazia — não sobrescreve edições manuais.
  _popularReceitasRecorrentesPadrao(ss, log);

  Logger.log(log.join('\n'));
  return { ok: true, log, msg: 'Migração v7 concluída. Meta anual já está em ' + CONFIG.META_ANUAL + ' no código — não precisa alterar planilha para isso.' };
}

function _popularReceitasRecorrentesPadrao(ss, log) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.RECEITAS_RECORRENTES);
  if (!sh) return;
  if (sh.getLastRow() > 1) { log.push('Receitas_Recorrentes já tinha dados — não sobrescrito.'); return; }
  const shProf = ss.getSheetByName(CONFIG.SHEETS.PROFISSIONAIS);
  const profs = shProf && shProf.getLastRow() > 1 ? shProf.getRange(2,1,shProf.getLastRow()-1,2).getValues() : [];
  const yasmin = profs.find(r => String(r[1]).toUpperCase().indexOf('YASMIN') !== -1);
  const yasminId = yasmin ? yasmin[0] : '';
  const yasminNome = yasmin ? yasmin[1] : 'YASMIN SILVA';
  sh.appendRow(['REC001', 'Aluguel de Espaço — Piscina', 'profissional', yasminId, yasminNome, 'Aluguel de Espaço', 425, 'ATE_DIA_5', 'PIX', '', '', 'SIM', 'Cadastro inicial — confirmar periodicidade e forma de pagamento com a gestora.', 'SISTEMA', new Date()]);
  log.push('Receita recorrente inicial cadastrada: Aluguel de Espaço — Piscina (Yasmin, R$425).');
}

// Garante que uma aba tenha determinadas colunas ao final do cabeçalho,
// sem apagar linhas nem colunas existentes. Preenche células novas com ''.
function _garantirColunas(ss, nomeAba, colunasNovas) {
  const log = [];
  const sh = ss.getSheetByName(nomeAba);
  if (!sh) { log.push('AVISO: aba ' + nomeAba + ' não encontrada, colunas não adicionadas.'); return log; }
  const lastCol = sh.getLastColumn();
  const cabecalhoAtual = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
  colunasNovas.forEach(col => {
    if (cabecalhoAtual.indexOf(col) === -1) {
      const novaCol = sh.getLastColumn() + 1;
      sh.getRange(1, novaCol).setValue(col).setFontWeight('bold').setBackground('#0049AF').setFontColor('#FFFFFF');
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        // valor padrão seguro: 'concluido' assume SIM (não travar comissão de guias antigas),
        // as demais ficam em branco para preenchimento manual.
        const padrao = col === 'concluido' ? 'SIM' : '';
        sh.getRange(2, novaCol, lastRow-1, 1).setValue(padrao);
      }
      log.push('Coluna "' + col + '" adicionada em ' + nomeAba);
    } else {
      log.push('Coluna "' + col + '" já existia em ' + nomeAba);
    }
  });
  return log;
}

function _criarAba(ss, nome, cabecalhos) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  else sh.clearContents();
  const r = sh.getRange(1, 1, 1, cabecalhos.length);
  r.setValues([cabecalhos]).setFontWeight('bold').setBackground('#0049AF').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

// Pode ser chamada internamente pelo setupInicial(ss) ou manualmente pelo
// editor do Apps Script sem argumento. Nunca usa uma planilha indefinida.
function _popularDadosIniciais(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Nenhuma planilha ativa encontrada. Abra o projeto pelo Google Sheets correto.');
  const shU = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  shU.getRange(2,1,4,10).setValues([
    ['USR001','Admin Master','admin','admin','SIM',_hashSenha('admin'),new Date(),'','#0049AF','Administrador'],
    ['USR002','Gestora Financeira','gestor@institutodador.com.br','gestor','SIM',_hashSenha('gestor2026'),new Date(),'','#FAAF34','Gestora'],
    ['USR003','Recepção Principal','recepcao@institutodador.com.br','recepcao','SIM',_hashSenha('recepcao2026'),new Date(),'','#22c55e','Recepcionista'],
    ['USR004','Bruno Nascimento','fisio@institutodador.com.br','fisioterapeuta','SIM',_hashSenha('fisio2026'),new Date(),'','#8b5cf6','Fisioterapeuta'],
  ]);

  const shC = ss.getSheetByName(CONFIG.SHEETS.CONVENIOS);
  shC.getRange(2,1,12,7).setValues([
    ['CONV001','GEAP',90,'','SIM','',''], ['CONV002','AMIL',30,'','SIM','',''],
    ['CONV003','CASSÍ',30,'','SIM','',''], ['CONV004','CASSIND',60,'','SIM','',''],
    ['CONV005','CAPESESP',60,'','SIM','',''], ['CONV006','BLUE',60,'','SIM','',''],
    ['CONV007','PETROBRAS',30,'','SIM','',''], ['CONV008','ASSEC',60,'','SIM','',''],
    ['CONV009','CASSE',60,'','SIM','',''], ['CONV010','CASEC',60,'','SIM','',''],
    ['CONV011','CASSIND2',60,'','SIM','',''], ['CONV012','PARTICULAR',0,'','SIM','',''],
  ]);

  const shP = ss.getSheetByName(CONFIG.SHEETS.PROFISSIONAIS);
  shP.getRange(2,1,7,11).setValues([
    ['PROF001','BRUNO NASCIMENTO','Fisioterapia Motora / Liberação','PJ','','SIM',new Date(),'#0049AF','','',''],
    ['PROF002','LETICIA HELEN','Fisioterapia Motora','CLT','','SIM',new Date(),'#e63946','','',''],
    ['PROF003','LUCAS REZENDE','RPG / Fisioterapia Vestibular','PJ','','SIM',new Date(),'#FAAF34','','',''],
    ['PROF004','MARIANA MENDONÇA','TT Dor / Liberação / DTM','PJ','','SIM',new Date(),'#8b5cf6','','',''],
    ['PROF005','MILA PIRES','Fisioterapia Pélvica','PJ','','SIM',new Date(),'#2a9d8f','','',''],
    ['PROF006','MILENA MORAES','Acupuntura','PJ','','SIM',new Date(),'#22c55e','','',''],
    ['PROF007','OSMALÍ SILVA','Fisioterapia Motora','PJ','','SIM',new Date(),'#f4a261','','','']
  ]);

  // Catálogo completo — alinhado com a tabela real "VALORES_SERVIÇOS -
  // INSTITUTO DA DOR (PARA VALIDAÇÃO)". Colunas: id, nome, categoria,
  // valor_particular (sessão avulsa, referência), ativo, duracao_minutos.
  // Pilates não tem "sessão avulsa" — usa matriz própria em
  // getMatrizPrecoServico() (MODULO_CONFIG_FINANCEIRA.gs), o valor aqui é
  // só um placeholder de referência para telas que esperam um número único.
  const shSv = ss.getSheetByName(CONFIG.SHEETS.SERVICOS);
  shSv.getRange(2,1,18,6).setValues([
    ['SRV001','FISIOTERAPIA MOTORA','Fisioterapia',130,'SIM',50],
    ['SRV002','FISIOTERAPIA PÉLVICA','Fisioterapia',200,'SIM',50],
    ['SRV003','RPG','Fisioterapia',180,'SIM',60],
    ['SRV004','AVALIAÇÃO INICIAL','Avaliação',180,'SIM',60],
    ['SRV005','ACUPUNTURA','Terapia Complementar',180,'SIM',50],
    ['SRV006','HOME CARE - FISIOTERAPIA','Home Care',180,'SIM',60],
    ['SRV007','QUIROPRAXIA','Terapia Manual',220,'SIM',50],
    ['SRV008','DRENAGEM LINFÁTICA','Estética',180,'SIM',60],
    ['SRV009','FISIOTERAPIA AQUÁTICA','Fisioterapia',110,'SIM',50],
    ['SRV010','DTM','Fisioterapia',180,'SIM',50],
    ['SRV011','TRATAMENTO DA DOR','Terapia da Dor',220,'SIM',50],
    ['SRV012','CONSULTA ORTOPEDISTA','Consulta',400,'SIM',30],
    ['SRV013','INFILTRAÇÃO COM CORTICOIDE','Procedimento',300,'SIM',30],
    ['SRV014','INFILTRAÇÃO COM ÁCIDO','Procedimento',1000,'SIM',30],
    ['SRV015','LIBERAÇÃO','Terapia Manual',180,'SIM',40],
    ['SRV016','NUTRIÇÃO','Consulta',200,'SIM',40],
    ['SRV017','PILATES','Pilates',220,'SIM',50],
    ['SRV018','PILATES - CASA/HOME','Pilates',210,'SIM',50],
  ]);

  const shCd = ss.getSheetByName(CONFIG.SHEETS.CHECKLIST_DEF);
  shCd.getRange(2,1,24,9).setValues([
    ['CK_R001','recepcao','Abrir o sistema e verificar a agenda do dia','Confirme todos os horários marcados para hoje','Abertura','SIM',1,'SIM','calendar'],
    ['CK_R002','recepcao','Confirmar presença dos pacientes do dia','Ligar ou mandar mensagem para os pacientes agendados','Pacientes','SIM',2,'SIM','phone'],
    ['CK_R003','recepcao','Registrar chegada dos pacientes','Marcar check-in de quem chegou na agenda','Atendimentos','SIM',3,'SIM','user-check'],
    ['CK_R004','recepcao','Lançar atendimentos particulares realizados','Registrar todos os pagamentos recebidos no dia','Financeiro','SIM',4,'SIM','dollar-sign'],
    ['CK_R005','recepcao','Registrar novas guias de convênio','Lançar as guias dos atendimentos de convênio do dia','Financeiro','SIM',5,'SIM','file-text'],
    ['CK_R006','recepcao','Verificar pacientes que faltaram','Marcar falta e tentar reagendar','Pacientes','SIM',6,'SIM','user-x'],
    ['CK_R007','recepcao','Confirmar agenda do dia seguinte','Verificar os agendamentos de amanhã e confirmar','Agenda','SIM',7,'SIM','calendar-check'],
    ['CK_R008','recepcao','Guardar comprovantes de pagamento','Organizar todos os recibos do dia','Organização','NÃO',8,'SIM','folder'],
    ['CK_R009','recepcao','Registrar qualquer reclamação ou elogio','Anotar feedbacks dos pacientes','Qualidade','NÃO',9,'SIM','message-square'],
    ['CK_G001','gestor','Revisar lançamentos do dia anterior','Conferir se todos os atendimentos foram registrados corretamente','Financeiro','SIM',1,'SIM','check-circle'],
    ['CK_G002','gestor','Verificar guias a vencer em 7 dias','Identificar convênios com prazo próximo de vencimento','Convênios','SIM',2,'SIM','alert-triangle'],
    ['CK_G003','gestor','Conferir conformidade da recepção','Ver o relatório de checklist da equipe de recepção','Gestão','SIM',3,'SIM','bar-chart-2'],
    ['CK_G004','gestor','Lançar despesas do dia','Registrar todas as saídas financeiras do dia','Financeiro','SIM',4,'SIM','trending-down'],
    ['CK_G005','gestor','Aprovar ou rejeitar lançamentos pendentes','Revisar entradas com status pendente','Financeiro','NÃO',5,'SIM','thumbs-up'],
    ['CK_G006','gestor','Atualizar status de guias pagas','Confirmar recebimentos de convênios','Convênios','SIM',6,'SIM','credit-card'],
    ['CK_G007','gestor','Verificar projeção de caixa da semana','Analisar o fluxo previsto para os próximos 7 dias','Financeiro','NÃO',7,'SIM','trending-up'],
    ['CK_F001','fisioterapeuta','Verificar agenda do dia','Conferir seus atendimentos programados para hoje','Agenda','SIM',1,'SIM','calendar'],
    ['CK_F002','fisioterapeuta','Registrar evolução dos pacientes','Atualizar prontuário após cada atendimento','Clínico','SIM',2,'SIM','clipboard'],
    ['CK_F003','fisioterapeuta','Confirmar materiais necessários','Verificar se há insumos para os procedimentos do dia','Operacional','NÃO',3,'SIM','package'],
    ['CK_F004','fisioterapeuta','Solicitar reposição de materiais','Avisar recepção sobre materiais em falta','Operacional','NÃO',4,'SIM','shopping-cart'],
    ['CK_F005','fisioterapeuta','Avaliar altas e encaminhamentos','Revisar pacientes com possível alta ou necessidade de encaminhamento','Clínico','SIM',5,'SIM','arrow-right'],
    ['CK_F006','fisioterapeuta','Reportar intercorrências','Comunicar qualquer problema ou situação especial ao gestor','Qualidade','SIM',6,'SIM','alert-circle'],
    ['CK_F007','fisioterapeuta','Organizar área de trabalho ao final do dia','Deixar a sala limpa e organizada','Organização','NÃO',7,'SIM','home'],
    ['CK_F008','fisioterapeuta','Assinar guias dos convênios atendidos','Rubricar as guias dos atendimentos realizados','Administrativo','SIM',8,'SIM','pen-tool'],
  ]);

  _log('SISTEMA', 'SETUP_V6', 'Setup inicial v6.0 realizado');
}

// ============================================================
//  SEGURANÇA
// ============================================================
function _hashSenha(senha) {
  if (!senha) return '';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha + 'ID_SALT_2026', Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function _checarRateLimit(email) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGIN_ATTEMPTS);
  if (!sh) return { bloqueado: false };
  const data = sh.getDataRange().getValues();
  const agora = new Date().getTime();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      const bloqAte = data[i][3] ? new Date(data[i][3]).getTime() : 0;
      if (bloqAte > agora) {
        const mins = Math.ceil((bloqAte - agora) / 60000);
        return { bloqueado: true, msg: `Conta bloqueada. Tente novamente em ${mins} minuto(s).` };
      }
      if (data[i][1] >= CONFIG.MAX_LOGIN_ATTEMPTS) {
        const nd = new Date(agora + CONFIG.LOCKOUT_MINUTES * 60000);
        sh.getRange(i + 1, 4).setValue(nd);
        return { bloqueado: true, msg: `Bloqueado por ${CONFIG.LOCKOUT_MINUTES} minutos.` };
      }
      return { bloqueado: false, row: i + 1, tentativas: data[i][1] };
    }
  }
  return { bloqueado: false, row: null, tentativas: 0 };
}

function _registrarTentativa(email, sucesso) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGIN_ATTEMPTS);
  if (!sh) return;
  const agora = new Date();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      if (sucesso) sh.getRange(i+1,1,1,4).setValues([[email,0,agora,'']]);
      else sh.getRange(i+1,1,1,3).setValues([[email,(data[i][1]||0)+1,agora]]);
      return;
    }
  }
  if (!sucesso) sh.appendRow([email,1,agora,'']);
}

function _verificarUsuario(usuario, perfisPermitidos) {
  if (!usuario || !usuario.id) throw new Error('Usuário não autenticado.');
  if (perfisPermitidos && !perfisPermitidos.includes(usuario.perfil)) throw new Error('Acesso não autorizado.');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USUARIOS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === usuario.id && data[i][4] === 'SIM') return true;
  }
  throw new Error('Sessão inválida.');
}

// ============================================================
//  AUTENTICAÇÃO
// ============================================================
function login(email, senha) {
  try {
    if (!email || !senha) return { ok: false, msg: 'Preencha e-mail e senha.' };
    email = email.toLowerCase().trim();
    _garantirAcessoInicial();
    const rl = _checarRateLimit(email);
    if (rl.bloqueado) return { ok: false, msg: rl.msg };
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USUARIOS);
    const dados = sh.getDataRange().getValues();
    const hash = _hashSenha(senha);
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      const identificadorValido = String(row[2] || '').toLowerCase() === email || (email === 'admin' && row[3] === 'admin');
      if (identificadorValido && row[4] === 'SIM') {
        if (row[5] === hash || row[5] === senha) {
          _registrarTentativa(email, true);
          sh.getRange(i+1,8).setValue(new Date());
          if (row[5] === senha) sh.getRange(i+1,6).setValue(hash);
          const user = { id:row[0], nome:row[1], email:row[2], perfil:row[3], cor:row[8], cargo:row[9] };
          _log(user.nome, 'LOGIN_OK', `Perfil: ${user.perfil}`);
          return { ok: true, user };
        } else {
          _registrarTentativa(email, false);
          return { ok: false, msg: 'E-mail ou senha incorretos.' };
        }
      }
    }
    _registrarTentativa(email, false);
    return { ok: false, msg: 'Usuário não encontrado.' };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function registrarLogout(usuario) {
  if (usuario) _log(usuario.nome, 'LOGOUT', 'Sessão encerrada');
}

// ============================================================
//  HELPERS GERAIS (usados por todos os módulos)
//
//  CORREÇÃO CRÍTICA v7.1 — bug raiz do dashboard mostrando valores
//  muito abaixo do esperado (ex: "R$2.590,82" de receita total num
//  período que devia somar centenas de milhares de reais):
//
//  1) _sanNum() usava só `parseFloat(v)`, que lê número BR com vírgula
//     decimal (ex: "150,00") corretamente por acaso, mas TRUNCA
//     silenciosamente qualquer valor com separador de milhar em ponto
//     (ex: "1.234,56" virava 1.234 em vez de 1234.56 — um valor ~1000x
//     menor). Como várias colunas de valor vêm da planilha como texto
//     formatado em BR, isso subestimava boa parte dos lançamentos.
//  2) _mesDoDate() exigia receber sempre um objeto Date (`d.getMonth()`
//     direto) — mas é chamado em vários lugares passando a STRING crua
//     da planilha (ex: `_mesDoDate(r.data)` no dashboard), o que
//     lançava TypeError e derrubava silenciosamente aquele lançamento
//     do agrupamento por mês (ele não entrava em nenhum mês, nem em
//     "Todos os meses").
//  3) _dateStr() para string fazia só `.slice(0,10)`, assumindo que a
//     string já vinha em yyyy-MM-dd — quebra para qualquer outro
//     formato de data armazenado na planilha.
//
//  As versões abaixo tratam os dois formatos de número (BR e US) e
//  aceitam Date ou string em qualquer uma dessas 4 funções, sempre
//  tratando a data como horário LOCAL (nunca UTC — ver _parseDataLocal),
//  evitando o clássico bug de "dia 1º do mês vira dia 30/31 do mês
//  anterior" que a interpretação UTC de strings yyyy-MM-dd causa em
//  fusos negativos como o do Brasil.
// ============================================================
function _parseDataLocal(data) {
  if (data instanceof Date) return data;
  if (data === null || data === undefined || data === '') return new Date(NaN);
  const s = String(data);
  const mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mIso) return new Date(parseInt(mIso[1],10), parseInt(mIso[2],10)-1, parseInt(mIso[3],10));
  const mBr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mBr) return new Date(parseInt(mBr[3],10), parseInt(mBr[2],10)-1, parseInt(mBr[1],10));
  return new Date(s); // fallback: deixa o JS tentar (ex: já serializado pelo Sheets)
}
const _san = v => v ? String(v).replace(/[<>'"&]/g,'').trim() : '';
function _sanNum(valor, padrao) {
  const p = (padrao === undefined) ? 0 : padrao;
  if (valor === null || valor === undefined || valor === '') return p;
  if (typeof valor === 'number') return isNaN(valor) ? p : valor;
  let s = String(valor).trim();
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    // tem os dois separadores -> formato BR (1.234,56): ponto é milhar, vírgula é decimal
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') !== -1) {
    // só vírgula -> decimal BR simples (ex: "150,00")
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? p : n;
}
function _mesDoDate(data) {
  const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
  const d = _parseDataLocal(data);
  if (isNaN(d.getTime())) return '';
  return MESES[d.getMonth()];
}
function _dateStr(data) {
  const d = _parseDataLocal(data);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd');
}

function _getSheet(ss, nome) {
  const sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
}

function _toObj(headers, row) {
  const obj = {};
  headers.forEach((h,i) => obj[h] = row[i] !== undefined ? row[i] : '');
  return obj;
}

function _log(usuario, acao, detalhes) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG);
    if (sh) sh.appendRow([new Date(), usuario, acao, detalhes, '']);
  } catch(e) {}
}

function _getOrCreateFolder(nome) {
  const folders = DriveApp.getFoldersByName(nome);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(nome);
}

// ============================================================
//  LISTAS (para popular selects no front-end)
// ============================================================
function getListas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profissionais = _getSheet(ss, CONFIG.SHEETS.PROFISSIONAIS).filter(r=>r[5]==='SIM').map(r=>({id:r[0],nome:r[1],especialidade:r[2],cor:r[7]}));
  const pacientes = _getSheet(ss, CONFIG.SHEETS.PACIENTES).filter(r=>r[18]==='SIM').map(r=>({id:r[0],nome:r[1],convenio:r[15],cpf:r[2],telefone:r[6]}));
  const servicos = _getSheet(ss, CONFIG.SHEETS.SERVICOS).filter(r=>r[4]==='SIM').map(r=>({id:r[0],nome:r[1],categoria:r[2],valor:r[3],duracao:r[5]}));
  const convenios = _getSheet(ss, CONFIG.SHEETS.CONVENIOS).filter(r=>r[4]==='SIM').map(r=>({id:r[0],nome:r[1],prazo:r[2]}));
  const codigos = _getSheet(ss, CONFIG.SHEETS.CODIGOS).filter(r=>r[6]==='SIM').map(r=>({id:r[0],convenio_id:r[1],convenio_nome:r[2],codigo:r[3],descricao:r[4],valor:r[5]}));
  return { profissionais, pacientes, servicos, convenios, codigos };
}

// ============================================================
//  CHECKLIST — Por perfil, histórico por usuário (mantido)
// ============================================================
function getChecklistDefinicoes(usuario) {
  try {
    const perfil = usuario ? usuario.perfil : 'recepcao';
    const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.CHECKLIST_DEF);
    const headers = ['id','perfil','titulo','descricao','categoria','obrigatorio','ordem','ativo','icone'];
    return rows.map(r => _toObj(headers, r))
      .filter(r => r.ativo === 'SIM' && (r.perfil === perfil || r.perfil === 'todos'))
      .sort((a,b) => parseInt(a.ordem) - parseInt(b.ordem));
  } catch(e) { return []; }
}

function getChecklistHoje(usuario) {
  try {
    const hoje = _dateStr(new Date());
    const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.CHECKLIST_HIST);
    const headers = ['id','data','usuario_id','usuario_nome','perfil','checklist_id','titulo','concluido','hora_conclusao','observacao','criado_em'];
    return rows.map(r => _toObj(headers,r))
      .filter(r => _dateStr(r.data) === hoje && r.usuario_id === usuario.id);
  } catch(e) { return []; }
}

function salvarChecklistDia(items, usuario) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.SHEETS.CHECKLIST_HIST);
    const hoje = _dateStr(new Date());
    const agora = new Date();

    const allData = sh.getDataRange().getValues();
    const linhasParaRemover = [];
    for (let i = allData.length - 1; i >= 1; i--) {
      const d = _dateStr(allData[i][1]);
      if (d === hoje && allData[i][2] === usuario.id) linhasParaRemover.push(i + 1);
    }
    linhasParaRemover.forEach(l => sh.deleteRow(l));

    items.forEach(item => {
      const id = 'CKH' + new Date().getTime() + '_' + item.checklist_id;
      sh.appendRow([id, hoje, usuario.id, usuario.nome, usuario.perfil, item.checklist_id, item.titulo, item.concluido ? 'SIM' : 'NÃO', item.concluido ? Utilities.formatDate(agora,'America/Sao_Paulo','HH:mm') : '', item.observacao || '', agora]);
    });

    const total = items.length;
    const obrigatorios = items.filter(i => i.obrigatorio === 'SIM' || i.obrigatorio === true);
    const concluidos = obrigatorios.filter(i => i.concluido);
    const pct = obrigatorios.length > 0 ? Math.round((concluidos.length / obrigatorios.length) * 100) : 100;

    if (usuario.perfil === 'recepcao' && pct < 80) {
      _criarNotificacao('conformidade',
        `⚠️ Checklist incompleto — ${usuario.nome}`,
        `${usuario.nome} enviou o checklist com ${pct}% de conformidade (${concluidos.length}/${obrigatorios.length} obrigatórios).`,
        'gestor', usuario.nome);
    }

    _log(usuario.nome, 'CHECKLIST_ENVIADO', `${concluidos.length}/${total} itens | ${pct}% conformidade`);
    return { ok: true, pct, concluidos: concluidos.length, total };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function getRelatorioConformidade(dias) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hist = _getSheet(ss, CONFIG.SHEETS.CHECKLIST_HIST);
    const defs = _getSheet(ss, CONFIG.SHEETS.CHECKLIST_DEF);
    const headers = ['id','data','usuario_id','usuario_nome','perfil','checklist_id','titulo','concluido','hora_conclusao','observacao','criado_em'];
    const headDefs = ['id','perfil','titulo','descricao','categoria','obrigatorio','ordem','ativo','icone'];

    const n = dias || 7;
    const hoje = new Date();
    const limite = new Date(hoje - n * 86400000);

    const registros = hist.map(r => _toObj(headers,r))
      .filter(r => new Date(r.data) >= limite && r.perfil === 'recepcao');

    const defObrig = defs.map(r => _toObj(headDefs,r))
      .filter(r => r.perfil === 'recepcao' && r.obrigatorio === 'SIM' && r.ativo === 'SIM');

    const mapa = {};
    registros.forEach(r => {
      const chave = _dateStr(r.data) + '|' + r.usuario_id;
      if (!mapa[chave]) mapa[chave] = { data: _dateStr(r.data), usuario: r.usuario_nome, itens: [], enviou: true };
      mapa[chave].itens.push(r);
    });

    const resultado = Object.values(mapa).map(entry => {
      const obrigConcluidos = entry.itens.filter(i => {
        const def = defObrig.find(d => d.id === i.checklist_id);
        return def && i.concluido === 'SIM';
      });
      return {
        ...entry,
        total_obrigatorios: defObrig.length,
        concluidos_obrigatorios: obrigConcluidos.length,
        pct: Math.round((obrigConcluidos.length / defObrig.length) * 100)
      };
    }).sort((a,b) => b.data.localeCompare(a.data));

    return { ok: true, dados: resultado, defObrig: defObrig.length };
  } catch(e) { return { ok: false, dados: [], msg: e.toString() }; }
}

function salvarChecklistDefinicao(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin', 'gestor']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CHECKLIST_DEF);
    if (dados.id) {
      const allData = sh.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === dados.id) {
          sh.getRange(i+1,1,1,9).setValues([[dados.id, dados.perfil, _san(dados.titulo), _san(dados.descricao), _san(dados.categoria), dados.obrigatorio||'SIM', parseInt(dados.ordem)||1, dados.ativo||'SIM', dados.icone||'check-circle']]);
          _log(usuario.nome, 'EDIT_CHECKLIST_DEF', dados.titulo);
          return { ok: true };
        }
      }
    }
    const id = 'CK_' + dados.perfil.substring(0,1).toUpperCase() + String(sh.getLastRow()).padStart(3,'0');
    sh.appendRow([id, dados.perfil, _san(dados.titulo), _san(dados.descricao), _san(dados.categoria), dados.obrigatorio||'SIM', parseInt(dados.ordem)||99, 'SIM', dados.icone||'check-circle']);
    _log(usuario.nome, 'NOVO_CHECKLIST_DEF', dados.titulo);
    return { ok: true, id };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ============================================================
//  USUÁRIOS (mantido)
// ============================================================
function getUsuarios(usuario) {
  try {
    _verificarUsuario(usuario, ['admin']);
    return _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.USUARIOS)
      .map(r => ({id:r[0],nome:r[1],email:r[2],perfil:r[3],ativo:r[4],cargo:r[9],cor:r[8],criado_em:r[6],ultimo_login:r[7]}));
  } catch(e) { return []; }
}

function criarUsuario(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USUARIOS);
    const id = 'USR' + String(sh.getLastRow()).padStart(3,'0');
    const cores = {admin:'#0049AF',gestor:'#FAAF34',recepcao:'#22c55e',fisioterapeuta:'#8b5cf6'};
    sh.appendRow([id,_san(dados.nome),_san(dados.email).toLowerCase(),dados.perfil,'SIM',_hashSenha(dados.senha),new Date(),'',cores[dados.perfil]||'#666',_san(dados.cargo)||'']);
    _log(usuario.nome,'NOVO_USUARIO',dados.nome);
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function excluirUsuario(id, usuario) {
  try {
    _verificarUsuario(usuario, ['admin']);
    if (id === usuario.id) throw new Error('Não é possível excluir sua própria conta.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USUARIOS);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) { sh.getRange(i+1,5).setValue('NÃO'); _log(usuario.nome,'DESATIVAR_USUARIO',id); return {ok:true}; }
    }
    return {ok:false,msg:'Não encontrado'};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  PACIENTES (mantido)
// ============================================================
function salvarPaciente(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PACIENTES);
    const agora = new Date();
    if (dados.id) {
      const allData = sh.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === dados.id) {
          sh.getRange(i+1,1,1,21).setValues([[dados.id,_san(dados.nome),_san(dados.cpf),'','','',_san(dados.telefone),_san(dados.email),_san(dados.cep),_san(dados.logradouro),_san(dados.numero),_san(dados.complemento),_san(dados.bairro),_san(dados.cidade),'',_san(dados.convenio_principal),'',_san(dados.observacoes),'SIM',allData[i][19],agora]]);
          return {ok:true};
        }
      }
    }
    const id = 'PAC' + String(sh.getLastRow()).padStart(3,'0');
    sh.appendRow([id,_san(dados.nome),_san(dados.cpf),'','','',_san(dados.telefone),_san(dados.email),_san(dados.cep),_san(dados.logradouro),_san(dados.numero),_san(dados.complemento),_san(dados.bairro),_san(dados.cidade),'',_san(dados.convenio_principal||'PARTICULAR'),'',_san(dados.observacoes),'SIM',agora,agora]);
    _log(usuario.nome,'NOVO_PACIENTE',dados.nome);
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  PROFISSIONAIS (mantido — a regra financeira detalhada agora
//  vive em Regras_Comissao, ver Modulo_Comissionamento.gs)
// ============================================================
function salvarProfissional(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PROFISSIONAIS);
    const id = 'PROF' + String(sh.getLastRow()).padStart(3,'0');
    sh.appendRow([id,_san(dados.nome),_san(dados.especialidade),_san(dados.tipo_vinculo||'PJ'),_sanNum(dados.percentual),'SIM',new Date(),_san(dados.cor||'#0049AF'),_san(dados.cro),'','']);
    _log(usuario.nome,'NOVO_PROFISSIONAL',dados.nome);
    // ATENÇÃO (poka-yoke): todo profissional novo PRECISA de uma linha em
    // Regras_Comissao para que a comissão dele seja calculada corretamente.
    // Criamos aqui um rascunho com 0% para forçar o preenchimento manual
    // em vez de deixar o profissional sem nenhuma regra (o que geraria
    // comissão silenciosamente igual a zero sem avisar ninguém).
    _garantirRascunhoRegraComissao(id, _san(dados.nome), _san(dados.tipo_vinculo||'PJ'));
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  AGENDA (mantido)
// ============================================================
function getAgendamentos(filtros) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.AGENDA);
  const headers = ['id','data','hora','hora_fim','profissional_id','profissional_nome','paciente_id','paciente_nome','servico_id','servico_nome','tipo','status','observacao','criado_por','criado_em','atualizado_em','cor_profissional','duracao_minutos'];
  let result = rows.map(r => _toObj(headers,r));
  if (filtros && filtros.data) result = result.filter(r => _dateStr(r.data) === filtros.data);
  return result;
}

function salvarAgendamento(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.SHEETS.AGENDA);
    const agora = new Date();
    const id = 'AGD' + agora.getTime();
    const profs = _getSheet(ss, CONFIG.SHEETS.PROFISSIONAIS);
    const profRow = profs.find(r => r[0] === dados.profissional_id);
    const cor = profRow ? profRow[7] : '#0049AF';
    sh.appendRow([id,dados.data,dados.hora,'',dados.profissional_id,dados.profissional_nome,dados.paciente_id,dados.paciente_nome,dados.servico_id||'',dados.servico_nome||'','Sessão','confirmado',dados.observacao||'',usuario.nome,agora,agora,cor,dados.duracao_minutos||60]);
    _log(usuario.nome,'NOVO_AGENDAMENTO',`${dados.paciente_nome} | ${dados.data} ${dados.hora}`);
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function atualizarStatusAgendamento(id, status, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.AGENDA);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) {
        sh.getRange(i+1,12).setValue(status);
        sh.getRange(i+1,16).setValue(new Date());
        _log(usuario.nome,'STATUS_AGENDA',`${id} → ${status}`);
        return {ok:true};
      }
    }
    return {ok:false,msg:'Não encontrado'};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  PARTICULARES (mantido)
// ============================================================
function salvarParticular(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PARTICULARES);
    const id = 'PART' + new Date().getTime();
    const mes = _mesDoDate(new Date(dados.data));
    sh.appendRow([id,dados.data,mes,dados.paciente_id,dados.paciente_nome,dados.profissional_id,dados.profissional_nome,dados.servico_id,dados.servico_nome,_sanNum(dados.valor),dados.forma_pgto,parseInt(dados.quantidade)||1,dados.tipo_qtd,dados.observacao||'','confirmado',usuario.nome,new Date(),dados.agenda_id||'']);
    _criarNotificacao('lancamento','Novo lançamento particular',`${usuario.nome} — ${dados.paciente_nome} | R$ ${_sanNum(dados.valor).toFixed(2)}`,'gestor',usuario.nome);
    _log(usuario.nome,'NOVO_PARTICULAR',`R$ ${dados.valor}`);
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function getParticulares(filtros) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.PARTICULARES);
  const headers = ['id','data','mes','paciente_id','paciente_nome','profissional_id','profissional_nome','servico_id','servico_nome','valor','forma_pgto','quantidade','tipo_qtd','observacao','status','lancado_por','criado_em','agenda_id'];
  let result = rows.map(r => _toObj(headers,r));
  if (filtros && filtros.mes && filtros.mes !== 'todos') result = result.filter(r => r.mes === filtros.mes);
  return result;
}

// ============================================================
//  DESPESAS (mantido)
// ============================================================
function salvarDespesa(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.DESPESAS);
    const id = 'DESP' + new Date().getTime();
    const mes = _mesDoDate(new Date(dados.data||dados.data_vencimento||new Date()));
    sh.appendRow([id,dados.data||'',mes,dados.categoria,dados.descricao,dados.fornecedor||'',_sanNum(dados.valor),dados.forma_pgto||'',dados.tipo,dados.status,dados.data_vencimento||'',dados.data_pgto||'','',dados.observacao||'',usuario.nome,new Date(),'']);
    _log(usuario.nome,'NOVA_DESPESA',`R$ ${dados.valor}`);
    return {ok:true,id};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

function getDespesas(filtros) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.DESPESAS);
  const headers = ['id','data','mes','categoria','descricao','fornecedor','valor','forma_pgto','tipo','status','data_vencimento','data_pgto','comprovante_url','observacao','lancado_por','criado_em','origem_recorrente_id'];
  let result = rows.map(r => _toObj(headers,r));
  if (filtros && filtros.mes && filtros.mes !== 'todos') result = result.filter(r => r.mes === filtros.mes);
  return result;
}

function uploadComprovante(base64, filename, despesa_id, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const folder = _getOrCreateFolder('Instituto_Dor_Comprovantes');
    const decoded = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(decoded, filename.endsWith('.pdf') ? 'application/pdf' : 'image/jpeg', filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.DESPESAS);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === despesa_id) { sh.getRange(i+1,13).setValue(url); break; }
    }
    return {ok:true,url};
  } catch(e) { return {ok:false,msg:e.toString()}; }
}

// ============================================================
//  NOTIFICAÇÕES (mantido)
// ============================================================
function _criarNotificacao(tipo, titulo, mensagem, para_perfil, criado_por) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.NOTIFICACOES);
    sh.appendRow(['NOT'+new Date().getTime(),tipo,titulo,mensagem,para_perfil||'all','NAO',new Date(),criado_por||'sistema']);
  } catch(e) {}
}

function getNotificacoes(perfil) {
  const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.NOTIFICACOES);
  const headers = ['id','tipo','titulo','mensagem','para_perfil','lida','criado_em','lancado_por'];
  return rows.map(r => _toObj(headers,r))
    .filter(r => r.lida==='NAO' && (r.para_perfil==='all'||r.para_perfil===perfil))
    .slice(-30).reverse();
}

function marcarNotificacaoLida(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.NOTIFICACOES);
  const allData = sh.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0]===id) { sh.getRange(i+1,6).setValue('SIM'); return {ok:true}; }
  }
  return {ok:false};
}

// ============================================================
//  LOG DE AUDITORIA (mantido)
// ============================================================
function getLog(usuario) {
  try {
    _verificarUsuario(usuario, ['admin']);
    const rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.LOG);
    return rows.slice(-200).reverse().map(r => ({timestamp:r[0],usuario:r[1],acao:r[2],detalhes:r[3]}));
  } catch(e) { return []; }
}

// Garante que uma implantação nova consiga entrar antes da execução manual
// do setupInicial(). Não limpa nem altera planilhas existentes.
function _garantirAcessoInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.USUARIOS);
    sh.getRange(1,1,1,10).setValues([['id','nome','email','perfil','ativo','senha','criado_em','ultimo_login','cor_avatar','cargo']]);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() < 2) {
    sh.appendRow(['USR001','Admin Master','admin','admin','SIM',_hashSenha('admin'),new Date(),'','#0049AF','Administrador']);
  }
  let rate = ss.getSheetByName(CONFIG.SHEETS.LOGIN_ATTEMPTS);
  if (!rate) {
    rate = ss.insertSheet(CONFIG.SHEETS.LOGIN_ATTEMPTS);
    rate.getRange(1,1,1,4).setValues([['email','tentativas','ultimo_at','bloqueado_ate']]);
    rate.setFrozenRows(1);
  }
}

// ============================================================
//  PLANO DE CONTAS — CRUD financeiro independente dos módulos clínicos
// ============================================================
const PLANO_CONTAS_HEADERS = ['id','codigo','nome','tipo','parent_id','ativo','criado_em','atualizado_em'];
function _shPlanoContas(){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PLANO_CONTAS); }
function getPlanoContas(usuario){
  _verificarUsuario(usuario, ['admin','gestor']);
  const sh=_shPlanoContas(); if(!sh) return [];
  return _getSheet(SpreadsheetApp.getActiveSpreadsheet(),CONFIG.SHEETS.PLANO_CONTAS).map(r=>_toObj(PLANO_CONTAS_HEADERS,r));
}
function salvarContaPlano(dados, usuario){
  _verificarUsuario(usuario,['admin','gestor']);
  const sh=_shPlanoContas(); if(!sh) throw new Error('Execute setupInicial ou migrarV7.');
  const agora=new Date(), id=dados.id||Utilities.getUuid(), tipo=dados.tipo==='RECEITA'?'RECEITA':'DESPESA';
  const row=[id,_san(dados.codigo),_san(dados.nome),tipo,_san(dados.parent_id||''),dados.ativo==='NAO'?'NAO':'SIM',dados.criado_em||agora,agora];
  const vals=sh.getDataRange().getValues(); const i=vals.findIndex((r,n)=>n>0&&String(r[0])===id);
  if(i>0) sh.getRange(i+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  _log(usuario,'SALVAR_CONTA_PLANO',id); return {ok:true,id:id};
}
function excluirContaPlano(id, usuario){
  _verificarUsuario(usuario,['admin','gestor']); const sh=_shPlanoContas();
  const vals=sh.getDataRange().getValues(); const i=vals.findIndex((r,n)=>n>0&&String(r[0])===String(id));
  if(i<1) return {ok:false,msg:'Conta não encontrada'};
  sh.getRange(i+1,6).setValue('NAO'); _log(usuario,'INATIVAR_CONTA_PLANO',id); return {ok:true};
}

// ============================================================
//  CARGA INICIAL MASSIVA — v1
//  Executar uma única vez depois de setupInicial() ou migrarV7().
//  A carga é idempotente: identifica o que já existe por ID e nunca
//  apaga lançamentos ou configurações manuais.
// ============================================================
const PLANO_CONTAS_CARGA_INICIAL = [
  ['PC001','1','Benefícios','DESPESA',''],['PC002','1.1','Plano de Saúde','DESPESA','PC001'],
  ['PC003','1.2','Plano de Saúde Colaborador','DESPESA','PC001'],['PC004','1.3','Previdência Privada','DESPESA','PC001'],
  ['PC005','2','Cartão de Crédito','DESPESA',''],['PC006','2.1','Clínica','DESPESA','PC005'],['PC007','2.2','Clínica Caixa','DESPESA','PC005'],
  ['PC008','2.3','Visa','DESPESA','PC005'],['PC009','2.4','Elo','DESPESA','PC005'],['PC010','2.5','Cartão','DESPESA','PC005'],['PC011','2.6','Pilates','DESPESA','PC005'],['PC012','2.7','Clínica 2','DESPESA','PC005'],
  ['PC013','3','Empréstimos e Investimentos','DESPESA',''],['PC014','3.1','Parcela/Aporte','DESPESA','PC013'],['PC015','3.2','Seguro','DESPESA','PC013'],
  ['PC016','4','Folha de Pagamento','DESPESA',''],['PC017','4.1','Salário/Encargos','DESPESA','PC016'],
  ['PC018','5','Impostos e Taxas','DESPESA',''],['PC019','5.1','Bombeiros','DESPESA','PC018'],['PC020','5.2','Taxas Diversas / Licenças','DESPESA','PC018'],['PC021','5.3','Sindicato / SASE','DESPESA','PC018'],['PC022','5.4','Simples Nacional','DESPESA','PC018'],['PC023','5.5','FGTS','DESPESA','PC018'],['PC024','5.6','DARF','DESPESA','PC018'],['PC025','5.7','Taxas Bancárias','DESPESA','PC018'],['PC026','5.8','IPTU','DESPESA','PC018'],['PC027','5.9','INSS','DESPESA','PC018'],
  ['PC028','6','Infraestrutura','DESPESA',''],['PC029','6.1','Aluguel','DESPESA','PC028'],['PC030','6.2','Condomínio','DESPESA','PC028'],['PC031','6.3','Energia Elétrica','DESPESA','PC028'],['PC032','6.4','Água e Saneamento','DESPESA','PC028'],['PC033','6.5','Telefonia e Internet','DESPESA','PC028'],
  ['PC034','7','Insumos','DESPESA',''],['PC035','7.1','Copa','DESPESA','PC034'],['PC036','7.2','Papelaria','DESPESA','PC034'],['PC037','7.3','Fardamento','DESPESA','PC034'],
  ['PC038','8','Manutenção e Obras','DESPESA',''],['PC039','8.1','Reforma/Obras','DESPESA','PC038'],['PC040','8.2','Manutenção Geral','DESPESA','PC038'],
  ['PC041','9','Marketing','DESPESA',''],['PC042','9.1','Publicidade/Mídia','DESPESA','PC041'],
  ['PC043','10','Repasse a Profissionais','DESPESA',''],['PC044','10.1','Geral / Não Especificado','DESPESA','PC043'],['PC045','10.2','Convênio (Diversos)','DESPESA','PC043'],['PC046','10.3','Fonoaudiologia','DESPESA','PC043'],['PC047','10.4','Hidroterapia','DESPESA','PC043'],['PC048','10.5','Pró-Labore','DESPESA','PC043'],['PC049','10.6','Particular','DESPESA','PC043'],['PC050','10.7','Pilates','DESPESA','PC043'],['PC051','10.8','Atendimento Domiciliar','DESPESA','PC043'],['PC052','10.9','Convênio (CASSI)','DESPESA','PC043'],['PC053','10.10','Convênio (GEAP)','DESPESA','PC043'],['PC054','10.11','Convênio (CASSI + GEAP)','DESPESA','PC043'],
  ['PC055','11','Serviços Administrativos','DESPESA',''],['PC056','11.1','Contabilidade','DESPESA','PC055'],['PC057','11.2','Consultoria','DESPESA','PC055'],
  ['PC058','12','Sistemas e TI','DESPESA',''],['PC059','12.1','Software / Mensalidade','DESPESA','PC058'],
  ['PC060','13','Veículos','DESPESA',''],['PC061','13.1','Manutenção / Combustível / Parcela','DESPESA','PC060'],
  ['PC100','100','Receitas de Atendimento','RECEITA',''],['PC101','100.1','Particular','RECEITA','PC100'],['PC102','100.2','Convênios','RECEITA','PC100'],['PC103','100.3','Receita Recorrente','RECEITA','']
];

const PROFISSIONAIS_CARGA_INICIAL = {
  PROF008:['YANNA MENEZES','Fisioterapia Pélvica'],PROF009:['YASMIN SILVA','Fisioterapia Aquática'],PROF010:['LUIZA GABRIELA','Pilates / RPG / Quiropraxia / DTM'],
  PROF011:['JAMILLE GONÇALVES','Psicologia'],PROF012:['JULIANA LINHARES','Psicologia'],PROF013:['JOSI NASCIMENTO','Fisioterapia Motora'],PROF014:['JULIANA ARCHIMINO','RPG'],
  PROF015:['LAIS MARINHO','Fisioterapia Pélvica'],PROF016:['MARIA APARECIDA','Reabilitação'],PROF017:['GRACE KELLY','Fisioterapia Aquática'],PROF018:['DANIELA','Fisioterapia Aquática'],
  PROF019:['VIVIANE COSTA','Fisioterapia Motora / Home Care'],PROF020:['LUCIANA','Fisioterapia Motora / Home Care'],PROF021:['BRENA MIRELI','Nutrição'],PROF022:['VINICIUS SOBRAL','Ortopedia']
};

function _popularPlanoContasCargaInicial(ss, log){
  let sh=ss.getSheetByName(CONFIG.SHEETS.PLANO_CONTAS);
  if(!sh){sh=ss.insertSheet(CONFIG.SHEETS.PLANO_CONTAS);sh.getRange(1,1,1,PLANO_CONTAS_HEADERS.length).setValues([PLANO_CONTAS_HEADERS]);sh.setFrozenRows(1);}
  const existentes=new Set(_getSheet(ss,CONFIG.SHEETS.PLANO_CONTAS).map(r=>String(r[0]))), agora=new Date();
  PLANO_CONTAS_CARGA_INICIAL.forEach(([id,codigo,nome,tipo,parent])=>{if(!existentes.has(id))sh.appendRow([id,codigo,nome,tipo,parent,'SIM',agora,agora]);});
  log.push('Plano de Contas: '+PLANO_CONTAS_CARGA_INICIAL.length+' contas verificadas/cadastradas.');
}

function _popularProfissionaisCargaInicial(ss, log){
  const sh=ss.getSheetByName(CONFIG.SHEETS.PROFISSIONAIS); if(!sh)return;
  const existentes=new Set(_getSheet(ss,CONFIG.SHEETS.PROFISSIONAIS).map(r=>String(r[0])));
  Object.keys(PROFISSIONAIS_CARGA_INICIAL).forEach(id=>{if(!existentes.has(id)){const p=PROFISSIONAIS_CARGA_INICIAL[id];sh.appendRow([id,p[0],p[1],'PJ','', 'SIM',new Date(),'#0049AF','','','']);}});
  log.push('Profissionais adicionais: '+Object.keys(PROFISSIONAIS_CARGA_INICIAL).length+' verificadas/cadastradas.');
}

function popularDadosMassivo(usuario){
  try{
    _garantirAcessoInicial();
    const operador=usuario||{id:'USR001',nome:'Admin Master',perfil:'admin'};
    _verificarUsuario(operador,['admin','gestor']);
    const ss=SpreadsheetApp.getActiveSpreadsheet(), cfg=ss.getSheetByName(CONFIG.SHEETS.CONFIG), log=[];
    const marcadores=cfg&&cfg.getLastRow()>1?cfg.getRange(2,1,cfg.getLastRow()-1,2).getValues().map(r=>String(r[0])):[];
    if(marcadores.includes('CARGA_MASSIVA_V1_CONCLUIDA')) return {ok:true,jaExecutada:true,msg:'Carga massiva já executada; nada foi alterado.'};
    _popularProfissionaisCargaInicial(ss,log);
    _popularPlanoContasCargaInicial(ss,log);
    _popularRegrasComissaoPadrao(ss,log);
    _popularRegrasConvenioPadrao(ss,log);
    const historico=importarHistoricoReal(operador);
    if(!historico.ok) throw new Error(historico.msg||'Falha na importação do histórico.');
    // A importação pode cadastrar convênios novos; garante também a linha
    // correspondente em Regras_Convenio para revisão da gestora.
    _popularRegrasConvenioPadrao(ss,log);
    if(cfg) cfg.appendRow(['CARGA_MASSIVA_V1_CONCLUIDA',new Date()]);
    _log(operador.nome,'CARGA_MASSIVA_V1','Plano de contas, profissionais, regras, convênios, entradas e saídas importados.');
    return {ok:true,log,historico};
  }catch(e){return {ok:false,msg:e.toString()};}
}

// ===== MODULO_COMISSIONAMENTO.gs =====
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

// ===== MODULO_CONFIG_FINANCEIRA.gs =====
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

// ===== MODULO_CONVENIOS_FATURAMENTO.gs =====
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

// ===== MODULO_CURVA_ABC.gs =====
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

// ===== MODULO_DASHBOARD_KPI.gs =====
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

// Versão leve para a primeira pintura do Web App. Consultas detalhadas
// (glosas, lotes, conformidade e projeções) não podem bloquear o dashboard.
function getDashboardDataRapido(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao','fisioterapeuta']);
    const mes=(filtros&&filtros.mes)||'todos';
    const particulares=getParticulares({mes});
    const guias=getGuias({mes});
    const despesas=getDespesas({mes});
    const recParticular=particulares.reduce((s,r)=>s+_sanNum(r.valor),0);
    const recConvenio=guias.reduce((s,r)=>s+_sanNum(r.valor_total),0);
    const despesasTotal=despesas.reduce((s,r)=>s+_sanNum(r.valor),0);
    const porMes={}; MESES_ORDEM.forEach(m=>porMes[m]={particular:0,convenio:0,despesa:0});
    particulares.forEach(r=>{if(porMes[r.mes])porMes[r.mes].particular+=_sanNum(r.valor);});
    guias.forEach(r=>{if(porMes[r.mes])porMes[r.mes].convenio+=_sanNum(r.valor_total);});
    despesas.forEach(r=>{if(porMes[r.mes])porMes[r.mes].despesa+=_sanNum(r.valor);});
    const totalFaturado=recParticular+recConvenio;
    return {ok:true,kpis:{totalFaturado,totalRecebido:recParticular,totalGlosado:guias.reduce((s,r)=>s+_sanNum(r.valor_glosado),0),totalAReceber:guias.filter(r=>r.status!=='Pago').reduce((s,r)=>s+_sanNum(r.valor_total)-_sanNum(r.valor_glosado),0),recParticular,recConvenio,totalDespesas:despesasTotal,resultado:recParticular-despesasTotal,qtdGuiasPendentes:guias.filter(r=>r.status==='Pendente').length,ticketMedio:particulares.length?recParticular/particulares.length:0,metaAnual:CONFIG.META_ANUAL,pctMeta:(totalFaturado/CONFIG.META_ANUAL*100).toFixed(1)},porMes,porConvenio:[],porProfissional:{},aVencer:[],agendaHoje:[],conformidade:{ok:true,dados:[]},statusLotes:{},resumoGlosas:{ok:true,totalGlosado:0}};
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

// ===== MODULO_DESPESAS_RECORRENTES.gs =====
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

// ============================================================
//  FLUXO DE CAIXA HISTÓRICO — v7.1 (nova)
//
//  Diferente de getProjecaoFluxoCaixa() (que só olha do mês atual
//  em diante), esta função monta o fluxo de caixa REALIZADO dos
//  meses passados: quanto entrou (convênio + particular + receitas
//  recorrentes), quanto saiu (despesas), e o saldo de cada mês —
//  usando dados já lançados no sistema, não estimativa.
//
//  Retorna também, para os meses que já têm despesa recorrente
//  cadastrada com valor_padrao (tipo FIXO), uma comparação
//  "planejado x realizado": planejado = soma dos valor_padrao das
//  recorrentes ativas naquele mês; realizado = soma do que
//  efetivamente foi lançado em Despesas para aquele mês. Isso ajuda
//  a gestora ver se gastou mais ou menos do que o esperado.
// ============================================================
function getFluxoCaixaHistorico(qtdMesesPassados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
    const hoje = new Date();
    const n = qtdMesesPassados || 12;

    const recorrentes = getDespesasRecorrentes(usuario).filter(r => r.ativo === 'SIM');

    const meses = [];
    for (let i = n; i >= 0; i--) {
      const d = new Date(hoje.getFullYear(), hoje.getMonth()-i, 1);
      const mesTx = MESES[d.getMonth()];
      const ano = d.getFullYear();

      const particularesMes = getParticulares({mes:mesTx}).filter(p => new Date(p.data).getFullYear() === ano);
      const guiasMes = getGuias({mes:mesTx}).filter(g => new Date(g.data).getFullYear() === ano);
      const despesasMes = getDespesas({mes:mesTx}).filter(x => new Date(x.data||x.data_vencimento).getFullYear() === ano);
      const receitasRecMes = getResumoReceitasRecorrentes({mes:mesTx}, usuario);

      const entradaParticular = particularesMes.reduce((s,p)=>s+_sanNum(p.valor),0);
      const entradaConvenio = guiasMes.reduce((s,g)=>s+_sanNum(g.valor_total)-_sanNum(g.valor_glosado),0);
      const entradaRecorrente = receitasRecMes.ok ? receitasRecMes.total : 0;
      const totalEntradas = entradaParticular + entradaConvenio + entradaRecorrente;
      const totalSaidas = despesasMes.reduce((s,x)=>s+_sanNum(x.valor),0);
      const saldoMes = totalEntradas - totalSaidas;

      // planejado x realizado (só para despesas do tipo FIXO recorrente)
      const planejadoDespesa = recorrentes.filter(r => r.tipo==='FIXO' &&
        (!r.data_inicio || new Date(r.data_inicio) <= new Date(ano, d.getMonth()+1, 0)) &&
        (!r.data_fim || new Date(r.data_fim) >= new Date(ano, d.getMonth(), 1))
      ).reduce((s,r)=>s+_sanNum(r.valor_padrao),0);

      meses.push({
        mes: mesTx, ano, ehMesAtual: (i===0),
        entradaParticular, entradaConvenio, entradaRecorrente, totalEntradas,
        totalSaidas, saldoMes,
        planejadoDespesa, realizadoDespesa: totalSaidas,
        diferencaDespesa: totalSaidas - planejadoDespesa,
        qtdLancamentosParticular: particularesMes.length, qtdGuias: guiasMes.length
      });
    }

    const saldoAcumulado = meses.reduce((s,m)=>s+m.saldoMes,0);
    return { ok:true, meses, saldoAcumulado };
  } catch(e) { return { ok:false, meses:[], msg:e.toString() }; }
}

// ============================================================
//  PROJEÇÃO FUTURA COM BASE NO HISTÓRICO — v7.1
//
//  Junta o realizado (getFluxoCaixaHistorico) com a projeção futura
//  (getProjecaoFluxoCaixa), numa única série temporal contínua, para
//  o dashboard poder desenhar um gráfico único: "realizado até aqui"
//  seguido de "projeção para os próximos meses" — útil para visualizar
//  a tendência sem trocar de gráfico.
// ============================================================
function getFluxoCaixaCompleto(qtdMesesPassados, qtdMesesFuturos, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const historico = getFluxoCaixaHistorico(qtdMesesPassados||6, usuario);
    const projecao = getProjecaoFluxoCaixa(qtdMesesFuturos||3, usuario);
    if (!historico.ok) return { ok:false, msg:historico.msg };

    const serieHistorico = historico.meses.map(m => ({
      mes: m.mes, ano: m.ano, tipo: 'realizado',
      entradas: m.totalEntradas, saidas: m.totalSaidas, saldo: m.saldoMes
    }));
    const serieProjecao = (projecao.ok ? projecao.meses : []).filter(m => {
      // remove o mês atual da projeção se já vier no histórico, pra não duplicar no gráfico
      return !serieHistorico.some(h => h.mes === m.mes && h.ano === m.ano);
    }).map(m => ({
      mes: m.mes, ano: m.ano, tipo: 'projetado',
      entradas: m.receitaTotalEstimada, saidas: m.despesasPrevistas, saldo: m.lucroProjetado
    }));

    return { ok:true, serie: serieHistorico.concat(serieProjecao), aviso: projecao.aviso };
  } catch(e) { return { ok:false, serie:[], msg:e.toString() }; }
}

// ===== MODULO_GLOSAS.gs =====
// ============================================================
//  MODULO_GLOSAS.gs — v6.0
//  Cadastro manual de glosas + resumo para o dashboard.
//  status possíveis: Aberta, Recorrida, Mantida, Revertida
// ============================================================

const GLOSAS_HEADERS = ['id','guia_id','convenio_nome','data_glosa','valor_glosado','motivo','status','observacao','lancado_por','criado_em'];

function getGlosas(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    let rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.GLOSAS).map(r => _toObj(GLOSAS_HEADERS, r));
    if (filtros && filtros.status) rows = rows.filter(r => r.status === filtros.status);
    if (filtros && filtros.convenio_nome) rows = rows.filter(r => r.convenio_nome === filtros.convenio_nome);
    return rows;
  } catch(e) { return []; }
}

function salvarGlosa(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (!dados.guia_id) throw new Error('Informe a guia relacionada à glosa.');
    if (!_sanNum(dados.valor_glosado)) throw new Error('Informe o valor glosado (maior que zero).');
    if (!dados.motivo) throw new Error('Informe o motivo da glosa (poka-yoke: motivo em branco impede análise futura).');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
    const id = 'GLO' + new Date().getTime();
    sh.appendRow([id, dados.guia_id, _san(dados.convenio_nome), dados.data_glosa||_dateStr(new Date()), _sanNum(dados.valor_glosado), _san(dados.motivo), dados.status||'Aberta', _san(dados.observacao), usuario.nome, new Date()]);
    _log(usuario.nome, 'NOVA_GLOSA', `${dados.guia_id} | R$ ${dados.valor_glosado}`);
    return { ok:true, id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function atualizarStatusGlosa(id, status, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (!['Aberta','Recorrida','Mantida','Revertida'].includes(status)) throw new Error('status inválido.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
    const allData = sh.getDataRange().getValues();
    const colStatus = GLOSAS_HEADERS.indexOf('status') + 1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === id) { sh.getRange(i+1,colStatus).setValue(status); _log(usuario.nome,'STATUS_GLOSA',`${id} → ${status}`); return {ok:true}; }
    }
    return { ok:false, msg:'Glosa não encontrada' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Chamado automaticamente por atualizarStatusGuia() quando alguém marca
// valor_glosado > 0 direto na guia, sem passar pelo formulário de Glosas —
// evita ter valor glosado "solto" na guia sem nenhum registro rastreável.
function _garantirRegistroGlosa(guia_id, convenio_nome, valor, usuario) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.GLOSAS);
  if (!sh) return;
  const allData = sh.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) if (allData[i][1] === guia_id) return; // já existe
  sh.appendRow(['GLO'+new Date().getTime(), guia_id, convenio_nome, _dateStr(new Date()), valor, 'PREENCHER MOTIVO — gerado automaticamente ao marcar status da guia', 'Aberta', '', usuario ? usuario.nome : 'sistema', new Date()]);
}

function getResumoGlosas(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const glosas = getGlosas(filtros, usuario);
    const totalGlosado = glosas.reduce((s,g)=>s+_sanNum(g.valor_glosado),0);
    const porConvenio = {};
    glosas.forEach(g => { porConvenio[g.convenio_nome] = (porConvenio[g.convenio_nome]||0) + _sanNum(g.valor_glosado); });
    const porStatus = {};
    glosas.forEach(g => { porStatus[g.status] = (porStatus[g.status]||0) + 1; });
    return { ok:true, totalGlosado, porConvenio, porStatus, qtd: glosas.length };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ===== MODULO_GUIAS.gs =====
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

// ===== MODULO_LOTES.gs =====
// ============================================================
//  MODULO_LOTES.gs — v6.0
//  Rastreamento de lotes/protocolos enviados aos convênios.
//  Pipeline de status sugerido (editável — é só texto na planilha):
//   Aberto -> Enviado -> Protocolado -> Pago Parcial / Pago Total -> Glosado
// ============================================================

const LOTES_HEADERS = ['id','lote','convenio_id','convenio_nome','data_fechamento','data_envio','data_prev_recebimento','valor_total_lote','valor_recebido','valor_glosado','status','guias_ids','observacao','criado_em','atualizado_em'];
const STATUS_LOTE_VALIDOS = ['Aberto','Enviado','Protocolado','Pago Parcial','Pago Total','Glosado'];

function getLotes(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    let rows = _getSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.LOTES).map(r => _toObj(LOTES_HEADERS, r));
    if (filtros && filtros.status) rows = rows.filter(r => r.status === filtros.status);
    if (filtros && filtros.convenio_id) rows = rows.filter(r => r.convenio_id === filtros.convenio_id);
    return rows;
  } catch(e) { return []; }
}

function criarLote(dados, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor','recepcao']);
    if (!dados.convenio_id) throw new Error('Selecione o convênio do lote.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
    const id = 'LOTE' + new Date().getTime();
    const agora = new Date();
    sh.appendRow([id, dados.lote||id, dados.convenio_id, _san(dados.convenio_nome), dados.data_fechamento||'', dados.data_envio||'', dados.data_prev_recebimento||'', 0, 0, 0, 'Aberto', '', _san(dados.observacao), agora, agora]);
    _log(usuario.nome, 'NOVO_LOTE', id);
    return { ok:true, id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Vincula uma guia recém-criada a um lote existente (chamado por salvarGuia
// em Modulo_Guias.gs quando dados.lote_id vem preenchido).
function _vincularGuiaAoLote(lote_id, guia_id, valorGuia) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
  if (!sh) return;
  const allData = sh.getDataRange().getValues();
  const colGuias = LOTES_HEADERS.indexOf('guias_ids');
  const colValor = LOTES_HEADERS.indexOf('valor_total_lote');
  const colAtualizado = LOTES_HEADERS.indexOf('atualizado_em');
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === lote_id) {
      const guiasAtuais = allData[i][colGuias] ? String(allData[i][colGuias]).split(',') : [];
      guiasAtuais.push(guia_id);
      sh.getRange(i+1, colGuias+1).setValue(guiasAtuais.join(','));
      sh.getRange(i+1, colValor+1).setValue(_sanNum(allData[i][colValor]) + valorGuia);
      sh.getRange(i+1, colAtualizado+1).setValue(new Date());
      return;
    }
  }
}

function atualizarStatusLote(lote_id, status, valor_recebido, valor_glosado, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    if (STATUS_LOTE_VALIDOS.indexOf(status) === -1) throw new Error('status inválido. Use um de: ' + STATUS_LOTE_VALIDOS.join(', '));
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOTES);
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === lote_id) {
        sh.getRange(i+1, LOTES_HEADERS.indexOf('status')+1).setValue(status);
        if (valor_recebido !== undefined) sh.getRange(i+1, LOTES_HEADERS.indexOf('valor_recebido')+1).setValue(_sanNum(valor_recebido));
        if (valor_glosado !== undefined) sh.getRange(i+1, LOTES_HEADERS.indexOf('valor_glosado')+1).setValue(_sanNum(valor_glosado));
        sh.getRange(i+1, LOTES_HEADERS.indexOf('atualizado_em')+1).setValue(new Date());
        _log(usuario.nome, 'STATUS_LOTE', `${lote_id} → ${status}`);
        return { ok:true };
      }
    }
    return { ok:false, msg:'Lote não encontrado' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// Visão do "status dos lotes no mês" para o dashboard/gráfico.
function getStatusLotesMes(filtros, usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const lotes = getLotes(filtros, usuario);
    const porStatus = {};
    STATUS_LOTE_VALIDOS.forEach(s => porStatus[s] = { qtd:0, valor:0 });
    lotes.forEach(l => {
      if (!porStatus[l.status]) porStatus[l.status] = { qtd:0, valor:0 };
      porStatus[l.status].qtd++;
      porStatus[l.status].valor += _sanNum(l.valor_total_lote);
    });
    return { ok:true, porStatus, totalLotes: lotes.length };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ===== MODULO_RECEITAS_RECORRENTES.gs =====
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

// ===== MODULO_IMPORTACAO_HISTORICO.gs =====
// ============================================================
//  MODULO_IMPORTACAO_HISTORICO.gs — v2.0
//
//  CORREÇÃO CRÍTICA HERDADA: esta versão só funciona corretamente
//  porque Código_Principal.gs (v7.1) corrigiu _sanNum(), _mesDoDate()
//  e _dateStr() — que estavam quebradas desde o início do projeto
//  (nunca existiam com a assinatura certa, ou truncavam valores em
//  formato brasileiro como "1.234,56"). Se você está lendo isto e
//  ainda vê números errados no dashboard depois de importar, confirme
//  primeiro que substituiu Código_Principal.gs por esta versão.
//
//  NOVIDADES v2.0 (a pedido da gestora, depois de revisar o dashboard):
//   1) Importa também as GUIAS DE CONVÊNIO reais extraídas do relatório
//      de NFS-e emitidas (Jan-Jun/2026) — antes só os particulares e
//      despesas eram importados, os convênios ficavam de fora.
//   2) Ao final da importação, GERA AUTOMATICAMENTE os registros de
//      Recebimentos mês a mês (particular + convênio), para o Fluxo
//      de Caixa Histórico (getFluxoCaixaHistorico, em
//      Modulo_Despesas_Recorrentes.gs) já aparecer preenchido assim
//      que você importar, sem precisar rodar mais nada.
//   3) Cadastra os 4 convênios que apareciam nas notas fiscais mas não
//      existiam na aba Convênios (Banco do Brasil, CODEVASF, FUNASA,
//      Associação de Habitação e Obras Públicas) — de forma idempotente,
//      não duplica se já existirem.
//
//  IMPORTANTE — O QUE O RELATÓRIO DE NFS-e REPRESENTA:
//  O PDF de notas fiscais emitidas mostra o VALOR TOTAL faturado por
//  convênio a cada mês (às vezes uma nota só consolida vários
//  atendimentos), não guia por guia com paciente identificado. Por
//  isso, a guia importada aqui é uma "guia consolidada mensal" por
//  convênio — id de paciente vazio, observação deixando isso claro.
//
//  NÃO é dado fictício de teste: são os valores reais das notas
//  fiscais já emitidas pela clínica (Prefeitura de Aracaju,
//  competência 01/2026 a 06/2026).
//
//  IDEMPOTENTE: cada linha importada carrega um id determinístico —
//  rodar de novo não duplica nada.
//
//  RODAR UMA VEZ: importarHistoricoReal(usuario)
// ============================================================

const MAPA_NOMES_PROFISSIONAL = {
  'MILA': 'MILA PIRES', 'MILA PIRES': 'MILA PIRES',
  'MARIANA': 'MARIANA MENDONÇA', 'MARIANA MENDONÇA': 'MARIANA MENDONÇA',
  'LETICIA': 'LETICIA HELEN', 'LETICIA HELEN': 'LETICIA HELEN',
  'LETICIA HELEN ': 'LETICIA HELEN',
  'LUCAS': 'LUCAS REZENDE', 'LUCAS REZENDE': 'LUCAS REZENDE',
  'BRUNO': 'BRUNO NASCIMENTO', 'BRUNO NASCIMENTO': 'BRUNO NASCIMENTO',
  'MILENA': 'MILENA MORAES', 'MILENA MORAES': 'MILENA MORAES',
  'JOSY': 'JOSI NASCIMENTO', 'JOSY NASCIMENTO': 'JOSI NASCIMENTO',
  'GABY MAYNARD': 'GABRYELLA MAYNARD', 'GABY': 'GABRYELLA MAYNARD',
  'OSMALI': 'OSMALÍ SILVA', 'OSMALI SILVA': 'OSMALÍ SILVA',
  'VINICIUS': 'VINICIUS SOBRAL'
};

function _normalizarNomeProfissional(nomePlanilha) {
  const chave = String(nomePlanilha||'').trim().toUpperCase();
  return MAPA_NOMES_PROFISSIONAL[chave] || chave;
}

// ------------------------------------------------------------
//  CONVÊNIOS QUE APARECEM NAS NOTAS FISCAIS MAS NÃO ESTAVAM
//  CADASTRADOS NA ABA CONVÊNIOS. Cadastrados de forma idempotente
//  na função importarHistoricoReal() abaixo.
// ------------------------------------------------------------
const CONVENIOS_FALTANTES = [
  ['CONV013', 'ASSOC. HABITAÇÃO E OBRAS PÚBLICAS', 60],
  ['CONV014', 'BANESE', 60],
  ['CONV015', 'BANCO DO BRASIL', 60],
  ['CONV016', 'CODEVASF', 60],
  ['CONV017', 'FUNASA', 60]
];

// ------------------------------------------------------------
//  DADOS REAIS — extraídos de PAGAMENTOS_PARTICULARES_-_2026.xlsx
//  (abas Janeiro a Julho, 155 lançamentos, ignorando as linhas de
//  resumo por profissional que ficam no rodapé de cada aba).
//  Colunas: [paciente, data, especialidade, profissional, valor,
//            forma_pgto, quantidade]
// ------------------------------------------------------------
const PARTICULARES_HISTORICO_REAL = [
["SERGIO","2026-01-05","FISIOTERAPIA PÉLVICA","MILA PIRES",150.0,"DÉBITO","1SS"],
  ["MARIA EDITE","2026-01-07","TT DOR","MARIANA MENDONÇA",150.0,"ESPÉCIE","1SS"],
  ["MARIA JOSE DE JESUS","2026-01-07","FISIOTERAPIA MOTORA","LETICIA HELEN",550.0,"ESPÉCIE","5SS"],
  ["KESIA/NOEMI","2026-01-12","TT DOR","MARIANA MENDONÇA",440.0,"ESPÉCIE","4SS"],
  ["TASSIA","2026-01-14","FISIOTERAPIA PÉLVICA","MILA PIRES",200.0,"DÉBITO","1AV"],
  ["ROSINALVA DOS SANTOS","2026-01-14","FISIOTERAPIA MOTORA","LETICIA HELEN",130.0,"DÉBITO","1AV"],
  ["MONICA","2026-01-14","FISIOTERAPIA PÉLVICA","MILA PIRES",150.0,"ESPÉCIE","1SS"],
  ["ZENILDA","2026-01-15","TT DOR","MARIANA MENDONÇA",220.0,"CRÉDITO","1AV"],
  ["ROSINALVA DOS SANTOS","2026-01-16","FISIOTERAPIA MOTORA","LETICIA HELEN",130.0,"DÉBITO","1SS"],
  ["ROSINALVA DOS SANTOS","2026-01-19","FISIOTERAPIA MOTORA","LETICIA HELEN",600.0,"DÉBITO","5SS"],
  ["NICOLAS","2026-01-16","RPG","LUCAS REZENDE",750.0,"CRÉDITO","5SS"],
  ["LIDIA NERY","2026-01-20","LIBERAÇÃO","MARIANA MENDONÇA",130.0,"DÉBITO","1SS"],
  ["FABRICIO","2026-01-21","AVALIAÇÃO","BRUNO NASCIMENTO",130.0,"CRÉDITO","1SS"],
  ["TASSIA","2026-01-21","FISIOTERAPIA PÉLVICA","MILA PIRES",170.0,"CRÉDITO","1SS"],
  ["MARIA EDITE","2026-01-21","TT DOR","MARIANA MENDONÇA",150.0,"ESPÉCIE","1SS"],
  ["BRUNO VIEIRA","2026-01-21","LIBERAÇÃO","MARIANA MENDONÇA",120.0,"PIX","1SS"],
  ["","2026-01-22","FISIOTERAPIA MOTORA","LUCAS REZENDE",400.0,"PIX","5SS"],
  ["MARIA JOSE DE JESUS","2026-01-26","FISIOTERAPIA MOTORA","LETICIA HELEN",550.0,"ESPÉCIE","5SS"],
  ["","2026-01-22","VESTIBULAR","LUCAS REZENDE",180.0,"PIX","1AV"],
  ["MARY NADJA","2026-01-27","TT DOR","MARIANA MENDONÇA",875.0,"PIX","5SS"],
  ["LIDIA NERY","2026-01-27","LIBERAÇÃO","MARIANA MENDONÇA",130.0,"DÉBITO","1SS"],
  ["AMANDA  OLINDA","2026-01-27","LIBERAÇÃO","JOSY NASCIMENTO",180.0,"ESPÉCIE","1SS"],
  ["BERIVALDO CHAGAS","2026-01-28","FISIOTERAPIA PÉLVICA","MILA PIRES",200.0,"ESPÉCIE","AV"],
  ["MONICA","2026-01-28","FISIOTERAPIA PÉLVICA","MILA PIRES",150.0,"ESPÉCIE","1SS"],
  ["ZENILDA","2026-01-29","TT DOR","MARIANA MENDONÇA",300.0,"CRÉDITO 2X","2SS"],
  ["DALMO","2026-01-30"," AVALIAÇÃO - TT DOR","MARIANA MENDONÇA",220.0,"CREDITO ","AV"],
  ["DALMO","2026-02-02","TT DOR","MARIANA MENDONÇA",900.0,"CREDITO","5SS"],
  ["CAITANO","2026-02-03","DRENAGEM  LINFATICA","GABY MAYNARD",300.0,"PIX","5SS"],
  ["BERIVALDO","2026-02-04","FISIOTERAPIA PELVICA","MILA PIRES",150.0,"ESPECIE","1SS"],
  ["BRUNO VIEIRA","2026-02-04","LIBERAÇAO","MARIANA MENDONÇA",120.0,"PIX","1SS"],
  ["MARIA EDITE ","2026-02-04","TRATAMENTO DA DOR","MARIANA MENDONÇA",150.0,"ESPECIE","1SS"],
  ["MONICA","2026-02-04","FISIOTERAPIA PELVICA","MILA PIRES",150.0,"ESPECIE","1SS"],
  ["JOSÉ FERREIRA","2026-02-06","ACUPUNTURA","MILENA MORAES",160.0,"CREDITO","AV"],
  ["JOSÉ FERREIRA","2026-02-13","ACUPUNTURA","MILENA MORAES",700.0,"CREDITO","5SS"],
  ["LEONARDO AUGUSTO","2026-02-23","FISIOTERAPIA PELVICA","MILA PIRES",200.0,"CREDITO","AV"],
  ["MONICA","2026-02-26","FISIOTERAPIA PELVICA","MILA PIRES",150.0,"ESPECIE","1SS"],
  ["LIDIA","2026-03-02","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1SS"],
  ["Hannah Silva Linhares","2026-03-05","LIBERAÇÃO","MARIANA",120.0,"PIX","1ss"],
  ["ZENILDA","2026-03-05","TT DOR","MARIANA",450.0,"CREDITO 3X","3SS"],
  ["Felipe Olimpio","2026-03-05","RPG","JULIANNA",1500.0,"CREDITO","10SS"],
  ["Milena","2026-03-09","FISIOTERAPIA PELVICA","Mila",170.0,"Pix","1ss"],
  ["Lidia Nery","2026-03-10","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1SS"],
  ["NICOLAS","2026-03-13","RPG","LUCAS",750.0,"CREDITO","5SS"],
  ["TALITA","2026-03-16","FISIOTERAPIA PELVICA","MILA",150.0,"DEBITO","1SS"],
  ["Donvina","2026-03-18","TT DOR","MARIANA",1600.0,"CREDITO 3X","10SS"],
  ["Dalmo","2026-03-19","TT DOR","MARIANA",900.0,"CREDITO ","5SS"],
  ["EDILSON","2026-03-19","FISIOTERAPIA PELVICA","MILA",200.0,"CREDITO","AV"],
  ["Mary Nadja","2026-03-23","TT DOR","Mariana",875.0,"PIX","5ss"],
  ["Roberto Tavares","2026-03-25","LIBERAÇÃO","BRUNO",180.0,"PIX","1ss"],
  ["Edilson ","2026-03-25","FISIOTERAPIA PELVICA","Mila",150.0,"DEBITO","1ss"],
  ["Daniel","2026-03-26","Recovery","Josy",180.0,"CREDITO","1SS"],
  ["Maria Melo","2026-03-20","ACUPUNTURA","MILENA",160.0,"DEBITO","AV"],
  ["BELIVALDO CHAGAS","2026-04-01","FISIOTERAPIA PELVICA ","MILA",150.0,"ESPECIE","1 SS"],
  ["EDILSON RIBEIRO","2026-04-01","FISIOTERAPIA PELVICA ","MILA",150.0,"DEBITO","1SS"],
  ["MONICA","2026-04-02","FISIOTERAPIA PELVICA ","MILA",150.0,"ESPECIE","1SS"],
  ["LIDIA","2026-04-07","LIBERAÇAO","MARIANA",130.0,"DEBITO","1SS"],
  ["BELIVALDO CHAGAS","2026-04-08","FISIOTERAPIA PELVICA ","MILA",150.0,"ESPECIE","1SS"],
  ["EDILSON RIBEIRO","2026-04-08","FISIOTERAPIA PELVICA ","MILA",150.0,"DEBITO","1ss"],
  ["ZENILDA RIBEIRO","2026-04-09","TT DOR","MARIANA",750.0,"CREDITO 3X","5SS"],
  ["ELIEZER BITTENCOURT","2026-04-14","LIBERAÇAO","OSMALI",160.0,"ESPECIE","1SS"],
  ["BELIVALDO CHAGAS","2026-04-15","FISIOTERAPIA PELVICA ","MILA",150.0,"ESPECIE","1ss"],
  [" EDILSON RIBEIRO","2026-04-15","FISIOTERAPIA PELVICA ","MILA",150.0,"DEBITO","1ss"],
  ["ELIEZER BITTENCOURT","2026-04-16","LIBERAÇAO","OSMALI",160.0,"PIX","1ss"],
  ["ALEXIA MORAES","2026-04-20","DTM","MARIANA",180.0,"CREDITO ","1SS"],
  ["ALEXSANDRA NUNES","2026-04-20","LIBERAÇAO","MARIANA",110.0,"PIX","1SS"],
  ["BELIVALDO CHAGAS","2026-04-22","FISIOTERAPIA PELVICA ","MILA",150.0,"ESPECIE","1SS"],
  ["EDILSON RIBEIRO","2026-04-22","FISIOTERAPIA PELVICA ","MILA",150.0,"DEBITO","1ss"],
  ["TALITA SIQUEIRA","2026-04-23","FISIOTERAPIA PELVICA ","MILA",150.0,"DEBITO","1SS"],
  ["ANTONELLA  MENESES","2026-04-24","ACUPUNTURA","MILENA",160.0,"PIX","AV"],
  ["ROMILDES MACHADO","2026-04-27","CONSULTA ORTOPEDISTA","VINICIUS",200.0,"PIX","1SS"],
  ["JONAS SILVA","2026-04-27","CONSULTA ORTOPEDISTA","VINICIUS",400.0,"CREDITO ","1SS"],
  ["LIDIA","2026-04-28","LIBERAÇAO","MARIANA",130.0,"DEBITO","1ss"],
  ["GUILHERME ANDRADE","2026-04-08","FISIOTERAPIA MOTORA","MARIANA",1750.0,"PIX","10SS"],
  ["JULIANA MARTINS","2026-04-24","LIBERAÇAO","MARIANA",150.0,"PIX","1SS"],
  ["DONVINA OLIVEIRA","2026-04-29","FISIOTERAPIA MOTORA","MARIANA",1200.0,"CREDITO 4X","10SS"],
  ["JULIANA MARTINS","2026-04-29","LIBERAÇAO","MARIANA",150.0,"PIX","1SS"],
  ["MARY NADJA","2026-04-29","FISIOTERAPIA MOTORA","MARIANA",875.0,"PIX","5SS"],
  ["JOAQUIM JORDAO ","2026-05-04","FISIOTERAPIA MOTORA","OSMALI",130.0,"PIX","AV"],
  ["LIDIA TEIXERA","2026-05-05","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1SS"],
  ["PEDRO NETO","2026-05-05","DTM","GABRYELLA ",180.0,"CREDITO","AV"],
  ["ELISABETH TEIXEIRA","2026-05-06","LIBERAÇÃO","MARIANA",625.0,"PIX","5SS"],
  ["GENISSON BARRETO","2026-05-07","LIBERAÇÃO","JOSY",180.0,"CREDITO","1SS"],
  ["JOAQUIM JORDAO ","2026-05-08","FISIOTERAPIA MOTORA","OSMALI",125.0,"CREDITO","1ss"],
  ["JULIETA SOARES ","2026-05-11","FISIOTERAPIA PELVICA","MILA",200.0,"CREDITO","AV"],
  ["VILMA LIMA","2026-05-11","TT DOR","MARIANA",220.0,"PIX","AV"],
  ["RONALDO LINHARES","2026-05-12","TT DOR","MARIANA",400.0,"PIX","5SS"],
  ["LIDIA TEIXERA","2026-05-12","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1SS"],
  ["PEDRO NETO","2026-05-12","DTM","GABRYELLA ",750.0,"CREDITO 2X","5SS"],
  ["MARIA CONSUELO SILVA","2026-05-13","TT DOR","MARIANA",150.0,"DEBITO","1SS"],
  ["VILMA LIMA","2026-05-13","TT DOR","MARIANA",900.0,"CREDITO 2X","5SS"],
  ["JOAQUIM JORDAO ","2026-05-13","FISIOTERAPIA MOTORA","OSMALI",125.0,"CREDITO","1SS"],
  ["PEDRO DRUMMOND","2026-05-14","LIBERAÇÃO","GABRYELLA ",160.0,"DEBITO","1SS"],
  ["TATIANE SANTOS","2026-05-14","FISIOTERAPIA PELVICA","MILA",200.0,"DEBITO","AV"],
  ["JOAQUIM JORDAO ","2026-05-15","FISIOTERAPIA MOTORA","OSMALI",600.0,"CREDITO","5SS"],
  ["ANA VALERIA SANTOS ","2026-05-18","FISIOTERAPIA PELVICA","MILA",200.0,"PIX","AV"],
  ["GENISSON BARRETO","2026-05-18","LIBERAÇÃO","JOSY",180.0,"CREDITO","1SS"],
  ["HILDETE CROPPER","2026-05-18","RPG","LUCAS",930.0,"DEBITO","AV 5SS"],
  ["LIDIA TEIXERA","2026-05-19","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1ss"],
  ["BELIVALDO CHAGAS","2026-05-20","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["MAURICIO SALMERON","2026-05-20","TT DOR","MARIANA",200.0,"ESPECIE","1SS"],
  ["JULIETA SOARES ","2026-05-20","FISIOTERAPIA PELVICA","MILA",1500.0,"PIX","10SS"],
  ["SANDRA MARIA","2026-05-21","TT DOR","MARIANA",135.0,"CREDITO","1SS"],
  ["DIEGO MATOS","2026-05-22","LIBERAÇÃO","OSMALI",160.0,"PIX","1ss"],
  ["TERESINHA LEMOS","2026-05-22","TT DOR","MARIANA",360.0,"PIX","2ss"],
  ["MARIA CONSUELO SILVA","2026-05-22","TT DOR","MARIANA",150.0,"ESPECIE","1SS"],
  ["ANTONELLA MENESES","2026-05-22","ACUPUNTURA","MILENA",700.0,"CREDITO 4X","5SS"],
  ["MARIA AUXILIADORA SILVA ","2026-05-25","CONSULTA ORTOPEDISTA","VINICIUS",200.0,"DEBITO","1SS"],
  ["LIDIA TEIXERA","2026-05-26","LIBERAÇÃO","MARIANA",130.0,"DEBITO","1SS"],
  ["BELIVALDO CHAGAS","2026-05-27","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["EDILSON RIBEIRO","2026-05-27","FISIOTERAPIA PELVICA","MILA",150.0,"DEBITO","1SS"],
  ["MARIA CONSUELO SILVA","2026-05-27","TT DOR","MARIANA",450.0,"CREDITO 2X","3SS"],
  ["ISRAEL SILVA","2026-05-28","FISIOTERAPIA MOTORA","LUCAS",800.0,"PIX","10SS"],
  ["ANA VALERIA SANTOS","2026-06-01","FISIOTERAPIA PELVICA","MILA ",170.0,"PIX","1SS"],
  ["IVALCY COSTA DE OLIVEIRA","2026-06-03","TT DOR","MARIANA ",1120.0,"ESPECIE","AV 5SS"],
  ["ELISABETH TEIXEIRA","2026-06-03","TT DOR","MARIANA ",625.0,"PIX","5SS"],
  ["EDILSON RIBEIRO","2026-06-03","FISIOTERAPIA PELVICA","MILA ",150.0,"DEBITO","1SS"],
  ["JOAQUIM JORDÃO","2026-06-05","FISIOTERAPIA MOTORA","OSMALI",600.0,"CREDITO","5SS"],
  ["OSMALI SILVA","2026-06-08","CONSULTA ORTOPEDISTA","VINICIUS",200.0,"PIX","1SS"],
  [" DENIA CAMILA SILVA","2026-06-10","FISIOTERAPIA MOTORA","BRUNO",130.0,"PIX","AV "],
  ["BELIVALDO CHAGAS","2026-06-10","FISIOTERAPIA PELVICA","MILA ",150.0,"ESPECIE","1SS"],
  ["RENILDE MARINO","2026-06-10","TT DOR","MARIANA ",150.0,"DEBITO","1SS"],
  ["MARIA AUXILIADORA COUTO","2026-06-10","TT DOR","MARIANA ",1970.0,"CREDITO 4X","AV 10SS"],
  ["EDILSON RIBEIRO","2026-06-10","FISIOTERAPIA PELVICA","MILA ",150.0,"DEBITO","1SS"],
  ["ALEXANDRE SOUZA","2026-06-15","FISIOTERAPIA MOTORA","LETICIA",130.0,"PIX","1SS"],
  ["BELIVALDO CHAGAS","2026-06-17","FISIOTERAPIA PELVICA","MILA ",150.0,"ESPECIE","1SS"],
  ["RANDY ORTEGA","2026-06-18","FISIOTERAPIA MOTORA","BRUNO",130.0,"CREDITO","AV"],
  ["NOEMI /KESIA","2026-06-22","LIBERAÇÃO","MARIANA ",110.0,"ESPECIE","1SS"],
  ["MARIA CONSUELO SILVA","2026-06-26","TT DOR","MARIANA ",100.0,"ESPECIE","1SS"],
  ["GUARACIABA APARECIDA OLIVA","2026-06-29","FISIOTERAPIA PELVICA","MILA ",200.0,"ESPECIE","AV"],
  ["MARY NADJA SEABRA","2026-06-29","TT DOR","MARIANA ",875.0,"PIX","5SS"],
  ["MARIA CONSUELO SILVA","2026-06-29","TT DOR","MARIANA ",700.0,"CREDITO 3X","9SS"],
  ["NILTON MOREIRA","2026-06-30","CONSULTA ORTOPEDISTA","VINICIUS",400.0,"CREDITO","1SS"],
  ["BELIVALDO CHAGAS","2026-07-01","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["SANDRA MARIA","2026-07-01","TT DOR","MARIANA",135.0,"CREDITO","1SS"],
  ["EDILSON RIBEIRO","2026-07-01","FISIOTERAPIA PELVICA","MILA",150.0,"DEBITO","1SS"],
  ["ELISABETH TEIXEIRA","2026-07-09","TT DOR","MARIANA",625.0,"PIX","5SS"],
  ["GUARACIABA APARECIDA OLIVA","2026-07-13","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["ALYSSON ALBERTO MELO","2026-07-13","FISIOTERAPIA MOTORA","LETICIA",130.0,"DEBITO","AV"],
  ["MARLY MARIA SANTOS","2026-07-14","RPG","LUCAS",180.0,"CREDITO","AV"],
  ["GENAMARA PEREIRA","2026-07-15","LIBERAÇÃO","JULIANA",180.0,"ESPECIE","1SS"],
  ["BELIVALDO CHAGAS","2026-07-15","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["EDILSON RIBEIRO","2026-07-15","FISIOTERAPIA PELVICA","MILA",150.0,"DEBITO","1SS"],
  ["ALYSSON ALBERTO MELO","2026-07-15","FISIOTERAPIA MOTORA","LETICIA",600.0,"CREDITO 2X","5SS"],
  ["IVALCY COSTA","2026-07-16","TT DOR","MARIANA",900.0,"ESPECIE","5SS"],
  ["MARIA ANTONIA BULHOES","2026-07-20","FISIOTERAPIA AQUATICA","JULIA",1000.0,"ESPECIE","10SS"],
  ["RANDY ORTEGA","2026-07-20","FISIOTERAPIA MOTORA","BRUNO",600.0,"DEBITO","5SS"],
  ["MARLY MARIA SANTOS","2026-07-21","RPG","LUCAS",750.0,"CREDITO 2X","5SS"],
  ["EDILSON RIBEIRO","2026-07-22","FISIOTERAPIA PELVICA","MILA",150.0,"DEBITO","1SS"],
  ["SONIA REGINA BUENO","2026-07-22","LIBERAÇÃO","JOSY",180.0,"PIX","1SS"],
  ["VINICIUS LOIOLA","2026-07-24","LIBERAÇÃO","JOSY",160.0,"CREDITO","1SS"],
  ["SANDRA MARIA","2026-07-24","TT DOR","MARIANA",850.0,"CREDITO","10SS"],
  ["MARIA AUXILIADORA COUTO","2026-07-24","TT DOR","MARIANA",1750.0,"1000 PIX 750 CREDITO 3X","10SS"],
  ["GUARACIABA APARECIDA OLIVA","2026-07-28","FISIOTERAPIA PELVICA","MILA",150.0,"ESPECIE","1SS"],
  ["ANDREA MOTTA","2026-07-28","FISIOTERAPIA PELVICA","MILA",150.0,"PIX","1ss"],
  ["SONIA REGINA BUENO","2026-07-29","TT DOR","MARIANA",220.0,"PIX","AV"]
];

// ------------------------------------------------------------
//  DADOS REAIS — extraídos de DESPESAS__1_.xlsx (abas Janeiro a
//  Julho), SEÇÃO OPERACIONAL, excluindo linhas de subtotal.
//  Colunas: [descricao, valor, data]
// ------------------------------------------------------------
const DESPESAS_OPERACIONAIS_HISTORICO_REAL = [
["CARTAO CLINICA",3000.0,"2026-01-05"],
  ["BOMBEIROS",157.6,"2026-01-06"],
  ["PONTO ELETRONICO",60.0,"2026-01-07"],
  ["GIL - PEDREIRO",560.0,"2026-01-08"],
  ["CARRO",2811.35,"2026-01-08"],
  ["ART PGRSS",108.39,"2026-01-09"],
  ["ALUGUEL CLINICA",2000.0,"2026-01-09"],
  ["COND GALERIA",660.0,"2026-01-10"],
  ["GIL - PEDREIRO",300.0,"2026-01-11"],
  ["CARRO VOLKSWAGEN",1536.58,"2026-01-14"],
  ["PGRSS - LUCAS ART",1000.0,"2026-01-14"],
  ["CARTAO CLINICA",3000.0,"2026-01-15"],
  ["CARTAO CLINICA - CAIXA",2576.11,"2026-01-15"],
  ["ART PMOC",108.39,"2026-01-16"],
  ["DEDETIZAÇÃO",500.0,"2026-01-19"],
  ["SASE",182.92,"2026-01-20"],
  ["ALUGUEL GALERIA",3000.0,"2026-01-20"],
  ["ENERGISA - SALA 5",146.28,"2026-01-20"],
  ["ENERGISA - SALA 6",111.82,"2026-01-20"],
  ["ENERGISA - CLINICA",1233.43,"2026-01-20"],
  ["SIMPLES - DASN",15901.59,"2026-01-20"],
  ["FGTS",659.11,"2026-01-20"],
  ["CONTABILIDADE",1518.0,"2026-01-20"],
  ["DARF",641.53,"2026-01-20"],
  ["CEMED",130.0,"2026-01-20"],
  ["INVS CLINICA",917.38,"2026-01-26"],
  ["UNIMED",38.58,"2026-01-26"],
  ["UNIMED",923.63,"2026-01-30"],
  ["ENERGISA - SALA 5",86.29,"2026-02-03"],
  ["ENERGISA - SALA 6",120.99,"2026-02-03"],
  ["CHAVEIRO",60.0,"2026-02-03"],
  ["CEMED",130.0,"2026-02-04"],
  ["ALUGUEL CLINICA",2000.0,"2026-02-04"],
  ["TAXAS BANCO NORDESTE",1605.0,"2026-02-04"],
  ["SASE",173.96,"2026-02-04"],
  ["IPTU 2025 - CLINICA",2063.38,"2026-02-04"],
  ["IPTU 2026 - CLINICA",664.84,"2026-02-04"],
  ["IPTU - GALERIA",174.05,"2026-02-04"],
  ["IPTU - GALERIA",158.41,"2026-02-04"],
  ["CARRO",2803.41,"2026-02-04"],
  ["TISS XML",60.68,"2026-02-04"],
  ["CARTAO CLINICA - CAIXA",2566.74,"2026-02-04"],
  ["CERTIFICADO DIGITAL",245.0,"2026-02-05"],
  ["COND GALERIA",700.0,"2026-02-05"],
  ["CARTAO CLINICA",7454.7,"2026-02-16"],
  ["SIMPLES - DASN",16576.21,"2026-02-20"],
  ["ALUGUEL GALERIA",3000.0,"2026-02-20"],
  ["CARRO - HB20",1729.31,"2026-02-23"],
  ["IGUA - TAXA SEMA",174.41,"2026-02-23"],
  ["ENERGISA - CLINICA",932.12,"2026-02-24"],
  ["IGUÁ - CLÍNICA",287.68,"2026-02-24"],
  ["FGTS",982.12,"2026-02-25"],
  ["INVEST CLINICA",918.71,"2026-03-02"],
  ["COPA",363.45,"2026-03-04"],
  ["ENERGISA - SALA 5",86.29,"2026-03-04"],
  ["ENERGISA - SALA 6",120.99,"2026-03-04"],
  ["TISS XML",60.65,"2026-03-05"],
  ["CARTAO CLINICA - CAIXA",2343.5,"2026-03-05"],
  ["MARCOS AGUA",118.0,"2026-03-06"],
  ["CEMED",130.0,"2026-03-06"],
  ["IPTU GALERIA",174.05,"2026-03-06"],
  ["IPTU GALERIA",158.41,"2026-03-06"],
  ["IPTU CLINICA",664.84,"2026-03-06"],
  ["SASE",175.08,"2026-03-06"],
  ["ALUGUEL CLINICA",2000.0,"2026-03-06"],
  ["INOVE - JULHO 2025",116.96,"2026-03-09"],
  ["COND GALERIA",700.0,"2026-03-10"],
  ["SIMPLES PARC FEV",1163.8,"2026-03-10"],
  ["SIMPLES PARC FEV",1163.8,"2026-03-12"],
  ["CONTABILIDADE",1621.0,"2026-03-15"],
  ["CARRO",2827.23,"2026-03-16"],
  ["CARRO - HB20",1544.48,"2026-03-16"],
  ["CARTAO CLINICA",2292.55,"2026-03-16"],
  ["AR CONDICIONADO - SALA 04",300.0,"2026-03-19"],
  ["SEGURO BNB",806.4,"2026-03-19"],
  ["FGTS",451.36,"2026-03-20"],
  ["TISS XML",55.0,"2026-03-20"],
  ["INSS",656.36,"2026-03-20"],
  ["SIMPLES",14068.86,"2026-03-20"],
  ["ALUGUEL GALERIA",3000.0,"2026-03-20"],
  ["CONTABILIDADE",1621.0,"2026-03-26"],
  ["INVEST CLINICA",909.15,"2026-03-26"],
  ["INVEST CLINICA",909.85,"2026-03-26"],
  ["CARRO EMP EMPRESTIMO",38697.0,"2026-03-30"],
  ["IGUÁ - CLINICA",295.43,"2026-03-31"],
  ["ENERGISA - CLINICA",1106.45,"2026-03-31"],
  ["ENERGISA - SALA 06",295.74,"2026-03-31"],
  ["ENERGISA - SALA 05",210.45,"2026-03-31"],
  ["ENERGISA - CLINICA",1115.79,"2026-04-01"],
  ["ARQUITETA - CLINICA",950.0,"2026-04-01"],
  ["UNIMED",924.23,"2026-04-02"],
  ["APLICAÇÃO PISO",700.0,"2026-04-02"],
  ["CARTAO CLINICA - CAIXA",3128.2,"2026-04-07"],
  ["IPTU GALERIA - SALA 06",177.41,"2026-04-08"],
  ["IPTU GALERIA - SALA 05",174.05,"2026-04-08"],
  ["ALUGUEL - CLINICA",2000.0,"2026-04-08"],
  ["APLICAÇÃO PISO - SALA 01",350.0,"2026-04-08"],
  ["CASA DAS TINTAS",303.3,"2026-04-09"],
  ["CEMED",130.0,"2026-04-09"],
  ["SASE",187.14,"2026-04-09"],
  ["CAPITAL - CLINICA",1427.26,"2026-04-09"],
  ["COND GALERIA",700.0,"2026-04-09"],
  ["IPTU - CLINICA",667.03,"2026-04-09"],
  ["PINTURA - CLINICA",1300.0,"2026-04-13"],
  ["ESTORNO PAC PILATES - GABRIEL BARRETO",260.0,"2026-04-14"],
  ["CARRO - HB20",1544.48,"2026-04-16"],
  ["IGUA - CLINICA",214.36,"2026-04-20"],
  ["SIMPLES - DASN",14920.87,"2026-04-20"],
  ["DARF",656.36,"2026-04-20"],
  ["FGTS",451.36,"2026-04-20"],
  ["FGTS",451.36,"2026-04-20"],
  ["SIMPLES - DASN",14920.87,"2026-04-20"],
  ["CARTAO - VISA",3183.59,"2026-04-20"],
  ["GIL PEDREIRO - CLINICA",400.0,"2026-04-20"],
  ["ALUGUEL - GALERIA",3000.0,"2026-04-20"],
  ["CONTABILIDADE",1621.0,"2026-04-20"],
  ["INVES CLINICA",907.83,"2026-04-20"],
  ["MARCOS AGUA",120.0,"2026-04-23"],
  ["INSTALAÇÃO PISO - SALA 04",380.0,"2026-04-24"],
  ["TISS XML",55.0,"2026-04-29"],
  ["ENERGISA - SALA 06 (abril)",199.84,"2026-05-03"],
  ["ENERGISA - SALA 05 (abril)",104.09,"2026-05-03"],
  ["CEMED",130.0,"2026-05-04"],
  ["SASE",180.0,"2026-05-04"],
  ["PROJETO - FACHADA",950.0,"2026-05-04"],
  ["SIMPLES - DASN PARC ABRIL",187.74,"2026-05-05"],
  ["EMPREST - CLINICA",1375.91,"2026-05-05"],
  ["CARTAO ELO",3905.57,"2026-05-05"],
  ["SIMPLES - DASN PARC MARÇO",1185.04,"2026-05-08"],
  ["SIMPLES DASN PARC MARÇO",187.74,"2026-05-08"],
  ["IPTU - CLINICA",664.84,"2026-05-08"],
  ["IPTU - SALA 05",158.41,"2026-05-08"],
  ["IPTU - SALA 06",174.05,"2026-05-08"],
  ["PARC SIMPLES",1185.04,"2026-05-08"],
  ["PARC SIMPLES",187.74,"2026-05-08"],
  ["PARC SIMPLES",187.74,"2026-05-08"],
  ["UNIMED",1003.98,"2026-05-08"],
  ["ALUGUEL - CLINICA",2000.0,"2026-05-10"],
  ["COND GALERIA",700.0,"2026-05-11"],
  ["CARTAO",2000.0,"2026-05-12"],
  ["CARRO HB20",1498.72,"2026-05-12"],
  ["PONTO ELETRONICO",210.0,"2026-05-14"],
  ["CARTAO VISA",1823.61,"2026-05-15"],
  ["ENERGISA - SALA 05 (maio)",78.82,"2026-05-20"],
  ["ENERGISA - SALA 06 (maio)",188.12,"2026-05-20"],
  ["ENERGISA - CLINICA",1241.49,"2026-05-20"],
  ["FGTS",581.04,"2026-05-20"],
  ["CONTABILIDADE",1621.0,"2026-05-20"],
  ["SIMPLES - DASN",15492.83,"2026-05-20"],
  ["DARF - INSS",777.93,"2026-05-20"],
  ["ALUGUEL - PILATES",3000.0,"2026-05-20"],
  ["VIVO - INTERNET+FIXO",279.83,"2026-05-25"],
  ["PLACAS - DIVULGAÇÃO CLINICA",600.0,"2026-05-28"],
  ["CONSULTORIA FINANCEIRA",2800.0,"2026-06-01"],
  ["INVEST CLINICA",910.75,"2026-06-01"],
  ["INVEST CLINICA",884.57,"2026-06-01"],
  ["UNIMED",1001.74,"2026-06-01"],
  ["PROJETO - FACHADA",950.0,"2026-06-02"],
  ["RESERVA - CLINICA",1512.0,"2026-06-03"],
  ["ALUGUEL - CLINICA",2000.0,"2026-06-08"],
  ["TISS XML",60.8,"2026-06-09"],
  ["CEMED",130.0,"2026-06-09"],
  ["CARTAO - CLINICA",1240.0,"2026-06-09"],
  ["SASE",187.14,"2026-06-09"],
  ["IPTU - GALERIA",158.93,"2026-06-09"],
  ["IPTU - GALERIA",174.62,"2026-06-09"],
  ["IPTU - CLINICA",667.03,"2026-06-09"],
  ["TLF",196.23,"2026-06-09"],
  ["CARTAO ELO",3378.3,"2026-06-09"],
  ["COND GALERIA",700.0,"2026-06-09"],
  ["EMPREST CLINICA",1433.22,"2026-06-10"],
  ["CURSO RPG - MARIANA",2000.0,"2026-06-15"],
  ["CARTAO VISA",2848.16,"2026-06-15"],
  ["IGUA - CLINICA",204.9,"2026-06-19"],
  ["SIMPLES - DASN",15379.24,"2026-06-19"],
  ["DARF",777.93,"2026-06-19"],
  ["SIMPLES - DASN",189.01,"2026-06-19"],
  ["CARRO HB20",1556.32,"2026-06-19"],
  ["MANUTENÇÃO AR CONDICIONADO",350.0,"2026-06-19"],
  ["HAPVIDA (maio)",335.36,"2026-06-20"],
  ["ALUGUEL - PILATES",3000.0,"2026-06-22"],
  ["CONTABILIDADE",1621.0,"2026-06-23"],
  ["ENERGISA - CLINICA",963.57,"2026-06-23"],
  ["TISS XML",55.0,"2026-06-23"],
  ["SIMPLES - DASN",189.01,"2026-06-30"],
  ["SIMPLES - DASN",1194.92,"2026-06-30"],
  ["HAPVIDA",292.18,"2026-06-30"],
  ["ENERGISA - SALA 05",123.59,"2026-06-30"],
  ["ENERGISA - SALA 06",179.9,"2026-06-30"],
  ["CONTABILIDADE - DECIMO",1621.0,"2026-06-30"],
  ["AJUSTES PILATES",100.0,"2026-06-19"],
  ["CONSULTORIA FINANCEIRA",2800.0,"2026-06-30"],
  ["PROJETO - FACHADA",950.0,"2026-06-30"],
  ["ALUGUEL - CLINICA",2000.0,"2026-07-03"],
  ["CARTAO ELO",4263.8,"2026-07-06"],
  ["PREVIDENCIA - MARIANA",237.0,"2026-07-06"],
  ["INOVE",100.0,"2026-07-06"],
  ["IPTU - CLINICA",664.84,"2026-07-06"],
  ["IPTU - GALERIA",158.41,"2026-07-06"],
  ["IPTU - GALERIA",174.05,"2026-07-06"],
  ["VIVO - FIXO",19.88,"2026-07-09"],
  ["VIVO - MOVEL",49.0,"2026-07-09"],
  ["VIVO - FIBRA",251.46,"2026-07-09"],
  ["COND GALERIA",912.0,"2026-07-10"],
  ["PAPEL TIMBRADO - 50%",85.0,"2026-07-10"],
  ["MARCOS AGUA",120.0,"2026-07-10"],
  ["JALECO - 50%",247.9,"2026-07-10"],
  ["MANUTENÇÃO ESTOFADO",1390.0,"2026-07-10"],
  ["MANUTENÇÃO PILATES",300.0,"2026-07-10"],
  ["PLACAS - MKT",600.0,"2026-07-10"],
  ["CARTAO CLINICA",1052.0,"2026-07-10"],
  ["CARTAO PILATES",340.0,"2026-07-10"],
  ["CARTAO CLINICA 2",328.0,"2026-07-10"],
  ["SIMPLES - DASN PARC MARÇO",1205.26,"2026-07-14"],
  ["EMPREST BNB",663.0,"2026-07-14"],
  ["BLUSAS PILATES",660.0,"2026-07-16"],
  ["TECHPRINT - MANUTENÇÃO IMPRESSORA",170.0,"2026-07-17"],
  ["CARTAO VISA",1107.46,"2026-07-17"],
  ["CARRO HB20",1548.42,"2026-07-17"],
  ["FGTS",594.07,"2026-07-20"],
  ["DARF - INSS",792.59,"2026-07-20"],
  ["SIMPLES - DASN JUNHO",14920.36,"2026-07-20"],
  ["ALUGUEL - PILATES",3152.0,"2026-07-22"],
  ["ENERGISA - SALA 06",122.12,"2026-07-24"],
  ["ENERGISA - SALA 05",107.7,"2026-07-24"],
  ["ENERGISA - CLINICA",960.3,"2026-07-24"],
  ["IGUA - CLINICA",184.21,"2026-07-24"],
  ["CONTABILIDADE",1621.0,"2026-07-24"],
  ["TISS XML",55.0,"2026-07-24"],
  ["SASE",195.99,"2026-07-24"],
  ["PLACAS ACRILICO",66.7,"2026-07-24"]
];

// ------------------------------------------------------------
//  DADOS REAIS — extraídos de DESPESAS__1_.xlsx, SEÇÃO
//  "PROFISSIONAIS" (repasses/comissões já pagas).
//  Colunas: [descricao, valor, data]
// ------------------------------------------------------------
const REPASSES_PROFISSIONAIS_HISTORICO_REAL = [
["OSMALI SILVA",971.0,"2026-01-06"],
  ["HANAMEEL VIEIRA",180.0,"2026-01-06"],
  ["LETICIA HELEN",2386.77,"2026-01-06"],
  ["LUCAS SILVA",2000.0,"2026-01-06"],
  ["MAYARA SOUZA",1649.15,"2026-01-06"],
  ["ROSILAINE DE JESUS (férias+2 dias comprados+emp 10+salario+passagem)",2263.13,"2026-01-06"],
  ["MARIANA MENDONÇA",2000.0,"2026-01-06"],
  ["APARECIDA - REABILIT",3546.12,"2026-01-08"],
  ["BRUNO NASCIMENTO",2218.18,"2026-01-08"],
  ["GABRYELLA MAYNARD",2527.4,"2026-01-08"],
  ["JAMILLE GONÇALVES",1599.84,"2026-01-08"],
  ["JOSY NASCIMENTO",1489.45,"2026-01-08"],
  ["JULIANNA ARCHIMINO",3883.33,"2026-01-08"],
  ["LAIS MARINHO",2705.85,"2026-01-08"],
  ["LUCAS SILVA - CONVENIOS",9519.05,"2026-01-08"],
  ["LUCIANA DANTAS - FONO",137.5,"2026-01-08"],
  ["MILA PIRES - CONVENIOS",7998.72,"2026-01-08"],
  ["MILENA MORAIS - CLINICA",1180.44,"2026-01-08"],
  ["VIVIANE COSTA ",558.12,"2026-01-08"],
  ["YASMIN SILVA - AQUATICA",5454.15,"2026-01-08"],
  ["JULIANA LINHARES",540.0,"2026-01-08"],
  ["YASMIN SILVA",195.72,"2026-01-08"],
  ["JANINY DE JESUS - MTK DIGITAL CLINICA",450.0,"2026-01-20"],
  ["JANINY DE JESUS - MTK DIGITAL PILATES",450.0,"2026-01-20"],
  ["APARECIDA - REABILIT",3421.48,"2026-02-03"],
  ["GRACE - HIDRO",6291.91,"2026-02-03"],
  ["JOSY NASCIMENTO",1322.93,"2026-02-03"],
  ["OSMALI SILVA",1024.27,"2026-02-03"],
  ["LETICIA HELEN",2388.32,"2026-02-03"],
  ["LUCAS SILVA - PRO LABORE",2000.0,"2026-02-03"],
  ["MAYARA SOUZA",1746.97,"2026-02-03"],
  ["HANAMEEL SILVA",180.0,"2026-02-03"],
  ["YASMIN SILVA",587.16,"2026-02-03"],
  ["MILA PARTICULAR",330.0,"2026-02-03"],
  ["MILA CONVENIOS",5953.44,"2026-02-03"],
  ["BRUNO NASCIMENTO",3650.62,"2026-02-04"],
  ["GABRYELLA MAYNARD",3509.0,"2026-02-04"],
  ["YASMIN - HIDRO",5599.71,"2026-02-04"],
  ["VIVIANE COSTA",1479.28,"2026-02-04"],
  ["JULIANNA ARCHIMINO",2494.48,"2026-02-04"],
  ["MILENA MORAIS - CLINICA",1223.49,"2026-02-04"],
  ["MILENA MORAIS - SOCORRO",394.78,"2026-02-04"],
  ["NIVANIA ADELINO - FONO",537.5,"2026-02-04"],
  ["LUCAS SILVA - CONVENIOS",8010.2,"2026-02-04"],
  ["MARIANA CONVENIOS",5000.0,"2026-02-04"],
  ["ROSILAINE DE JESUS",1588.41,"2026-02-04"],
  ["ANA BEATRIZ",421.74,"2026-02-04"],
  ["CAMILLA BENIGNO",292.5,"2026-02-13"],
  ["JAMILLE GONÇALVES ",1532.4,"2026-02-13"],
  ["ALANNA LETICIA - RESIDUAL",14.12,"2026-02-13"],
  ["BRENNA MIRELLE",18.0,"2026-02-13"],
  ["ANDRESSA DE SOUZA",32.5,"2026-02-13"],
  ["LUCIANA DANTAS - FONO",125.0,"2026-02-13"],
  ["MILA PIRES",820.44,"2026-02-13"],
  ["JANINY DE JESUS ",950.0,"2026-02-20"],
  ["ROSILAINE DE JESUS",1755.97,"2026-03-06"],
  ["MAYARA SOUZA",1755.97,"2026-03-06"],
  ["LETICIA HELEN",2397.32,"2026-03-06"],
  ["LUCAS PRO-LABORE",2000.0,"2026-03-06"],
  ["ANA BEATRIZ",1688.54,"2026-03-06"],
  ["HANAMEEL VIEIRA",150.0,"2026-03-06"],
  ["OSMALI SILVA",814.27,"2026-03-06"],
  ["APARECIDA - REABILIT",2235.24,"2026-03-06"],
  ["BRUNO NASCIMENTO",2375.04,"2026-03-06"],
  ["CAMILLA BENIGNO",650.0,"2026-03-06"],
  ["GABRYELLA MAYNARD",2170.39,"2026-03-06"],
  ["JULIANA LINHARES",504.0,"2026-03-06"],
  ["JULIANNA ARCHIMINO",828.7,"2026-03-06"],
  ["LAIS MARINHO",620.15,"2026-03-06"],
  ["LUCAS REZENDE - CONVENIOS",4144.8,"2026-03-06"],
  ["LUCIANA DANTAS",137.5,"2026-03-06"],
  ["MILA PIRES - CONVENIOS",6457.68,"2026-03-06"],
  ["MILENA - CLINICA",1150.44,"2026-03-06"],
  ["MILENA - SOCORRO",955.2,"2026-03-06"],
  ["VIVIANE COSTA",558.12,"2026-03-06"],
  ["YASMIN - CONVENIOS",832.08,"2026-03-06"],
  ["LUCAS REZENDE - HIDRO",1541.76,"2026-03-10"],
  ["JOSY NASCIMENTO",1185.45,"2026-03-10"],
  ["YASMIN - HIDRO",4082.0,"2026-03-10"],
  ["LUCAS - RPG PARTICULAR",690.0,"2026-03-16"],
  ["JANINY BRAZ",400.0,"2026-03-20"],
  ["JANINY BRAZ",350.0,"2026-03-20"],
  ["MILA PIRES - PARTICULAR",655.0,"2026-03-20"],
  ["MILENA MORAIS - PARTICULAR",430.0,"2026-03-20"],
  ["MAYARA SOUZA",1746.0,"2026-04-02"],
  ["ROSILAINE DE JESUS",1746.0,"2026-04-02"],
  ["OSMALI SILVA",2200.0,"2026-04-02"],
  ["LUCAS PRO LABORE",2000.0,"2026-04-02"],
  ["LETICIA HELEN",2388.32,"2026-04-02"],
  ["STEFANY DE OLIVEIRA (dias trabalhados+passagem mar e abril)",1038.0,"2026-04-02"],
  ["GILVANIRA AZEVEDO (dias trabalhados+passagrm março)",245.16,"2026-04-02"],
  ["HANAMEEL DIAS",150.0,"2026-04-02"],
  ["APARECIDA - REABILIT",6371.32,"2026-04-08"],
  ["BRUNO NASCIMENTO",2922.42,"2026-04-08"],
  ["GRACE - HIDRO",6281.72,"2026-04-08"],
  ["JULIANNA ARCHIMINO",3414.45,"2026-04-08"],
  ["JULIANNA LINHARES",648.0,"2026-04-08"],
  ["MILA PIRES - CONVENIOS",9972.06,"2026-04-08"],
  ["JOSY NASCIMENTO",1260.59,"2026-04-08"],
  ["LAIS MARINHO",1005.65,"2026-04-08"],
  ["LUCIANA DANTAS",375.0,"2026-04-08"],
  ["NIVANIA ADELINO",375.0,"2026-04-08"],
  ["LUCAS - CONVENIOS",4803.74,"2026-04-08"],
  ["VIVIANE COSTA",2693.0,"2026-04-08"],
  ["YASMIN - HIDRO",3715.12,"2026-04-08"],
  ["JAMILLE TEIXEIRA",2252.88,"2026-04-08"],
  ["MILENA MORAIS - CLINICA",1770.44,"2026-04-08"],
  ["MILENA MORAIS - SOCORRO",338.04,"2026-04-08"],
  ["GABRYELLA MAYNARD",3856.37,"2026-04-10"],
  ["JANINY BRAZ - MKT",950.0,"2026-04-20"],
  ["MILENA MORAIS - PARTICULAR",80.0,"2026-04-22"],
  ["OSMALI SILVA",2200.0,"2026-05-07"],
  ["HANAMEEL DIAS",180.0,"2026-05-07"],
  ["LETICIA HELEN",2388.32,"2026-05-07"],
  ["MAYARA SOUZA",1746.97,"2026-05-07"],
  ["ROSILAINE DE JESUS",1746.97,"2026-05-07"],
  ["LUCAS REZENDE - PRO LABORE",2300.0,"2026-05-07"],
  ["JANINY BRAZ - MKT",950.0,"2026-05-07"],
  ["STEFANY DE OLIVEIRA",1679.43,"2026-05-07"],
  ["YASMIN DE OLIVEIRA - HIDRO",4196.56,"2026-05-08"],
  ["YASMIN DE OLIVEIRA - CLINICA",660.06,"2026-05-08"],
  ["VIVIANE COSTA",1123.58,"2026-05-08"],
  ["GRACE KELLY - AQUATICA",3148.0,"2026-05-08"],
  ["APARECIDA - REABILIT",4598.68,"2026-05-08"],
  ["LUCAS REZENDE - CONVENIOS",6315.88,"2026-05-08"],
  ["JULIANA LINHARESS",756.0,"2026-05-08"],
  ["JOSY NASCIMENTO",1047.14,"2026-05-08"],
  ["LAIS MARINHO",1437.75,"2026-05-08"],
  ["BRENNA MIRELLE",372.0,"2026-05-08"],
  ["DANIELA CRUZ - HIDRO",313.04,"2026-05-08"],
  ["NIVANIA ADELINO",175.0,"2026-05-08"],
  ["VINICIUS SOBRAL - PARTICULAR",300.0,"2026-05-11"],
  ["MILENA MORAIS - SOCORRO",800.08,"2026-05-11"],
  ["MILENA MORAIS - CLINICA",1785.52,"2026-05-11"],
  ["GABRYELLA MAYNARD",2297.75,"2026-05-11"],
  ["JULIANNA ARCHIMINO",480.0,"2026-05-11"],
  ["BRUNO NASCIMENTO",2478.14,"2026-05-11"],
  ["JAMILLE GONÇALVES",1335.6,"2026-05-11"],
  ["MARIANA MENDONÇA - CONVENIOS",5000.0,"2026-05-11"],
  ["MILA PIRES - PARTICULAR",406.0,"2026-05-15"],
  ["LUCAS REZENDE - PARTICULAR",1645.0,"2026-05-28"],
  ["MILA PIRES - PARTICULAR",1075.0,"2026-05-28"],
  ["MARIANA MENDONÇA - PARTICULAR",990.93,"2026-06-01"],
  ["ROSILAINE DE JESUS - PASSAGENS",189.0,"2026-06-02"],
  ["ROSILAINE DE JESUS - SALARIO",1566.97,"2026-06-03"],
  ["MAYARA SOUZA - SALARIO",1746.0,"2026-06-03"],
  ["MAYARA SOUZA - ID CURSOS",350.0,"2026-06-03"],
  ["LUCAS REZENDE - PRO LABORE",2300.0,"2026-06-03"],
  ["STEFANY DE OLIVEIRA",1679.43,"2026-06-03"],
  ["LETICIA HELEN",2308.32,"2026-06-03"],
  ["OSMALI SILVA",2200.0,"2026-06-08"],
  ["APARECIDA - REABILIT",2369.34,"2026-06-08"],
  ["BRUNO NASCIMENTO",1192.22,"2026-06-08"],
  ["GABRYELLA MAYNARD",3112.45,"2026-06-08"],
  ["JOSY NASCIMENTO",613.22,"2026-06-08"],
  ["JULIANNA ARCHIMINO",2526.08,"2026-06-08"],
  ["LAIS MARINHO",1974.45,"2026-06-08"],
  ["LUCAS REZENDE - CONVENIOS",5218.56,"2026-06-08"],
  ["MILA PIRES - CONVENIOS",11285.88,"2026-06-08"],
  ["MILENA MORAIS - CLINICA",1010.0,"2026-06-08"],
  ["MILENA MORAIS - SOCORRO",156.0,"2026-06-08"],
  ["OSMALI SILVA - DOMICILIAR",321.9,"2026-06-08"],
  ["VIVIANE COSTA",1510.26,"2026-06-08"],
  ["HANAMEEL DIAS",180.0,"2026-06-08"],
  ["MARIANA MENDONÇA - PARTICULAR",625.0,"2026-06-09"],
  ["MARIANA MENDONÇA - PARTICULAR",729.04,"2026-06-09"],
  ["MARIANA MENDONÇA - RETIRADA",840.0,"2026-06-10"],
  ["APARECIDA - REABILIT (CASSI)",935.52,"2026-06-10"],
  ["JOSY NASCIMENTO (CASSI)",1059.68,"2026-06-10"],
  ["JULIANNA ARCHIMINO (CASSI)",1559.2,"2026-06-10"],
  ["DANIELA CRUZ",375.65,"2026-06-10"],
  ["BRENNA MIRELLE",560.0,"2026-06-10"],
  ["GABRYELLA MAYNARD (CASSI)",779.6,"2026-06-10"],
  ["GRACE - HIDRO",3063.12,"2026-06-10"],
  ["LUCAS REZENDE (CASSI)",3656.64,"2026-06-10"],
  ["BRUNO NASCIMENTO (CASSI)",946.89,"2026-06-10"],
  ["YASMIN DE OLIVEIRA - HIDRO",4630.68,"2026-06-10"],
  ["JULIANA LINHARES  (CASSI)",249.84,"2026-06-10"],
  ["VIVIANE COSTA (CASSI)",244.72,"2026-06-10"],
  ["MILENA MORAIS (CASSI)",935.52,"2026-06-10"],
  ["JAMILLE TEIXEIRA",1351.68,"2026-06-10"],
  ["MARIANA MENDONÇA - PRO LABORE",5000.0,"2026-06-10"],
  ["LUCIANA LOPO",100.0,"2026-06-17"],
  ["VINICIUS SOBRAL",100.0,"2026-06-17"],
  ["MILA PIRES - PARTICULAR",305.0,"2026-06-17"],
  ["VIVO - INTERNET+FIXO",265.22,"2026-06-17"],
  ["DIARISTA (150 EM ESPECIE)",50.0,"2026-06-17"],
  ["MARIANA MENDONÇA - PARTICULAR",3900.0,"2026-06-30"],
  ["LUCAS REZENDE - PRO LABORA",2300.0,"2026-07-03"],
  ["MAYARA SOUZA - SALARIO",1755.97,"2026-07-03"],
  ["MAYARA SOUZA - ID CURSOS",350.0,"2026-07-03"],
  ["ROSILAINE DE JESUS (-200 ADIANTAMENTO SOLICITADO POR FUNCIONARIA)",1555.97,"2026-07-03"],
  ["STEFANY DE OLIVEIRA",1465.62,"2026-07-03"],
  ["LETICIA HELEN",2545.62,"2026-07-06"],
  ["BRENNA MIRELLE",216.0,"2026-07-06"],
  ["APARECIDA - REABILIT",4738.68,"2026-07-06"],
  ["DANIELA - HIDRO",643.8,"2026-07-06"],
  ["JOSY NASCIMENTO",595.72,"2026-07-06"],
  ["JULIANNA ARCHIMINO",810.0,"2026-07-06"],
  ["LUCAS REZENDE - CONVENIOS",5432.96,"2026-07-06"],
  ["MILA PIRES - CONVENIOS",6368.46,"2026-07-06"],
  ["YASMIN SILVA - CLINICA",220.02,"2026-07-06"],
  ["VIVIANE COSTA ",536.5,"2026-07-06"],
  ["MILENA MORAIS - SOCORRO",321.9,"2026-07-06"],
  ["MILENA MORAIS - CLINICA",1240.0,"2026-07-06"],
  ["HANAMEEL DIAS",180.0,"2026-07-06"],
  ["OSMALI SILVA - SALARIO",2200.0,"2026-07-06"],
  ["OSMALI SILVA - DOMICILIAR",321.9,"2026-07-06"],
  ["MILENA MORAIS - SOCORRO (DIFERENÇA)",78.18,"2026-07-06"],
  ["JANINY BRAZ - MKT",712.5,"2026-07-06"],
  ["BRUNO NASCIMENTO",1128.6,"2026-07-09"],
  ["GABRYELLA MAYNARD",1462.5,"2026-07-09"],
  ["LAIS MARINHO",1316.3,"2026-07-09"],
  ["MARIANA MENDONAÇA - PRO LABORE",5000.0,"2026-07-10"],
  ["LUCAS REZENDE - PARTICULAR HIDRO",1000.0,"2026-07-10"],
  ["BRUNO NASCIMENTO - GEAP",1352.37,"2026-07-14"],
  ["APARECIDA - REABILIT (CASSI+GEAP)",2932.91,"2026-07-14"],
  ["BRENNA MIRELLE - GEAP",36.0,"2026-07-14"],
  ["DANIELA - HIDRO (GEAP)",375.65,"2026-07-14"],
  ["JAMILLE TEIXEIRA - GEAP",689.04,"2026-07-14"],
  ["JOSY NASCIMENTO - CASSI+GEAP",1542.37,"2026-07-14"],
  ["LAIS MARINHO - GEAP",975.0,"2026-07-14"],
  ["LUCAS REZENDE - GEAP+CASSI",2404.0,"2026-07-14"],
  ["GRACE - HIDRO",3694.02,"2026-07-14"],
  ["YASMIN SILVA - HIDRO",4209.56,"2026-07-14"],
  ["LUCIANA LOPO",230.0,"2026-07-14"],
  ["MILENA MORAIS - CASSI",935.52,"2026-07-14"],
  ["MILA PIRES - GEAP",780.0,"2026-07-14"],
  ["MILA PIRES - PARTICULAR",250.0,"2026-07-14"],
  ["MARIANA MENDONÇA - PARTICULAR",2800.0,"2026-07-20"],
  ["ROSILAINE DE JESUS RESTANTE DAS FERIAS FUNCIONARIA NÃO GOZARA, SOLICITOU VENDA)",854.67,"2026-07-24"],
  ["ALANNA LETICIA - RESIDUAL (ENCERRAMENTO DE DIVIDA)",2495.01,"2026-07-24"],
  ["LUCAS REZENDE - PARTICULAR",840.0,"2026-07-24"]
];

// ------------------------------------------------------------
//  DADOS REAIS — extraídos do Relatório de NFS-e Emitidas
//  (Prefeitura de Aracaju, competência 01/2026 a 06/2026),
//  agregados por convênio × mês. Fonte: 59 notas fiscais emitidas
//  para 10 convênios distintos.
//  Colunas: [convenio_id, convenio_nome, mes, valor_total, qtd_notas]
// ------------------------------------------------------------
const GUIAS_CONVENIO_HISTORICO_REAL = [
["CONV015","BANCO DO BRASIL","JUNHO",31713.54,2],
  ["CONV016","CODEVASF","JUNHO",737.5,1],
  ["CONV007","PETROBRAS","JUNHO",49003.6,1],
  ["CONV004","CASSIND","JUNHO",1242.6,1],
  ["CONV014","BANESE","JUNHO",949.48,1],
  ["CONV006","BLUE","JUNHO",2610.0,1],
  ["CONV013","ASSOC. HABITAÇÃO E OBRAS PÚBLICAS","JUNHO",473.0,1],
  ["CONV002","AMIL","JUNHO",11840.0,1],
  ["CONV015","BANCO DO BRASIL","MAIO",46890.43,2],
  ["CONV017","FUNASA","MAIO",610.2,1],
  ["CONV007","PETROBRAS","MAIO",51607.1,1],
  ["CONV004","CASSIND","MAIO",207.06,1],
  ["CONV014","BANESE","MAIO",674.6,1],
  ["CONV006","BLUE","MAIO",1430.0,1],
  ["CONV002","AMIL","MAIO",6520.0,1],
  ["CONV015","BANCO DO BRASIL","ABRIL",26691.83,1],
  ["CONV001","GEAP","ABRIL",11314.0,1],
  ["CONV004","CASSIND","ABRIL",2878.61,2],
  ["CONV007","PETROBRAS","ABRIL",59868.2,1],
  ["CONV014","BANESE","ABRIL",963.6,1],
  ["CONV017","FUNASA","ABRIL",210.2,1],
  ["CONV006","BLUE","ABRIL",2030.0,1],
  ["CONV013","ASSOC. HABITAÇÃO E OBRAS PÚBLICAS","ABRIL",236.5,1],
  ["CONV002","AMIL","ABRIL",5440.0,1],
  ["CONV015","BANCO DO BRASIL","MARÇO",34538.44,2],
  ["CONV001","GEAP","MARÇO",19097.43,3],
  ["CONV017","FUNASA","MARÇO",1371.0,1],
  ["CONV004","CASSIND","MARÇO",393.49,2],
  ["CONV007","PETROBRAS","MARÇO",48968.7,1],
  ["CONV014","BANESE","MARÇO",481.8,1],
  ["CONV006","BLUE","MARÇO",400.0,1],
  ["CONV002","AMIL","MARÇO",7320.0,1],
  ["CONV015","BANCO DO BRASIL","FEVEREIRO",26504.78,2],
  ["CONV001","GEAP","FEVEREIRO",11006.99,1],
  ["CONV004","CASSIND","FEVEREIRO",1242.6,1],
  ["CONV017","FUNASA","FEVEREIRO",1934.39,1],
  ["CONV016","CODEVASF","FEVEREIRO",218.3,1],
  ["CONV007","PETROBRAS","FEVEREIRO",57032.6,1],
  ["CONV014","BANESE","FEVEREIRO",481.8,1],
  ["CONV013","ASSOC. HABITAÇÃO E OBRAS PÚBLICAS","FEVEREIRO",236.5,1],
  ["CONV002","AMIL","FEVEREIRO",4240.0,1],
  ["CONV015","BANCO DO BRASIL","JANEIRO",37436.2,2],
  ["CONV007","PETROBRAS","JANEIRO",60386.0,1],
  ["CONV001","GEAP","JANEIRO",14713.2,1],
  ["CONV004","CASSIND","JANEIRO",621.3,1],
  ["CONV014","BANESE","JANEIRO",674.6,1],
  ["CONV017","FUNASA","JANEIRO",850.2,1],
  ["CONV006","BLUE","JANEIRO",210.0,1],
  ["CONV013","ASSOC. HABITAÇÃO E OBRAS PÚBLICAS","JANEIRO",1655.4,1],
  ["CONV002","AMIL","JANEIRO",3780.0,1]
];

// ------------------------------------------------------------
//  FUNÇÃO PRINCIPAL — idempotente
// ------------------------------------------------------------
function importarHistoricoReal(usuario) {
  try {
    _verificarUsuario(usuario, ['admin','gestor']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultado = {
      particulares: {importados:0, pulados:0, profissionalNaoEncontrado:[]},
      despesasOperacionais: {importados:0, pulados:0},
      repasses: {importados:0, pulados:0},
      convenios: {cadastrados:0, jaExistiam:0},
      guiasConvenio: {importadas:0, puladas:0},
      recebimentosGerados: {gerados:0, pulados:0}
    };

    // -------- 0) CADASTRA CONVÊNIOS QUE FALTAVAM --------
    const shConv = ss.getSheetByName(CONFIG.SHEETS.CONVENIOS);
    const convExistentes = _getSheet(ss, CONFIG.SHEETS.CONVENIOS);
    const idsConvExistentes = new Set(convExistentes.map(r => String(r[0])));
    const nomesConvExistentes = new Set(convExistentes.map(r => String(r[1]).toUpperCase()));
    CONVENIOS_FALTANTES.forEach(row => {
      const [id, nome, prazo] = row;
      if (idsConvExistentes.has(id) || nomesConvExistentes.has(nome.toUpperCase())) {
        resultado.convenios.jaExistiam++;
        return;
      }
      shConv.appendRow([id, nome, prazo, '', 'SIM', '', 'Cadastrado automaticamente pela importação de histórico — confirme o prazo de pagamento real.']);
      resultado.convenios.cadastrados++;
    });

    // -------- 1) PARTICULARES --------
    const shPart = ss.getSheetByName(CONFIG.SHEETS.PARTICULARES);
    const listasProf = _getSheet(ss, CONFIG.SHEETS.PROFISSIONAIS);
    const profPorNome = {};
    listasProf.forEach(r => { profPorNome[String(r[1]).toUpperCase()] = { id: r[0], nome: r[1] }; });

    const listasServ = _getSheet(ss, CONFIG.SHEETS.SERVICOS);
    const servPorNome = {};
    listasServ.forEach(r => { servPorNome[String(r[1]).toUpperCase()] = { id: r[0], nome: r[1] }; });

    const partExistentes = _getSheet(ss, CONFIG.SHEETS.PARTICULARES);
    const idsPartExistentes = new Set(partExistentes.map(r => String(r[0])));

    PARTICULARES_HISTORICO_REAL.forEach(row => {
      const [paciente, data, especialidade, profissionalTxt, valor, formaPgto, quantidade] = row;
      const idDeterministico = 'HIST-PART-' + Utilities.base64EncodeWebSafe(
        Utilities.newBlob(paciente+'|'+data+'|'+valor+'|'+quantidade).getBytes()
      ).slice(0,16);
      if (idsPartExistentes.has(idDeterministico)) { resultado.particulares.pulados++; return; }

      const nomeNormalizado = _normalizarNomeProfissional(profissionalTxt);
      const profInfo = profPorNome[nomeNormalizado];
      if (!profInfo) resultado.particulares.profissionalNaoEncontrado.push(profissionalTxt);

      const servInfo = servPorNome[String(especialidade).toUpperCase()] || null;
      const mes = _mesDoDate(data);

      shPart.appendRow([
        idDeterministico, data, mes, '', paciente,
        profInfo ? profInfo.id : '', profInfo ? profInfo.nome : (nomeNormalizado + ' (CONFERIR — não encontrado no cadastro)'),
        servInfo ? servInfo.id : '', servInfo ? servInfo.nome : especialidade,
        valor, formaPgto, 1, quantidade,
        'Importado do histórico real (planilha PAGAMENTOS_PARTICULARES). ' + (profInfo ? '' : 'ATENÇÃO: profissional "' + profissionalTxt + '" não encontrado no cadastro — associe manualmente.'),
        'Confirmado', 'IMPORTAÇÃO HISTÓRICO', new Date(), ''
      ]);
      resultado.particulares.importados++;
    });

    // -------- 2) DESPESAS OPERACIONAIS --------
    const shDesp = ss.getSheetByName(CONFIG.SHEETS.DESPESAS);
    const despExistentes = _getSheet(ss, CONFIG.SHEETS.DESPESAS);
    const idsDespExistentes = new Set(despExistentes.map(r => String(r[0])));

    DESPESAS_OPERACIONAIS_HISTORICO_REAL.forEach(row => {
      const [descricao, valor, data] = row;
      const idDeterministico = 'HIST-DESP-' + Utilities.base64EncodeWebSafe(
        Utilities.newBlob(descricao+'|'+data+'|'+valor).getBytes()
      ).slice(0,16);
      if (idsDespExistentes.has(idDeterministico)) { resultado.despesasOperacionais.pulados++; return; }
      const mes = _mesDoDate(data);
      shDesp.appendRow([
        idDeterministico, data, mes, 'Histórico Importado', descricao, '', valor, '', 'FIXO', 'Pago',
        data, data, '', 'Importado do histórico real (planilha DESPESAS, seção operacional).',
        'IMPORTAÇÃO HISTÓRICO', new Date(), ''
      ]);
      resultado.despesasOperacionais.importados++;
    });

    // -------- 3) REPASSES PROFISSIONAIS (histórico) --------
    REPASSES_PROFISSIONAIS_HISTORICO_REAL.forEach(row => {
      const [descricao, valor, data] = row;
      const idDeterministico = 'HIST-REPASSE-' + Utilities.base64EncodeWebSafe(
        Utilities.newBlob(descricao+'|'+data+'|'+valor).getBytes()
      ).slice(0,16);
      if (idsDespExistentes.has(idDeterministico)) { resultado.repasses.pulados++; return; }
      const mes = _mesDoDate(data);
      shDesp.appendRow([
        idDeterministico, data, mes, 'Repasse Profissional (Histórico)', descricao, '', valor, '', 'FIXO', 'Pago',
        data, data, '', 'Importado do histórico real (planilha DESPESAS, seção PROFISSIONAIS). Use para validar se calcularComissaoGuia/calcularComissaoParticular batem com o valor que a gestora calculava manualmente neste período.',
        'IMPORTAÇÃO HISTÓRICO', new Date(), ''
      ]);
      resultado.repasses.importados++;
    });

    // -------- 4) GUIAS DE CONVÊNIO (agregadas por mês, do relatório de NFS-e) --------
    const shGuias = ss.getSheetByName(CONFIG.SHEETS.GUIAS);
    const guiasExistentes = _getSheet(ss, CONFIG.SHEETS.GUIAS);
    const idsGuiasExistentes = new Set(guiasExistentes.map(r => String(r[0])));
    const convAtualizados = _getSheet(ss, CONFIG.SHEETS.CONVENIOS);
    const convPorId = {};
    convAtualizados.forEach(r => { convPorId[String(r[0])] = { id:r[0], nome:r[1], prazo:r[2] }; });

    GUIAS_CONVENIO_HISTORICO_REAL.forEach(row => {
      const [convenioId, convenioNome, mesTexto, valorTotal, qtdNotas] = row;
      const idDeterministico = 'HIST-GUIA-' + Utilities.base64EncodeWebSafe(
        Utilities.newBlob(convenioId+'|'+mesTexto+'|'+valorTotal).getBytes()
      ).slice(0,16);
      if (idsGuiasExistentes.has(idDeterministico)) { resultado.guiasConvenio.puladas++; return; }

      const convInfo = convPorId[convenioId] || { nome: convenioNome, prazo: 60 };
      const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
      const mesIdx = MESES.indexOf(mesTexto);
      const ano = 2026;
      const dataRef = new Date(ano, mesIdx, 28);
      const dataStr = _dateStr(dataRef);
      const prazoDias = _sanNum(convInfo.prazo, 60);
      const dataPrevPgto = new Date(ano, mesIdx, 28 + prazoDias);

      shGuias.appendRow([
        idDeterministico, dataStr, mesTexto, convenioId, convInfo.nome,
        '', 'Guia consolidada mensal (' + qtdNotas + ' nota(s) fiscal(is))',
        '', '', '', '', '',
        valorTotal, prazoDias, dataStr, _dateStr(dataPrevPgto), dataStr, 'Pago', 0,
        'Importado do relatório de NFS-e Emitidas (Prefeitura de Aracaju, competência ' + mesTexto + '/2026). Valor consolidado de ' + qtdNotas + ' nota(s) fiscal(is) — não há detalhamento por paciente neste relatório. Marcado como "Pago" porque nota fiscal emitida já reflete faturamento reconhecido; ajuste o status se a gestora souber que algum valor ainda estava pendente de recebimento nesse período.',
        'IMPORTAÇÃO HISTÓRICO', new Date(), ''
      ]);
      resultado.guiasConvenio.importadas++;
    });

    // -------- 5) GERA RECEBIMENTOS MÊS A MÊS (para o Fluxo de Caixa Histórico) --------
    const shReceb = ss.getSheetByName(CONFIG.SHEETS.RECEBIMENTOS);
    const recebExistentes = _getSheet(ss, CONFIG.SHEETS.RECEBIMENTOS);
    const idsRecebExistentes = new Set(recebExistentes.map(r => String(r[0])));
    const MESES_ORDEM = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];

    MESES_ORDEM.slice(0,7).forEach(mesTx => {
      const totalParticularMes = getParticulares({mes:mesTx}).reduce((s,p)=>s+_sanNum(p.valor),0);
      if (totalParticularMes > 0) {
        const idP = 'HIST-RECEB-PART-' + mesTx;
        if (!idsRecebExistentes.has(idP)) {
          shReceb.appendRow([idP, _dateStr(new Date(2026, MESES_ORDEM.indexOf(mesTx), 28)), 'Particular (Consolidado Histórico)', '', '', '', totalParticularMes, '', mesTx, 'Soma consolidada dos lançamentos particulares importados do histórico real para este mês.', new Date()]);
          resultado.recebimentosGerados.gerados++;
        } else resultado.recebimentosGerados.pulados++;
      }
      const totalConvenioMes = getGuias({mes:mesTx}).reduce((s,g)=>s+_sanNum(g.valor_total)-_sanNum(g.valor_glosado),0);
      if (totalConvenioMes > 0) {
        const idC = 'HIST-RECEB-CONV-' + mesTx;
        if (!idsRecebExistentes.has(idC)) {
          shReceb.appendRow([idC, _dateStr(new Date(2026, MESES_ORDEM.indexOf(mesTx), 28)), 'Convênio (Consolidado Histórico)', '', '', '', totalConvenioMes, '', mesTx, 'Soma consolidada das guias de convênio importadas do relatório de NFS-e para este mês.', new Date()]);
          resultado.recebimentosGerados.gerados++;
        } else resultado.recebimentosGerados.pulados++;
      }
    });

    _log(usuario.nome, 'IMPORTAR_HISTORICO_REAL',
      `Particulares: ${resultado.particulares.importados} importados/${resultado.particulares.pulados} já existiam | ` +
      `Despesas: ${resultado.despesasOperacionais.importados} importadas/${resultado.despesasOperacionais.pulados} já existiam | ` +
      `Repasses: ${resultado.repasses.importados} importados/${resultado.repasses.pulados} já existiam | ` +
      `Convênios cadastrados: ${resultado.convenios.cadastrados} | ` +
      `Guias de convênio: ${resultado.guiasConvenio.importadas} importadas/${resultado.guiasConvenio.puladas} já existiam | ` +
      `Recebimentos consolidados: ${resultado.recebimentosGerados.gerados} gerados/${resultado.recebimentosGerados.pulados} já existiam`);

    return { ok:true, resultado };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}
