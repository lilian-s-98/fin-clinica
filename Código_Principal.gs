// ============================================================
//  INSTITUTO DA DOR — Codigo_Principal.gs  v6.0
//  Sistema de Gestão Clínica e Financeira
//  Gabillaud Consultoria | 2026
//
//  NOVIDADES v6.0:
//  - Removido: Assistente de IA (consultarIA e toda a integração
//    com a API da Anthropic). Não há mais chave de API armazenada.
//  - Novo módulo de Comissionamento configurável por profissional
//    (Modulo_Comissionamento.gs)
//  - Novo módulo de Regras de Convênio / Faturamento configurável
//    (Modulo_Convenios_Faturamento.gs)
//  - Novo módulo de Glosas (Modulo_Glosas.gs)
//  - Novo módulo de Lotes/Protocolos (Modulo_Lotes.gs)
//  - Dashboard e KPIs ampliados (Modulo_Dashboard_KPI.gs)
//  - Mantido 100% do que já existia: checklist por perfil,
//    histórico, notificações, guias, particulares, despesas,
//    agenda, autenticação, log de auditoria.
//
//  ATENÇÃO — LEIA ANTES DE RODAR QUALQUER FUNÇÃO:
//  setupInicial() CONTINUA existindo apenas para uma planilha NOVA
//  e vazia (ele limpa o conteúdo das abas). Se sua planilha já tem
//  dados de produção, NÃO rode setupInicial(). Rode migrarV6()
//  em vez disso — ela só ADICIONA as abas/colunas novas, sem
//  apagar nada do que já existe.
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
  META_ANUAL: 1200000,
  VERSION: '6.0',
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
    DESPESAS_RECORRENTES: 'Despesas_Recorrentes'
  }
};

// ============================================================
//  PONTO DE ENTRADA WEB
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
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

function _popularDadosIniciais(ss) {
  const shU = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  shU.getRange(2,1,4,10).setValues([
    ['USR001','Admin Master','admin@institutodador.com.br','admin','SIM',_hashSenha('admin2026'),new Date(),'','#0049AF','Administrador'],
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

  const shSv = ss.getSheetByName(CONFIG.SHEETS.SERVICOS);
  shSv.getRange(2,1,8,6).setValues([
    ['SRV001','FISIOTERAPIA MOTORA','Fisioterapia',130,'SIM',50],
    ['SRV002','FISIOTERAPIA PÉLVICA','Fisioterapia',200,'SIM',50],
    ['SRV003','RPG','Fisioterapia',180,'SIM',60],
    ['SRV004','AVALIAÇÃO INICIAL','Avaliação',180,'SIM',60],
    ['SRV005','ACUPUNTURA','Terapia Complementar',180,'SIM',50],
    ['SRV006','HOME CARE - FISIOTERAPIA','Home Care',180,'SIM',60],
    ['SRV007','QUIROPRAXIA','Terapia Manual',220,'SIM',50],
    ['SRV008','DRENAGEM LINFÁTICA','Estética',180,'SIM',60],
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
    const rl = _checarRateLimit(email);
    if (rl.bloqueado) return { ok: false, msg: rl.msg };
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USUARIOS);
    const dados = sh.getDataRange().getValues();
    const hash = _hashSenha(senha);
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (row[2].toLowerCase() === email && row[4] === 'SIM') {
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
// ============================================================
const _san = v => v ? String(v).replace(/[<>'"&]/g,'').trim() : '';
const _sanNum = (v,d) => { const n = parseFloat(v); return isNaN(n) ? (d||0) : n; };
const _mesDoDate = d => ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'][d.getMonth()];
const _dateStr = d => { if (!d) return ''; if (typeof d === 'string') return d.slice(0,10); return Utilities.formatDate(d,'America/Sao_Paulo','yyyy-MM-dd'); };

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
