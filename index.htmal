/**
 * PORTAL DE PEDIDOS — BACKEND (Google Apps Script)
 * ---------------------------------------------------------------
 * Como instalar:
 * 1. Abra a planilha Google Sheets com as abas de clientes.
 * 2. Menu Extensões → Apps Script.
 * 3. Apague o conteúdo do Code.gs e cole este arquivo inteiro.
 * 4. Adicione a aba USUARIOS (veja instruções que te mandei no chat).
 * 5. Menu Implantar → Nova implantação → tipo "App da Web".
 *      - Executar como: Eu (sua conta)
 *      - Quem pode acessar: Qualquer pessoa
 * 6. Copie a URL gerada e cole no PORTAL-VENDEDOR.html (variável API_URL).
 * ---------------------------------------------------------------
 */

// COLE AQUI O ID DA SUA PLANILHA (está na URL, entre /d/ e /edit).
// Assim o script funciona mesmo se não estiver "vinculado" à planilha.
const SPREADSHEET_ID = '19L0rpTYIKhXyKOJoqRsQlRE7Yn-l7EwAVF0a1uvS7SQ';

// Abas que contêm carteiras de cliente por vendedor.
// Para adicionar um novo vendedor no futuro: crie a aba com o mesmo
// formato de colunas e acrescente o nome dela aqui.
const VENDOR_SHEETS = ['VENDEDOR ALINA', 'VENDEDOR ANA LAMY', 'VENDEDOR FELIX', 'VENDEDOR MARIA SALETE'];

const VERSAO_CODIGO = 'v3-vendedor-sheets-2026-08-27';

const SHEET_USUARIOS = 'USUARIOS';
const SHEET_PEDIDOS = 'PEDIDOS';
const SHEET_ITENS = 'ITENS_PEDIDO';
const SHEET_GRI = 'TABELA GRI';
const SHEET_CPM = 'TABELA CPM';
const SHEET_CADASTROS = 'CADASTROS_CLIENTES';
// E-mail padrão para onde os cadastros de cliente novo são enviados.
// Pode ser sobrescrito pelo vendedor no próprio formulário do portal.
const EMAIL_PADRAO_CADASTRO = 'comercial@biowellamerica.com.br';

// Mapeia os vários nomes de coluna encontrados nas abas de vendedor
// para um nome de campo único (canônico).
const CLIENT_HEADER_MAP = {
  'CODIGO SAP': 'sap',
  'CÓDIGO DO PN': 'sap',
  'CODIGO DO PN': 'sap',
  'NOME DO PN': 'nome',
  'FANTASIA': 'fantasia',
  'CODIGO LOJA': 'codigoLoja',
  'CNPJ / CPF': 'cnpj',
  'UF': 'uf',
  'TABELA': 'tabela',
  'VENDENDOR': 'vendedor',
  'CÓDIGO DO VENDEDOR': 'vendedor',
  'CODIGO DO VENDEDOR': 'vendedor',
  'CÓDIGO DA CONDIÇÃO DE PAGAMENTO': 'condPagamento',
  'CODIGO DA CONDICAO DE PAGAMENTO': 'condPagamento',
  'CIDADE': 'cidade',
  'BAIRRO': 'bairro',
  'COMPLEMENTO': 'complemento',
  'ATIVO': 'ativo'
};

const PRODUCT_HEADER_MAP = {
  'LINHA': 'linha',
  'CÓDIGO': 'codigo',
  'CODIGO': 'codigo',
  'CAIXARIA': 'caixaria',
  'CÓDIGO DE BARRAS EAN': 'ean',
  'CÓDIGO DE BARRAS': 'ean',
  'CODIGO DE BARRAS': 'ean',
  'VEGANO': 'vegano',
  'DESCRIÇÃO DO PRODUTO': 'descricao',
  'DESCRICAO DO PRODUTO': 'descricao',
  'CUSTO GRI UND': 'custo',
  'CUSTO CPM': 'custo',
  'STATUS ESTOQUE': 'status'
};

/* ============================= ROTEAMENTO ============================= */

function doPost(e) {
  return handleRequest(e);
}
function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  Logger.log('Requisição recebida. e.parameter: ' + JSON.stringify(e && e.parameter) + ' | e.postData: ' + (e && e.postData ? e.postData.contents : 'nenhum'));
  let body = {};
  try {
    if (e.parameter && e.parameter.data) {
      body = JSON.parse(e.parameter.data);
    } else if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
    }
  } catch (err) {
    return jsonOut({ ok: false, error: 'Requisição inválida.' });
  }

  const action = body.action;
  let result;
  try {
    switch (action) {
      case 'login':
        result = actionLogin(body);
        break;
      case 'getClientes':
        result = actionGetClientes(body);
        break;
      case 'getProdutos':
        result = actionGetProdutos(body);
        break;
      case 'criarPedido':
        result = actionCriarPedido(body);
        break;
      case 'getPedidos':
        result = actionGetPedidos(body);
        break;
      case 'getVendedores':
        result = actionGetVendedores(body);
        break;
      case 'debugHeaders':
        result = actionDebugHeaders(body);
        break;
      case 'cadastrarClienteEmail':
        result = actionCadastrarClienteEmail(body);
        break;
      default:
        result = { ok: false, error: 'Ação desconhecida.' };
    }
  } catch (err) {
    result = { ok: false, error: 'Erro no servidor: ' + err.message };
  }
  return jsonOut(result);
}

function jsonOut(obj) {
  obj.versao = VERSAO_CODIGO;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================= AUTENTICAÇÃO ============================= */

function isAdminPerfil(perfil) {
  return String(perfil || '').trim().toLowerCase().indexOf('admin') === 0;
}

function findUser(login, senha) {
  const sheet = getSheet(SHEET_USUARIOS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim().toLowerCase());
  const idxNome = headers.indexOf('nome');
  const idxLogin = headers.indexOf('login');
  const idxSenha = headers.indexOf('senha');
  const idxPerfil = headers.indexOf('perfil');
  const idxCarteira = headers.indexOf('carteiravendedor');
  const idxStatus = headers.indexOf('status');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rowLogin = String(row[idxLogin] || '').trim().toLowerCase();
    const rowSenha = String(row[idxSenha] || '');
    if (rowLogin === String(login || '').trim().toLowerCase() && rowSenha === String(senha || '')) {
      const status = idxStatus >= 0 ? String(row[idxStatus] || '').trim().toLowerCase() : 'ativo';
      if (status && status !== 'ativo' && status !== 'sim') {
        return { blocked: true };
      }
      return {
        nome: row[idxNome],
        login: row[idxLogin],
        perfil: String(row[idxPerfil] || 'vendedor').trim().toLowerCase(),
        carteira: idxCarteira >= 0 ? String(row[idxCarteira] || '').trim() : ''
      };
    }
  }
  return null;
}

function requireAuth(body) {
  const user = findUser(body.login, body.senha);
  if (!user || user.blocked) return null;
  return user;
}

function actionLogin(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Login ou senha inválidos, ou usuário desativado.' };
  return { ok: true, nome: user.nome, perfil: user.perfil, carteira: user.carteira };
}

/* ============================= CLIENTES ============================= */

function actionGetClientes(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const todos = getAllClientes();
  if (isAdminPerfil(user.perfil)) {
    return { ok: true, clientes: todos };
  }
  const carteira = normalizeVendorName(user.carteira);
  const filtrados = todos.filter(c => normalizeVendorName(c.vendedor) === carteira);
  return { ok: true, clientes: filtrados };
}

function getAllClientes() {
  const out = [];
  VENDOR_SHEETS.forEach(sheetName => {
    const sheet = getSheet(sheetName, true);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return;
    const headers = rows[0].map(h => String(h).trim().toUpperCase());
    const fieldIdx = {};
    headers.forEach((h, i) => {
      const field = CLIENT_HEADER_MAP[h];
      if (field) fieldIdx[field] = i;
    });
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[fieldIdx.sap] && !row[fieldIdx.nome]) continue;
      const ativo = fieldIdx.ativo >= 0 ? String(row[fieldIdx.ativo] || '').trim().toLowerCase() : 'sim';
      if (ativo && ativo !== 'sim' && ativo !== 'ativo') continue;
      out.push({
        sap: String(row[fieldIdx.sap] || ''),
        nome: String(row[fieldIdx.nome] || ''),
        fantasia: fieldIdx.fantasia >= 0 ? String(row[fieldIdx.fantasia] || '') : '',
        codigoLoja: fieldIdx.codigoLoja >= 0 ? String(row[fieldIdx.codigoLoja] || '') : '',
        cnpj: fieldIdx.cnpj >= 0 ? String(row[fieldIdx.cnpj] || '') : '',
        tabela: normalizeTabela(fieldIdx.tabela >= 0 ? row[fieldIdx.tabela] : ''),
        vendedor: fieldIdx.vendedor >= 0 ? String(row[fieldIdx.vendedor] || '') : sheetName.replace(/^VENDEDOR\s+/i, ''),
        condPagamento: fieldIdx.condPagamento >= 0 ? String(row[fieldIdx.condPagamento] || '') : '',
        cidade: fieldIdx.cidade >= 0 ? String(row[fieldIdx.cidade] || '') : '',
        bairro: fieldIdx.bairro >= 0 ? String(row[fieldIdx.bairro] || '') : '',
        complemento: fieldIdx.complemento >= 0 ? String(row[fieldIdx.complemento] || '') : ''
      });
    }
  });
  return out;
}

function normalizeTabela(v) {
  const s = String(v || '').toUpperCase();
  if (s.indexOf('CPM') >= 0) return 'CPM';
  if (s.indexOf('GRI') >= 0) return 'GRI';
  return 'GRI'; // padrão caso não informado
}

function normalizeVendorName(v) {
  return String(v || '').trim().toUpperCase();
}

/* ============================= PRODUTOS ============================= */

function actionGetProdutos(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const tabela = String(body.tabela || 'GRI').toUpperCase();
  const sheetName = tabela === 'CPM' ? SHEET_CPM : SHEET_GRI;
  const sheet = getSheet(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim().toUpperCase());
  const fieldIdx = {};
  headers.forEach((h, i) => {
    const field = PRODUCT_HEADER_MAP[h];
    if (field) fieldIdx[field] = i;
  });

  const produtos = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[fieldIdx.codigo]) continue;
    produtos.push({
      codigo: row[fieldIdx.codigo],
      ean: fieldIdx.ean >= 0 ? String(row[fieldIdx.ean] || '') : '',
      descricao: String(row[fieldIdx.descricao] || '').trim(),
      caixaria: fieldIdx.caixaria >= 0 ? row[fieldIdx.caixaria] : '',
      vegano: fieldIdx.vegano >= 0 ? String(row[fieldIdx.vegano] || '') : '',
      custo: Number(row[fieldIdx.custo]) || 0,
      status: fieldIdx.status >= 0 ? String(row[fieldIdx.status] || 'OK') : 'OK'
    });
  }
  return { ok: true, tabela: tabela, produtos: produtos };
}

/* ============================= PEDIDOS ============================= */

function actionCriarPedido(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const cliente = body.cliente || {};
  const itens = body.itens || [];
  if (!cliente.sap && !cliente.nome) return { ok: false, error: 'Cliente não informado.' };
  if (!itens.length) return { ok: false, error: 'Pedido sem itens.' };

  const pedidoId = gerarPedidoId();
  const agora = new Date();

  let total = 0;
  itens.forEach(it => { total += (Number(it.preco) || 0) * (Number(it.quantidade) || 0); });

  const sheetPedidos = getSheet(SHEET_PEDIDOS, false, [
    'PedidoID', 'DataHora', 'Vendedor', 'ClienteSAP', 'ClienteNome', 'ClienteFantasia',
    'CNPJ', 'Tabela', 'CondPagamento', 'Cidade', 'ValorTotal', 'Observacao', 'Status'
  ]);
  sheetPedidos.appendRow([
    pedidoId, agora, user.nome, cliente.sap || '', cliente.nome || '', cliente.fantasia || '',
    cliente.cnpj || '', cliente.tabela || '', cliente.condPagamento || '', cliente.cidade || '',
    total, body.observacao || '', 'Enviado'
  ]);

  const sheetItens = getSheet(SHEET_ITENS, false, [
    'PedidoID', 'Codigo', 'Descricao', 'Quantidade', 'PrecoUnitario', 'Subtotal', 'Vendedor', 'ClienteNome'
  ]);
  itens.forEach(it => {
    const qtd = Number(it.quantidade) || 0;
    const preco = Number(it.preco) || 0;
    sheetItens.appendRow([pedidoId, it.codigo, it.descricao, qtd, preco, qtd * preco, user.nome, cliente.nome || cliente.fantasia || '']);
  });

  return { ok: true, pedidoId: pedidoId, total: total, data: agora.toISOString() };
}

function gerarPedidoId() {
  const sheet = getSheet(SHEET_PEDIDOS, false, [
    'PedidoID', 'DataHora', 'Vendedor', 'ClienteSAP', 'ClienteNome', 'ClienteFantasia',
    'CNPJ', 'Tabela', 'CondPagamento', 'Cidade', 'ValorTotal', 'Observacao', 'Status'
  ]);
  const ano = new Date().getFullYear();
  const lastRow = sheet.getLastRow();
  const seq = lastRow < 1 ? 1 : lastRow; // cabeçalho ocupa a linha 1
  const numero = String(seq).padStart(4, '0');
  return 'PC' + ano + numero;
}

function actionGetPedidos(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const sheetPedidos = getSheet(SHEET_PEDIDOS, true);
  const sheetItens = getSheet(SHEET_ITENS, true);
  if (!sheetPedidos) return { ok: true, pedidos: [] };

  const rows = sheetPedidos.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });

  const itensPorPedido = {};
  if (sheetItens) {
    const irows = sheetItens.getDataRange().getValues();
    if (irows.length > 1) {
      const ih = irows[0].map(h => String(h).trim());
      const iidx = {};
      ih.forEach((h, i) => { iidx[h] = i; });
      for (let i = 1; i < irows.length; i++) {
        const r = irows[i];
        const pid = r[iidx['PedidoID']];
        if (!itensPorPedido[pid]) itensPorPedido[pid] = [];
        itensPorPedido[pid].push({
          codigo: r[iidx['Codigo']],
          descricao: r[iidx['Descricao']],
          quantidade: r[iidx['Quantidade']],
          preco: r[iidx['PrecoUnitario']],
          subtotal: r[iidx['Subtotal']]
        });
      }
    }
  }

  const pedidos = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const vendedorPedido = String(r[idx['Vendedor']] || '');
    if (!isAdminPerfil(user.perfil) && vendedorPedido.trim().toUpperCase() !== String(user.nome).trim().toUpperCase()) {
      continue;
    }
    const pid = r[idx['PedidoID']];
    pedidos.push({
      pedidoId: pid,
      dataHora: r[idx['DataHora']],
      vendedor: vendedorPedido,
      clienteSap: r[idx['ClienteSAP']],
      clienteNome: r[idx['ClienteNome']],
      clienteFantasia: r[idx['ClienteFantasia']],
      cnpj: r[idx['CNPJ']],
      tabela: r[idx['Tabela']],
      condPagamento: r[idx['CondPagamento']],
      cidade: r[idx['Cidade']],
      valorTotal: r[idx['ValorTotal']],
      observacao: r[idx['Observacao']],
      status: r[idx['Status']],
      itens: itensPorPedido[pid] || []
    });
  }
  pedidos.reverse(); // mais recentes primeiro
  return { ok: true, pedidos: pedidos };
}

/* ============================= ADMIN ============================= */

function actionGetVendedores(body) {
  const user = requireAuth(body);
  if (!user || !isAdminPerfil(user.perfil)) return { ok: false, error: 'Acesso restrito ao administrador.' };

  const sheet = getSheet(SHEET_USUARIOS);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => String(h).trim().toLowerCase());
  const idxNome = headers.indexOf('nome');
  const idxPerfil = headers.indexOf('perfil');
  const idxCarteira = headers.indexOf('carteiravendedor');
  const idxStatus = headers.indexOf('status');

  const vendedores = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    vendedores.push({
      nome: row[idxNome],
      perfil: row[idxPerfil],
      carteira: idxCarteira >= 0 ? row[idxCarteira] : '',
      status: idxStatus >= 0 ? row[idxStatus] : 'Ativo'
    });
  }
  return { ok: true, vendedores: vendedores };
}

/* ============================= CADASTRO DE CLIENTE NOVO ============================= */

function actionCadastrarClienteEmail(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const d = body.dados || {};
  if (!d.razaoSocial) return { ok: false, error: 'Preencha ao menos a Razão Social.' };

  const destino = (body.emailDestino || EMAIL_PADRAO_CADASTRO || '').trim();
  if (!destino) return { ok: false, error: 'Informe o e-mail de destino do cadastro.' };

  // Monta o corpo do e-mail com todos os campos.
  const campos = [
    ['Razão social', d.razaoSocial],
    ['Nome fantasia', d.fantasia],
    ['CNPJ', d.cnpj],
    ['Inscrição estadual', d.inscricaoEstadual],
    ['Endereço completo com CEP', d.endereco],
    ['Complemento', d.complemento],
    ['E-mail para XML', d.emailXml],
    ['E-mail Boleto Bancário', d.emailBoleto],
    ['Telefone Comercial', d.telComercial],
    ['Telefone Financeiro', d.telFinanceiro],
    ['Condição de Pagamento Negociado', d.condPagamento],
    ['Tabela de Preço', d.tabelaPreco]
  ];
  let corpo = 'Novo cadastro de cliente solicitado pelo portal de pedidos.\n\n';
  corpo += 'Vendedor: ' + user.nome + '\n';
  corpo += 'Data: ' + new Date().toLocaleString('pt-BR') + '\n\n';
  campos.forEach(([label, valor]) => {
    corpo += label + ': ' + (valor || '-') + '\n';
  });
  if (d.observacao) corpo += '\nObservações: ' + d.observacao + '\n';

  // Prepara os anexos (recebidos em base64 do navegador).
  const anexos = body.anexos || [];
  const blobs = [];
  anexos.forEach(a => {
    try {
      const bytes = Utilities.base64Decode(a.data);
      blobs.push(Utilities.newBlob(bytes, a.mimeType || 'application/octet-stream', a.nome || 'anexo'));
    } catch (err) {
      // ignora anexo corrompido, mas segue com o resto
    }
  });

  const assunto = 'CADASTRO CLIENTE NOVO';
  const opcoes = {};
  if (blobs.length) opcoes.attachments = blobs;

  MailApp.sendEmail(destino, assunto, corpo, opcoes);

  // Registra a solicitação na planilha para histórico/rastreabilidade.
  try {
    const sheet = getSheet(SHEET_CADASTROS, false, [
      'DataHora', 'Vendedor', 'RazaoSocial', 'Fantasia', 'CNPJ', 'InscricaoEstadual',
      'Endereco', 'Complemento', 'EmailXml', 'EmailBoleto', 'TelComercial', 'TelFinanceiro',
      'CondPagamento', 'TabelaPreco', 'Observacao', 'QtdAnexos', 'EmailDestino'
    ]);
    sheet.appendRow([
      new Date(), user.nome, d.razaoSocial, d.fantasia || '', d.cnpj || '', d.inscricaoEstadual || '',
      d.endereco || '', d.complemento || '', d.emailXml || '', d.emailBoleto || '', d.telComercial || '',
      d.telFinanceiro || '', d.condPagamento || '', d.tabelaPreco || '', d.observacao || '', blobs.length, destino
    ]);
  } catch (err) {
    // não bloqueia o envio do e-mail se o registro na planilha falhar
  }

  return { ok: true };
}

/* ============================= DIAGNÓSTICO ============================= */

function actionDebugHeaders(body) {
  const user = requireAuth(body);
  if (!user) return { ok: false, error: 'Sessão inválida. Faça login novamente.' };

  const ss = getSpreadsheet();
  const todasAsAbas = ss.getSheets().map(s => s.getName());

  const out = {};
  VENDOR_SHEETS.forEach(sheetName => {
    const sheet = getSheet(sheetName, true);
    if (!sheet) { out[sheetName] = 'ABA NÃO ENCONTRADA'; return; }
    const rows = sheet.getDataRange().getValues();
    out[sheetName] = {
      linhas: rows.length,
      cabecalho: rows[0],
      exemploLinha2: rows.length > 1 ? rows[1] : null
    };
  });
  return { ok: true, nomePlanilha: ss.getName(), todasAsAbas: todasAsAbas, abas: out };
}

/* ============================= TESTE MANUAL (rode direto no editor) ============================= */

function testarDiagnostico() {
  const resultado = actionDebugHeaders({ login: 'adm', senha: 'adm123' });
  Logger.log(JSON.stringify(resultado, null, 2));
}

/* ============================= UTIL ============================= */

function getSpreadsheet() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('COLE_AQUI') === 0) {
    throw new Error('Configure a constante SPREADSHEET_ID no topo do script com o ID da sua planilha.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name, silent, headerRowIfCreate) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    if (silent) return null;
    if (headerRowIfCreate) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headerRowIfCreate);
      return sheet;
    }
    throw new Error('Aba "' + name + '" não encontrada na planilha.');
  }
  return sheet;
}
