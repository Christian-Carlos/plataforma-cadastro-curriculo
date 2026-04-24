function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cadastro Profissional')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

const ABA_CADASTROS = "CADASTROS";

const ID_PASTA_FOTOS = "13RaTpoFBimCnUIw3XqrJM3t0q0TGAnDs";
const ID_PASTA_CURRICULOS = "1k6Cin7gjaLZKK8fuS8Ajvon-Za-k_VeT";
const ID_PASTA_COMPROVANTE = "15EoUDSLRgCK9LpmuQCnUUxZ3Pm5asW6k";

function obterAbaCadastros_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CADASTROS);
  if (!aba) throw new Error('Aba CADASTROS não encontrada.');
  return aba;
}

function normalizarTexto_(valor) {
  return String(valor || '').trim();
}

function normalizarCPF_(valor) {
  return String(valor || '').replace(/\D/g, '').trim();
}

function normalizarEmail_(valor) {
  return String(valor || '').toLowerCase().trim();
}

function normalizarTelefone_(valor) {
  return String(valor || '').replace(/\D/g, '').trim();
}

function normalizarCEP_(valor) {
  return String(valor || '').replace(/\D/g, '').trim();
}

function formatarCPF_(valor) {
  const cpf = normalizarCPF_(valor);
  if (cpf.length !== 11) return normalizarTexto_(valor);
  return cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function formatarCEP_(valor) {
  const cep = normalizarCEP_(valor);
  if (cep.length !== 8) return normalizarTexto_(valor);
  return cep.replace(/(\d{5})(\d{3})/, '$1-$2');
}

function formatarTelefone_(valor) {
  const tel = normalizarTelefone_(valor);

  if (tel.length === 11) {
    return tel.replace(/(\d{2})(\d{1})(\d{4})(\d{4})/, '$1 $2 $3-$4');
  }

  if (tel.length === 10) {
    return tel.replace(/(\d{2})(\d{4})(\d{4})/, '$1 $2-$3');
  }

  return normalizarTexto_(valor);
}

function formatarRG_(valor) {
  const rg = String(valor || '').replace(/\D/g, '');

  if (rg.length === 10) {
    return rg.replace(/(\d{2})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
  }

  return normalizarTexto_(valor);
}

function validarCpfBackend(cpf) {
  cpf = normalizarCPF_(cpf);
  if (cpf.length !== 11 || /^(\d)\1+$/.test(cpf)) return false;

  let soma = 0;
  let resto = 0;

  for (let i = 1; i <= 9; i++) {
    soma += parseInt(cpf.substring(i - 1, i), 10) * (11 - i);
  }

  resto = (soma * 10) % 11;
  if (resto === 10 || resto === 11) resto = 0;
  if (resto !== parseInt(cpf.substring(9, 10), 10)) return false;

  soma = 0;
  for (let i = 1; i <= 10; i++) {
    soma += parseInt(cpf.substring(i - 1, i), 10) * (12 - i);
  }

  resto = (soma * 10) % 11;
  if (resto === 10 || resto === 11) resto = 0;

  return resto === parseInt(cpf.substring(10, 11), 10);
}

function validarEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizarTexto_(email));
}

function gerarHashSenha_(senha) {
  const texto = normalizarTexto_(senha);
  if (!texto) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto);
  return Utilities.base64Encode(digest);
}

function calcularIdade_(dataNasc) {
  const texto = normalizarTexto_(dataNasc);
  if (!texto) return '';

  let nascimento;
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
    const partes = texto.split('-').map(Number);
    nascimento = new Date(partes[0], partes[1] - 1, partes[2]);
  } else {
    nascimento = new Date(texto);
  }

  if (isNaN(nascimento.getTime())) return '';

  const hoje = new Date();
  let idade = hoje.getFullYear() - nascimento.getFullYear();
  const mesAtual = hoje.getMonth();
  const diaAtual = hoje.getDate();
  const mesNasc = nascimento.getMonth();
  const diaNasc = nascimento.getDate();

  if (mesAtual < mesNasc || (mesAtual === mesNasc && diaAtual < diaNasc)) {
    idade--;
  }

  return idade >= 0 ? idade : '';
}

function serializarParaInputDate_(valor) {
  if (!valor) return '';

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const texto = normalizarTexto_(valor);

  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
    return texto;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const partes = texto.split('/');
    return partes[2] + '-' + partes[1] + '-' + partes[0];
  }

  return texto;
}

function gerarId_(cpfLimpo) {
  return 'ID-' + cpfLimpo;
}

function obterPastaDrive_(idPasta, descricao) {
  if (!idPasta) {
    throw new Error('ID da pasta de ' + descricao + ' não configurado.');
  }

  try {
    return DriveApp.getFolderById(idPasta);
  } catch (e) {
    throw new Error('Pasta de ' + descricao + ' não encontrada ou sem permissão.');
  }
}

function processarUpload_(arquivo, pastaId, nomeArquivo, descricao) {
  if (!arquivo) return '';

  const pasta = obterPastaDrive_(pastaId, descricao);
  const nomeLimpo = normalizarTexto_(arquivo.name || 'arquivo')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w.\- ]/g, '')
    .replace(/\s+/g, '_');

  const blob = Utilities.newBlob(
    Utilities.base64Decode(arquivo.base64),
    arquivo.mimeType,
    nomeArquivo + '_' + nomeLimpo
  );

  const novoArquivo = pasta.createFile(blob);

  try {
    novoArquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return novoArquivo.getUrl();
}

function obterMapaColunas_(aba) {
  const cabecalhos = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 32)).getDisplayValues()[0];
  const mapa = {};
  cabecalhos.forEach(function(nome, i) {
    mapa[normalizarTexto_(nome).toLowerCase()] = i + 1;
  });

  return {
    id: mapa['id'] || 1,
    dataCadastro: mapa['data cadastro'] || 2,
    dataAtualizacao: mapa['data atualização'] || mapa['data atualizacao'] || 3,
    nome: mapa['nome completo'] || 4,
    cpf: mapa['cpf'] || 5,
    cpfNormalizado: mapa['cpf normalizado'] || 6,
    rg: mapa['rg'] || 7,
    dataNasc: mapa['data nascimento'] || 8,
    telefone: mapa['telefone'] || 9,
    email: mapa['e-mail'] || mapa['email'] || 10,
    cep: mapa['cep'] || 11,
    endereco: mapa['endereço'] || mapa['endereco'] || 12,
    genero: mapa['gênero'] || mapa['genero'] || 13,
    idade: mapa['idade'] || 14,
    escolaridade: mapa['escolaridade'] || 15,
    cursos: mapa['cursos'] || 16,
    experiencia: mapa['experiência'] || mapa['experiencia'] || 17,
    areaAtuacao: mapa['área de atuação'] || mapa['area de atuação'] || mapa['área de atuacao'] || mapa['area de atuacao'] || 18,
    objetivo: mapa['objetivo profissional'] || 19,
    linkFoto: mapa['link da foto'] || 20,
    linkCurriculo: mapa['link do currículo'] || mapa['link do curriculo'] || 21,
    linkComprovante: mapa['link do comprovante'] || 22,
    status: mapa['status'] || 23,
    senha: mapa['senha'] || 24,
    estadoCivil: mapa['estado civil'] || 25,
    possuiFilhos: mapa['possui filhos'] || 26,
    pcd: mapa['é pcd?'] || mapa['e pcd?'] || mapa['é pcd'] || mapa['e pcd'] || 27,
    possuiCnh: mapa['possui cnh'] || 28,
    categoriaCnh: mapa['categoria da cnh'] || 29,
    idioma: mapa['fala algum idioma'] || 30,
    primeiroEmprego: mapa['1º emprego'] || mapa['1o emprego'] || mapa['1° emprego'] || 31,
    disponibilidadeHorario: mapa['disponibilidade de horário'] || mapa['disponibilidade de horario'] || 32
  };
}

function montarLinhaVazia_(tamanho) {
  return Array.from({ length: tamanho }, function() { return ''; });
}

function montarObjetoLinha_(linha, numeroLinha, col) {
  return {
    linha: numeroLinha,
    encontrado: true,
    dados: {
      nome: linha[col.nome - 1] || '',
      cpf: linha[col.cpf - 1] || '',
      rg: linha[col.rg - 1] || '',
      dataNasc: serializarParaInputDate_(linha[col.dataNasc - 1]),
      telefone: linha[col.telefone - 1] || '',
      email: linha[col.email - 1] || '',
      cep: linha[col.cep - 1] || '',
      endereco: linha[col.endereco - 1] || '',
      genero: linha[col.genero - 1] || '',
      idade: linha[col.idade - 1] || '',
      estadoCivil: linha[col.estadoCivil - 1] || '',
      possuiFilhos: linha[col.possuiFilhos - 1] || '',
      pcd: linha[col.pcd - 1] || '',
      possuiCnh: linha[col.possuiCnh - 1] || '',
      categoriaCnh: linha[col.categoriaCnh - 1] || '',
      idioma: linha[col.idioma - 1] || '',
      primeiroEmprego: linha[col.primeiroEmprego - 1] || '',
      disponibilidadeHorario: linha[col.disponibilidadeHorario - 1] || '',
      escolaridade: linha[col.escolaridade - 1] || '',
      cursos: linha[col.cursos - 1] || '',
      experiencia: linha[col.experiencia - 1] || '',
      areaAtuacao: linha[col.areaAtuacao - 1] || '',
      objetivo: linha[col.objetivo - 1] || '',
      linkFoto: linha[col.linkFoto - 1] || '',
      linkCurriculo: linha[col.linkCurriculo - 1] || '',
      linkComprovante: linha[col.linkComprovante - 1] || '',
      status: linha[col.status - 1] || 'ATIVO'
    }
  };
}

function buscarCpf(cpfLimpoBusca) {
  const cpfLimpo = normalizarCPF_(cpfLimpoBusca);
  const aba = obterAbaCadastros_();
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = Math.max(aba.getLastColumn(), 31);
  const col = obterMapaColunas_(aba);

  if (ultimaLinha < 2) return { encontrado: false };

  const dados = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();

  for (let i = 0; i < dados.length; i++) {
    const cpfNormalizadoLinha = normalizarCPF_(dados[i][col.cpfNormalizado - 1] || '');
    const cpfFormatadoLinha = normalizarCPF_(dados[i][col.cpf - 1] || '');

    if (cpfNormalizadoLinha === cpfLimpo || cpfFormatadoLinha === cpfLimpo) {
      return montarObjetoLinha_(dados[i], i + 2, col);
    }
  }

  return { encontrado: false };
}

function validarSenhaEdicao(cpfInformado, senhaInformada) {
  const busca = buscarCpf(cpfInformado);
  if (!busca.encontrado) {
    return { ok: false, mensagem: 'Cadastro não encontrado.' };
  }

  const aba = obterAbaCadastros_();
  const col = obterMapaColunas_(aba);
  const ultimaColuna = Math.max(aba.getLastColumn(), 31);
  const linha = aba.getRange(busca.linha, 1, 1, ultimaColuna).getValues()[0];
  const hashSalvo = normalizarTexto_(linha[col.senha - 1] || '');

  if (!hashSalvo) {
    return { ok: true, semSenha: true, mensagem: 'Cadastro localizado sem senha cadastrada. Defina uma senha ao salvar.' };
  }

  const hashInformado = gerarHashSenha_(senhaInformada);
  if (hashInformado !== hashSalvo) {
    return { ok: false, mensagem: 'Senha incorreta.' };
  }

  return { ok: true, semSenha: false, mensagem: 'Senha validada com sucesso.' };
}

function validarDadosCadastro_(dadosForm, modo, buscaExistente, hashSenhaExistente) {
  const nome = normalizarTexto_(dadosForm.nome);
  const cpfLimpo = normalizarCPF_(dadosForm.cpf);
  const email = normalizarEmail_(dadosForm.email);
  const cep = normalizarCEP_(dadosForm.cep);
  const senha = normalizarTexto_(dadosForm.senha);
  const confirmarSenha = normalizarTexto_(dadosForm.confirmarSenha);

  if (!nome) throw new Error('Informe o nome completo.');
  if (!cpfLimpo) throw new Error('Informe o CPF.');
  if (!validarCpfBackend(cpfLimpo)) throw new Error('CPF inválido.');
  if (!email) throw new Error('Informe o e-mail.');
  if (!validarEmail_(email)) throw new Error('E-mail inválido.');
  if (!cep) throw new Error('Informe o CEP.');
  if (cep && cep.length !== 8) throw new Error('CEP inválido.');

  if (!normalizarTexto_(dadosForm.dataNasc)) throw new Error('Informe a data de nascimento.');
  if (!normalizarTexto_(dadosForm.telefone)) throw new Error('Informe o telefone / WhatsApp.');

  var telefoneNumeros = String(dadosForm.telefone || '').replace(/\D/g, '');
  if (telefoneNumeros.length !== 11) {
    throw new Error('O telefone deve conter exatamente 11 dígitos.');
  }

  var ddd = telefoneNumeros.substring(0, 2);
  var dddsValidos = {
    '11': true,'12': true,'13': true,'14': true,'15': true,'16': true,'17': true,'18': true,'19': true,
    '21': true,'22': true,'24': true,'27': true,'28': true,
    '31': true,'32': true,'33': true,'34': true,'35': true,'37': true,'38': true,
    '41': true,'42': true,'43': true,'44': true,'45': true,'46': true,
    '47': true,'48': true,'49': true,
    '51': true,'53': true,'54': true,'55': true,
    '61': true,'62': true,'63': true,'64': true,'65': true,'66': true,'67': true,'68': true,'69': true,
    '71': true,'73': true,'74': true,'75': true,'77': true,
    '79': true,
    '81': true,'82': true,'83': true,'84': true,'85': true,'86': true,'87': true,'88': true,'89': true,
    '91': true,'92': true,'93': true,'94': true,'95': true,'96': true,'97': true,'98': true,'99': true
  };

  if (!dddsValidos[ddd]) {
    throw new Error('O DDD informado no telefone é inválido.');
  }

  if (/^(\d)\1{10}$/.test(telefoneNumeros)) {
    throw new Error('O telefone informado é inválido.');
  }
  if (!normalizarTexto_(dadosForm.endereco)) throw new Error('Informe o endereço completo.');
  if (!normalizarTexto_(dadosForm.genero)) throw new Error('Informe o gênero.');
  if (!normalizarTexto_(dadosForm.estadoCivil)) throw new Error('Informe o estado civil.');
  if (!normalizarTexto_(dadosForm.possuiFilhos)) throw new Error('Informe se possui filhos.');
  if (!normalizarTexto_(dadosForm.pcd)) throw new Error('Informe se é PCD.');
  if (!normalizarTexto_(dadosForm.possuiCnh)) throw new Error('Informe se possui CNH.');
  if (!normalizarTexto_(dadosForm.primeiroEmprego)) throw new Error('Informe se é 1º emprego.');
  if (!normalizarTexto_(dadosForm.disponibilidadeHorario)) throw new Error('Informe a disponibilidade de horário.');
  if (!normalizarTexto_(dadosForm.escolaridade)) throw new Error('Informe a escolaridade.');
  if (!normalizarTexto_(dadosForm.cursos)) throw new Error('Selecione pelo menos um curso.');
  if (normalizarTexto_(dadosForm.primeiroEmprego).toLowerCase() !== 'sim' && !normalizarTexto_(dadosForm.experiencia)) {
    throw new Error('Informe a experiência profissional.');
  }
    if (!normalizarTexto_(dadosForm.objetivo)) throw new Error('Informe o objetivo profissional.');

  if (modo !== 'edicao' && buscaExistente.encontrado) {
    throw new Error('Já existe cadastro com este CPF.');
  }

  if (modo === 'novo') {
    if (!dadosForm.comp) throw new Error('Anexe o comprovante de escolaridade.');
    if (!senha) throw new Error('Informe uma senha para o cadastro.');
    if (senha.length < 4) throw new Error('A senha deve ter pelo menos 4 caracteres.');
    if (senha !== confirmarSenha) throw new Error('A confirmação da senha não confere.');
  } else {
    if (!hashSenhaExistente && !senha) {
      throw new Error('Defina uma senha para proteger este cadastro.');
    }
    if (senha || confirmarSenha) {
      if (senha.length < 4) throw new Error('A senha deve ter pelo menos 4 caracteres.');
      if (senha !== confirmarSenha) throw new Error('A confirmação da senha não confere.');
    }
  }
}

function salvarCadastro(dadosForm) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const aba = obterAbaCadastros_();
    const col = obterMapaColunas_(aba);
    const modo = normalizarTexto_(dadosForm.modo).toLowerCase() === 'edicao' ? 'edicao' : 'novo';
    const cpfLimpo = normalizarCPF_(dadosForm.cpf);

    const busca = buscarCpf(cpfLimpo);
    const ultimaColuna = Math.max(aba.getLastColumn(), 31);
    let hashSenhaExistente = '';
    let linkComprovanteExistente = '';
    if (busca.encontrado) {
      const linhaExistente = aba.getRange(busca.linha, 1, 1, ultimaColuna).getValues()[0];
      hashSenhaExistente = normalizarTexto_(linhaExistente[col.senha - 1] || '');
      linkComprovanteExistente = normalizarTexto_(linhaExistente[col.linkComprovante - 1] || '');
    }
    validarDadosCadastro_(dadosForm, modo, busca, hashSenhaExistente);

    if (modo === 'edicao' && !dadosForm.comp && !linkComprovanteExistente) {
      throw new Error('Anexe o comprovante de escolaridade.');
    }

    const agora = new Date();

    let linkFoto = '';
    let linkCurr = '';
    let linkComp = '';

    if (dadosForm.foto) {
      linkFoto = processarUpload_(dadosForm.foto, ID_PASTA_FOTOS, 'FOTO_' + cpfLimpo, 'fotos');
    }

    if (dadosForm.curr) {
      linkCurr = processarUpload_(dadosForm.curr, ID_PASTA_CURRICULOS, 'CV_' + cpfLimpo, 'currículos');
    }

    if (dadosForm.comp) {
      linkComp = processarUpload_(dadosForm.comp, ID_PASTA_COMPROVANTE, 'COMP_' + cpfLimpo, 'comprovantes');
    }

    const tamanhoLinha = Math.max(aba.getLastColumn(), 32);
    const linha = montarLinhaVazia_(tamanhoLinha);

    linha[col.id - 1] = gerarId_(cpfLimpo);
    linha[col.dataCadastro - 1] = agora;
    linha[col.dataAtualizacao - 1] = agora;
    linha[col.nome - 1] = normalizarTexto_(dadosForm.nome).toUpperCase();
    linha[col.cpf - 1] = formatarCPF_(cpfLimpo);
    linha[col.cpfNormalizado - 1] = cpfLimpo;
    linha[col.rg - 1] = formatarRG_(dadosForm.rg);
    linha[col.dataNasc - 1] = normalizarTexto_(dadosForm.dataNasc);
    linha[col.telefone - 1] = formatarTelefone_(dadosForm.telefone);
    linha[col.email - 1] = normalizarEmail_(dadosForm.email);
    linha[col.cep - 1] = formatarCEP_(dadosForm.cep);
    linha[col.endereco - 1] = normalizarTexto_(dadosForm.endereco);
    linha[col.genero - 1] = normalizarTexto_(dadosForm.genero);
    linha[col.idade - 1] = calcularIdade_(dadosForm.dataNasc);
    linha[col.estadoCivil - 1] = normalizarTexto_(dadosForm.estadoCivil);
    linha[col.possuiFilhos - 1] = normalizarTexto_(dadosForm.possuiFilhos);
    linha[col.pcd - 1] = normalizarTexto_(dadosForm.pcd);
    linha[col.possuiCnh - 1] = normalizarTexto_(dadosForm.possuiCnh);
    linha[col.categoriaCnh - 1] = normalizarTexto_(dadosForm.categoriaCnh);
    linha[col.idioma - 1] = normalizarTexto_(dadosForm.idioma);
    linha[col.primeiroEmprego - 1] = normalizarTexto_(dadosForm.primeiroEmprego) || 'Não';
    linha[col.disponibilidadeHorario - 1] = normalizarTexto_(dadosForm.disponibilidadeHorario);
    linha[col.escolaridade - 1] = normalizarTexto_(dadosForm.escolaridade);
    linha[col.cursos - 1] = normalizarTexto_(dadosForm.cursos);
    linha[col.experiencia - 1] = (normalizarTexto_(dadosForm.primeiroEmprego).toLowerCase() === 'sim') ? 'Sem experiência' : normalizarTexto_(dadosForm.experiencia);
    linha[col.areaAtuacao - 1] = normalizarTexto_(dadosForm.areaAtuacao);
    linha[col.objetivo - 1] = normalizarTexto_(dadosForm.objetivo);
    linha[col.linkFoto - 1] = linkFoto;
    linha[col.linkCurriculo - 1] = linkCurr;
    linha[col.linkComprovante - 1] = linkComp;
    linha[col.senha - 1] = normalizarTexto_(dadosForm.senha) ? gerarHashSenha_(dadosForm.senha) : hashSenhaExistente;

    if (modo === 'edicao' && busca.encontrado) {
      const antiga = aba.getRange(busca.linha, 1, 1, tamanhoLinha).getValues()[0];

      linha[col.dataCadastro - 1] = antiga[col.dataCadastro - 1] || agora;

      if (!linkFoto) linha[col.linkFoto - 1] = antiga[col.linkFoto - 1];
      if (!linkCurr) linha[col.linkCurriculo - 1] = antiga[col.linkCurriculo - 1];
      if (!linkComp) linha[col.linkComprovante - 1] = antiga[col.linkComprovante - 1];

      aba.getRange(busca.linha, 1, 1, tamanhoLinha).setValues([linha]);

      return {
        status: 'ok',
        modo: 'edicao',
        mensagem: 'Currículo atualizado com sucesso.'
      };
    }

    if (busca.encontrado) {
      throw new Error('Já existe cadastro com este CPF.');
    }

    aba.appendRow(linha);

    return {
      status: 'ok',
      modo: 'novo',
      mensagem: 'Cadastro realizado com sucesso.'
    };

  } catch (e) {
    throw new Error(e.message || 'Erro ao salvar cadastro.');
  } finally {
    lock.releaseLock();
  }
}
