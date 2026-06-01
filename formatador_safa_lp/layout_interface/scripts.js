const statusEl = document.getElementById("status");
const saidaEl = document.getElementById("saida");

document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));

    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

function coletarParametros() {
  const ids = [
    "componente",
    "modo_trabalho",
    "matriz_base",
    "etapa_matriz",
    "ano_alvo",
    "codigo_alvo",
    "dificuldade",
    "gabarito",
    "tipo_suporte",
    "tema_contexto",
    "texto_base",
    "fonte_autor",
    "descricao_suporte",
    "restricoes",
    "item_base"
  ];

  const dados = {};
  ids.forEach((id) => {
    const el = document.getElementById(id);
    dados[id] = el ? el.value : "";
  });

  return dados;
}

async function testarBackend() {
  try {
    if (!window.pywebview || !window.pywebview.api) {
      saidaEl.textContent =
        "Interface abriu, mas a API Python ainda não está disponível. " +
        "Se a tela foi aberta no navegador, isso é esperado. Rode com pywebview pelo app.py.";
      return;
    }

    const resp = await window.pywebview.api.ping();
    saidaEl.textContent = JSON.stringify(resp, null, 2);
  } catch (err) {
    saidaEl.textContent = "Erro ao testar backend: " + err;
  }
}

async function carregarMatrizes() {
  try {
    if (!window.pywebview || !window.pywebview.api) {
      saidaEl.textContent =
        "A API Python não está disponível. Não foi possível carregar matrizes pelo backend.";
      return;
    }

    const resp = await window.pywebview.api.carregar_matrizes();
    saidaEl.textContent = JSON.stringify(resp, null, 2);
  } catch (err) {
    saidaEl.textContent = "Erro ao carregar matrizes: " + err;
  }
}

async function gerarPromptTeste() {
  try {
    const parametros = coletarParametros();

    if (!window.pywebview || !window.pywebview.api) {
      saidaEl.textContent =
        "A API Python não está disponível. A tela está funcionando, mas o backend não respondeu.";
      return;
    }

    const resp = await window.pywebview.api.gerar_prompt(parametros);
    saidaEl.textContent = JSON.stringify(resp, null, 2);
  } catch (err) {
    saidaEl.textContent = "Erro ao gerar prompt: " + err;
  }
}

function limparSaida() {
  saidaEl.textContent = "Aguardando ação...";
}

function funcaoNaoImplementada(nome) {
  saidaEl.textContent =
    nome +
    " ainda não está implementado na interface provisória. " +
    "A próxima etapa é ligar este botão ao módulo correspondente em core/.";
}

window.addEventListener("load", () => {
  statusEl.textContent =
    "Interface carregada. Se você está vendo esta mensagem, index.html, style.css e scripts.js estão funcionando.";
});
