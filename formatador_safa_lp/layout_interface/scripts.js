// Alternância de Abas
function openTab(evt, tabId) {
    const tabContents = document.getElementsByClassName("tab-content");
    for (let i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove("active");
    }

    const tabLinks = document.getElementsByClassName("tab-link");
    for (let i = 0; i < tabLinks.length; i++) {
        tabLinks[i].classList.remove("active");
    }

    document.getElementById(tabId).classList.add("active");
    evt.currentTarget.classList.add("active");
    
    updateStatusBar(`Aba atual: ${evt.currentTarget.innerText}`);
}

// Funções de Utilitário Visual
function copyToClipboard(elementId) {
    const textarea = document.getElementById(elementId);
    textarea.select();
    document.execCommand("copy");
    
    const originalText = textarea.value;
    updateStatusBar("Copiado para a área de transferência!");
}

function clearField(elementId) {
    document.getElementById(elementId).value = "";
    updateStatusBar("Campo limpo.");
}

function updateStatusBar(message) {
    const bar = document.getElementById("status-bar");
    if (bar) bar.innerText = message;
}

// listeners de exemplo para conexão com backend futuro
window.addEventListener('load', () => {
    console.log("Interface visual carregada e pronta.");
});
let MATRIZES = [];
let CODIGOS_ATUAIS = [];
let CODIGO_SELECIONADO = null;

function openTab(evt, tabId) {
    const tabContents = document.getElementsByClassName("tab-content");
    for (let i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove("active");
    }

    const tabLinks = document.getElementsByClassName("tab-link");
    for (let i = 0; i < tabLinks.length; i++) {
        tabLinks[i].classList.remove("active");
    }

    document.getElementById(tabId).classList.add("active");
    evt.currentTarget.classList.add("active");
    updateStatusBar(`Aba atual: ${evt.currentTarget.innerText}`);
}

function copyToClipboard(elementId) {
    const textarea = document.getElementById(elementId);
    textarea.select();
    document.execCommand("copy");
    updateStatusBar("Copiado para a área de transferência.");
}

function clearField(elementId) {
    document.getElementById(elementId).value = "";
    updateStatusBar("Campo limpo.");
}

function updateStatusBar(message) {
    const bar = document.getElementById("status-bar");
    if (bar) bar.innerText = message;
}

function mostrarAlerta(mensagem, tipo = "erro") {
    const box = document.getElementById("f-alerta");
    if (!box) return;
    if (!mensagem) {
        box.innerHTML = "";
        return;
    }
    box.innerHTML = `<div class="alert ${tipo}">${mensagem}</div>`;
}

function getCampo(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
}

function normalizarCodigo(codigo) {
    codigo = String(codigo || "").trim().toUpperCase().replace(/\s+/g, "");
    if (/^D0*\d+$/.test(codigo)) {
        const n = parseInt(codigo.substring(1), 10);
        return `D${n}`;
    }
    return codigo;
}

function getCodigo(meta) {
    return String(meta?.código || meta?.codigo || "").trim();
}

function getDescricao(meta) {
    return String(meta?.descrição || meta?.descricao || "").trim();
}

function matrizSelecionada() {
    const id = getCampo("f-matriz");
    return MATRIZES.find((m) => m.id === id) || null;
}

function etapaSelecionada() {
    const matriz = matrizSelecionada();
    const etapaId = getCampo("f-etapa");
    if (!matriz) return null;
    return (matriz.etapas || []).find((e) => e.id === etapaId) || null;
}

function preencherMatrizes() {
    const select = document.getElementById("f-matriz");
    select.innerHTML = "";

    if (!MATRIZES.length) {
        select.innerHTML = '<option value="">Nenhuma matriz encontrada</option>';
        preencherEtapas();
        return;
    }

    MATRIZES.forEach((matriz) => {
        const opt = document.createElement("option");
        opt.value = matriz.id;
        opt.textContent = matriz.nome || matriz.id;
        select.appendChild(opt);
    });

    preencherEtapas();
}

function preencherEtapas() {
    const matriz = matrizSelecionada();
    const select = document.getElementById("f-etapa");
    select.innerHTML = "";

    if (!matriz) {
        select.innerHTML = '<option value="">Selecione a matriz primeiro</option>';
        CODIGOS_ATUAIS = [];
        atualizarCodigoCompleto();
        return;
    }

    const etapas = matriz.etapas || [];
    if (!etapas.length) {
        select.innerHTML = '<option value="">Nenhuma etapa encontrada</option>';
        CODIGOS_ATUAIS = [];
        atualizarCodigoCompleto();
        return;
    }

    etapas.forEach((etapa) => {
        const opt = document.createElement("option");
        opt.value = etapa.id;
        opt.textContent = etapa.nome || etapa.id;
        select.appendChild(opt);
    });

    carregarCodigosDaEtapa();
}

function carregarCodigosDaEtapa() {
    const etapa = etapaSelecionada();
    CODIGOS_ATUAIS = [];
    CODIGO_SELECIONADO = null;

    if (etapa) {
        CODIGOS_ATUAIS = (etapa.codigos || []).map((codigo) => {
            const cod = getCodigo(codigo);
            const desc = getDescricao(codigo);
            return {
                display: desc ? `${cod} - ${desc}` : cod,
                meta: codigo,
                codigo_normalizado: normalizarCodigo(cod),
            };
        });
    }

    atualizarCodigoCompleto();
}

function atualizarCodigoCompleto() {
    const input = document.getElementById("f-codigo");
    const completo = document.getElementById("f-codigo-completo");
    const codigoDigitado = input ? input.value : "";
    const alvo = normalizarCodigo(codigoDigitado);

    CODIGO_SELECIONADO = null;

    if (!alvo) {
        completo.value = "";
        completo.placeholder = "Digite o código para exibir a descrição completa";
        return;
    }

    const encontrado = CODIGOS_ATUAIS.find((item) => item.codigo_normalizado === alvo);

    if (!encontrado) {
        completo.value = "";
        completo.placeholder = "Código não encontrado para a matriz e etapa selecionadas";
        return;
    }

    CODIGO_SELECIONADO = encontrado;
    completo.value = encontrado.display;
}

async function carregarMatrizes() {
    try {
        if (window.pywebview && window.pywebview.api) {
            const resp = await window.pywebview.api.carregar_matrizes();
            if (!resp.ok) throw new Error(resp.erro || "Falha ao carregar matrizes.");
            MATRIZES = resp.matrizes && resp.matrizes.length ? resp.matrizes : (resp.dados?.matrizes || []);
        } else {
            // Fallback apenas para quando o HTML for aberto direto no navegador.
            const response = await fetch("../data/matrizes_portugues.json");
            const dados = await response.json();
            MATRIZES = (dados.matrizes || []).filter((m) => m.ativo !== false);
        }

        MATRIZES.sort((a, b) => (a.ordem || 99) - (b.ordem || 99));
        preencherMatrizes();
        updateStatusBar("Matrizes carregadas.");
    } catch (err) {
        MATRIZES = [];
        preencherMatrizes();
        updateStatusBar("Erro ao carregar matrizes.");
        mostrarAlerta(`Erro ao carregar matrizes: ${err.message || err}`);
    }
}

function validarFormularioPrompt() {
    const obrigatorios = [
        ["f-modo", "Informe o modo de trabalho."],
        ["f-matriz", "Informe a matriz-base."],
        ["f-etapa", "Informe a etapa da matriz."],
        ["f-ano-alvo", "Informe o ano em que o item será feito."],
        ["f-codigo", "Informe o código-alvo."],
        ["f-dificuldade", "Informe a dificuldade."],
        ["f-tipo-suporte", "Informe o tipo de suporte."],
    ];

    for (const [id, msg] of obrigatorios) {
        if (!getCampo(id)) return msg;
    }

    if (!CODIGO_SELECIONADO) {
        return "Código-alvo não encontrado para a matriz e etapa selecionadas. Digite um código válido, como D1 ou D01.";
    }

    const modo = getCampo("f-modo");
    const gabarito = getCampo("f-gabarito");
    if (modo === "Novo Item" && (!gabarito || gabarito === "Não informado")) {
        return "No modo Novo Item, informe o gabarito desejado.";
    }

    const textoBase = getCampo("f-texto-base");
    const fonteAutor = getCampo("f-fonte-autor");
    if (textoBase && !fonteAutor) {
        return "Informe a fonte/autor do texto-base.";
    }

    const tipoSuporte = getCampo("f-tipo-suporte");
    const descricaoSuporte = getCampo("f-descricao-suporte");
    const suportesQueExigemDescricao = ["Imagem", "Tabela", "Gráfico", "Malha", "Outro"];
    if (suportesQueExigemDescricao.includes(tipoSuporte) && !descricaoSuporte) {
        return "O suporte selecionado exige descrição/conteúdo.";
    }

    const itemBase = getCampo("f-item-base");
    if (!itemBase) {
        return modo === "Revisão" ? "Informe o item para revisão." : "Informe o item-base ou item de referência.";
    }

    return "";
}

function coletarDadosPrompt() {
    const matriz = matrizSelecionada();
    const etapa = etapaSelecionada();
    const meta = CODIGO_SELECIONADO?.meta || {};
    const codigo = getCodigo(meta) || getCampo("f-codigo");
    const descricao = getDescricao(meta);

    return {
        componente: "Língua Portuguesa",
        modo_trabalho: getCampo("f-modo"),
        matriz_base: matriz?.nome || getCampo("f-matriz"),
        matriz_id: matriz?.id || getCampo("f-matriz"),
        etapa_matriz: etapa?.nome || getCampo("f-etapa"),
        etapa_id: etapa?.id || getCampo("f-etapa"),
        ano_alvo: getCampo("f-ano-alvo"),
        codigo_alvo: codigo,
        codigo_digitado: getCampo("f-codigo"),
        codigo_completo: document.getElementById("f-codigo-completo").value,
        descricao_codigo: descricao,
        dificuldade: getCampo("f-dificuldade"),
        gabarito: getCampo("f-gabarito"),
        tipo_suporte: getCampo("f-tipo-suporte"),
        descricao_suporte: getCampo("f-descricao-suporte"),
        tema_contexto: getCampo("f-tema-contexto"),
        texto_base: getCampo("f-texto-base"),
        fonte_autor: getCampo("f-fonte-autor"),
        restricoes: getCampo("f-restricoes"),
        item_base_ou_revisao: getCampo("f-item-base"),
    };
}

async function gerarPrompt() {
    mostrarAlerta("");
    const erro = validarFormularioPrompt();
    if (erro) {
        mostrarAlerta(erro);
        updateStatusBar("Corrija os campos obrigatórios.");
        return;
    }

    const dados = coletarDadosPrompt();

    try {
        if (!window.pywebview || !window.pywebview.api) {
            throw new Error("API Python não disponível. Abra pelo app.py com pywebview.");
        }
        const resp = await window.pywebview.api.gerar_prompt(dados);
        if (!resp.ok) throw new Error(resp.erro || "Erro ao gerar prompt.");
        document.getElementById("f-prompt-result").value = resp.prompt;
        updateStatusBar("Prompt gerado com sucesso.");
    } catch (err) {
        mostrarAlerta(`Erro ao gerar prompt: ${err.message || err}`);
        updateStatusBar("Erro ao gerar prompt.");
    }
}

window.addEventListener("load", () => {
    updateStatusBar("Interface visual carregada. Carregando matrizes...");

    const matriz = document.getElementById("f-matriz");
    const etapa = document.getElementById("f-etapa");
    const codigo = document.getElementById("f-codigo");
    const btnGerar = document.getElementById("btn-gerar-prompt");

    if (matriz) matriz.addEventListener("change", preencherEtapas);
    if (etapa) etapa.addEventListener("change", carregarCodigosDaEtapa);
    if (codigo) codigo.addEventListener("input", atualizarCodigoCompleto);
    if (btnGerar) btnGerar.addEventListener("click", gerarPrompt);

    carregarMatrizes();
});
