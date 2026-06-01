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
