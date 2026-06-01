/**
 * assets/app.js
 * Adaptativa 2026 — Integração HTML ↔ Python (pywebview)
 *
 * Todas as chamadas ao backend usam: window.pywebview.api.<método>(args)
 * que retornam Promises resolvidas com { ok, ... } ou { ok: false, error }.
 */

/* ══════════════════════════════════════════════════════════════════════════
   Utilitários
══════════════════════════════════════════════════════════════════════════ */

const log = document.getElementById("log");
const backendStatus = document.getElementById("backendStatus");

/** Adiciona uma linha ao painel de log. */
function appendLog(msg) {
  if (!log) return;
  log.textContent += msg + "\n";
  log.scrollTop = log.scrollHeight;
}

/** Atualiza o badge de status do backend na topbar. */
function setStatus(text, ok = true) {
  if (!backendStatus) return;
  backendStatus.textContent = text;
  backendStatus.style.setProperty(
    "--status-color",
    ok ? "var(--success)" : "var(--danger)"
  );
}

/** Retorna a API Python ou null (fallback para browser sem backend). */
function api() {
  return window.pywebview?.api ?? null;
}

/** Wrapper seguro: chama o backend e trata erros de conexão. */
async function call(method, ...args) {
  const a = api();
  if (!a) {
    const msg = "Backend Python não disponível neste ambiente.";
    appendLog("[ERRO] " + msg);
    return { ok: false, error: msg };
  }
  try {
    return await a[method](...args);
  } catch (err) {
    const msg = `Falha ao chamar ${method}(): ${err}`;
    appendLog("[ERRO] " + msg);
    return { ok: false, error: msg };
  }
}

/** Exibe resultado em um <pre> e no log. */
function showResult(preId, res) {
  const pre = document.getElementById(preId);
  if (!pre) return;
  if (res.ok) {
    pre.style.color = "var(--success)";
    const detail =
      res.path       ? `✔ ${res.path}`          :
      res.output_dir ? `✔ Pasta: ${res.output_dir}` :
      res.text       ? res.text                  :
      JSON.stringify(res, null, 2);
    pre.textContent = detail;
    appendLog("[OK] " + (res.path || res.output_dir || "Operação concluída."));
  } else {
    pre.style.color = "var(--danger)";
    pre.textContent = "✘ " + (res.error || "Erro desconhecido.");
    appendLog("[ERRO] " + res.error);
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   Inicialização: ping ao backend
══════════════════════════════════════════════════════════════════════════ */

async function initBackend() {
  // pywebview injeta a API de forma assíncrona; aguarda até 5 s
  for (let i = 0; i < 50; i++) {
    if (api()) break;
    await new Promise(r => setTimeout(r, 100));
  }

  if (!api()) {
    setStatus("Sem backend Python", false);
    appendLog("[AVISO] pywebview não detectado — backend inativo.");
    return;
  }

  const res = await call("ping");
  if (res.ok) {
    setStatus("Backend conectado");
    appendLog("[OK] " + res.message);
  } else {
    setStatus("Erro no backend", false);
  }
}

initBackend();

// Atualiza logs do backend periodicamente
setInterval(async () => {
  if (!api()) return;
  const res = await call("get_logs");
  if (res.ok && res.logs?.length) {
    // Só imprime linhas novas (compara pelo último texto do log)
    const current = log?.textContent ?? "";
    res.logs.forEach(line => {
      if (!current.includes(line)) appendLog(line);
    });
  }
}, 3000);

/* ══════════════════════════════════════════════════════════════════════════
   TAB: Gerador de TXT
══════════════════════════════════════════════════════════════════════════ */

document.getElementById("gerarTxt")?.addEventListener("click", async () => {
  const data = {
    item_base:   document.getElementById("item_base")?.value  ?? "",
    componente:  document.getElementById("componente")?.value ?? "",
    matriz_base: document.getElementById("matriz_base")?.value ?? "",
    etapa:       document.getElementById("etapa")?.value       ?? "",
    codigo_alvo: document.getElementById("codigo_alvo")?.value ?? "",
    ano_alvo:    document.getElementById("ano_alvo")?.value    ?? "",
    dificuldade: document.querySelector('input[name="dificuldade"]:checked')?.value ?? "Médio",
    tema:        document.getElementById("tema")?.value        ?? "",
    suporte:     document.getElementById("suporte")?.value     ?? "",
    restricoes:  document.getElementById("restricoes")?.value  ?? "",
    gabarito:    (document.getElementById("gabarito")?.value   ?? "A").toUpperCase(),
  };

  const res = await call("generate_txt", data);
  const saida = document.getElementById("txt_saida");
  if (saida) {
    saida.value = res.ok ? res.text : ("ERRO: " + res.error);
  }
  if (!res.ok) appendLog("[ERRO] " + res.error);
});

/* ══════════════════════════════════════════════════════════════════════════
   TAB: Formatador de item
══════════════════════════════════════════════════════════════════════════ */

/**
 * Valida os títulos do texto bruto antes de formatar.
 * Retorna um array de strings de erro (vazio = OK).
 *
 * Validação ativa:
 *  - Alternativas fora da ordem A→B→C→D (aceita maiúsculas ou minúsculas)
 *
 * Obs.: capitalização das seções e Texto-Base são aceitos como vierem.
 */
function validateTitles(raw) {
  const errors = [];
  const lines = raw.split(/\r?\n/);

  const altAny       = /^([A-Da-d])\)/;
  const expectedOrder = ['a', 'b', 'c', 'd'];
  let altIndex = 0;

  lines.forEach((line, i) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    const lineNum = i + 1;

    // Verifica ordem das alternativas (case-insensitive)
    const match = trimmed.match(altAny);
    if (match) {
      const letter = match[1].toLowerCase();
      const expected = expectedOrder[altIndex];
      if (expected && letter !== expected) {
        errors.push(`Linha ${lineNum}: alternativa "${match[1]})" fora de ordem — esperado "${expectedOrder[altIndex].toUpperCase()})"`);
      }
      altIndex++;
    }
  });

  return errors;
}

document.getElementById("fmt_btn")?.addEventListener("click", async () => {
  const raw  = document.getElementById("fmt_raw")?.value  ?? "";
  const path = document.getElementById("fmt_out")?.value  ?? "";
  const validationDiv = document.getElementById("fmt_validation");
  const resultPre     = document.getElementById("fmt_result");

  // ── Melhoria 5: Validação de títulos ──
  if (validationDiv) validationDiv.style.display = "none";
  if (resultPre) { resultPre.textContent = ""; resultPre.style.color = ""; }

  if (raw.trim()) {
    const erros = validateTitles(raw);
    if (erros.length > 0) {
      validationDiv.textContent =
        "⚠ Bloqueado — corrija os erros abaixo antes de formatar:\n\n" +
        erros.map(e => "• " + e).join("\n");
      validationDiv.style.display = "block";
      appendLog("[VALIDAÇÃO] " + erros.length + " erro(s) encontrado(s) nos títulos.");
      return; // bloqueia a geração
    }
  }

  // Se o campo de caminho estiver vazio, abre diálogo de salvar
  let outPath = path.trim();
  if (!outPath) {
    const dlg = await call("choose_file", "Escolha o local de saída (.docx)", [
      ["Documento Word", "*.docx"],
    ]);
    if (!dlg.ok || !dlg.path) return;
    outPath = dlg.path;
    const el = document.getElementById("fmt_out");
    if (el) el.value = outPath;
  }

  const res = await call("format_item_to_word", raw, outPath);
  showResult("fmt_result", res);

  // ── Melhoria 4: Abrir documento automaticamente após gerar ──
  if (res.ok && res.path) {
    appendLog("[INFO] Abrindo documento: " + res.path);
    // Tenta os nomes de método mais comuns no backend Python
    let opened = false;
    for (const method of ["open_file", "open_document", "os_startfile", "shell_open"]) {
      const a = api();
      if (a && typeof a[method] === "function") {
        try {
          const r = await a[method](res.path);
          if (r?.ok !== false) { opened = true; break; }
        } catch (_) {}
      }
    }
    if (!opened) {
      appendLog("[AVISO] Nenhum método de abertura encontrado no backend.");
      appendLog("[AVISO] Adicione ao seu api.py:");
      appendLog('  import os, subprocess, sys');
      appendLog('  def open_file(self, path):');
      appendLog('    try:');
      appendLog('      if sys.platform == "win32": os.startfile(path)');
      appendLog('      elif sys.platform == "darwin": subprocess.Popen(["open", path])');
      appendLog('      else: subprocess.Popen(["xdg-open", path])');
      appendLog('      return {"ok": True}');
      appendLog('    except Exception as e: return {"ok": False, "error": str(e)}');
    }
  }
});

/* ══════════════════════════════════════════════════════════════════════════
   TAB: Padrão SAFA
══════════════════════════════════════════════════════════════════════════ */

// Botão "Selecionar Word(s)"
document.getElementById("safa_pick")?.addEventListener("click", async () => {
  const res = await call("choose_files", "Selecionar arquivos Word", [
    ["Documentos Word", "*.docx"],
  ]);
  if (!res.ok || !res.paths?.length) return;
  const ta = document.getElementById("safa_files");
  if (ta) ta.value = res.paths.join("\n");
});

// Botão "Selecionar" pasta de saída
document.getElementById("safa_out_pick")?.addEventListener("click", async () => {
  const res = await call("choose_folder", "Selecionar pasta de saída");
  if (!res.ok || !res.path) return;
  const el = document.getElementById("safa_out");
  if (el) el.value = res.path;
});

// Botão "Processar Padrão SAFA"
document.getElementById("safa_btn")?.addEventListener("click", async () => {
  const filesText = document.getElementById("safa_files")?.value ?? "";
  const files = filesText.split("\n").map(s => s.trim()).filter(Boolean);
  const outDir   = document.getElementById("safa_out")?.value     ?? "";
  const ensinart = document.getElementById("safa_ensinart")?.checked ?? true;
  const selection= document.getElementById("safa_selection")?.value ?? "";

  if (!files.length) {
    appendLog("[AVISO] Nenhum arquivo Word informado.");
    return;
  }

  const res = await call("convert_safa", files, outDir, ensinart, selection);
  showResult("safa_result", res);

  if (res.ok && res.results) {
    res.results.forEach(r => {
      if (r.ok && r.report) {
        const rpt = r.report;
        appendLog(
          `[SAFA] ${Path.basename(r.file)} — ` +
          `${rpt.itens_ok} ok / ${rpt.itens_com_erro} erros / ` +
          `${rpt.itens_ignorados} ignorados`
        );
      } else if (!r.ok) {
        appendLog(`[ERRO] ${r.file}: ${r.error}`);
      }
    });
  }
});

/* ══════════════════════════════════════════════════════════════════════════
   TAB: Montagem das imagens (Word → PNG)
══════════════════════════════════════════════════════════════════════════ */

// Botão "Selecionar Word(s)"
document.getElementById("png_pick")?.addEventListener("click", async () => {
  const res = await call("choose_files", "Selecionar arquivos Word", [
    ["Documentos Word", "*.docx"],
  ]);
  if (!res.ok || !res.paths?.length) return;
  const ta = document.getElementById("png_files");
  if (ta) ta.value = res.paths.join("\n");
});

// Botão "Selecionar" pasta de saída
document.getElementById("png_out_pick")?.addEventListener("click", async () => {
  const res = await call("choose_folder", "Selecionar pasta de saída");
  if (!res.ok || !res.path) return;
  const el = document.getElementById("png_out");
  if (el) el.value = res.path;
});

// Botão "Converter Word para PNG"
document.getElementById("png_btn")?.addEventListener("click", async () => {
  const filesText = document.getElementById("png_files")?.value ?? "";
  const files = filesText.split("\n").map(s => s.trim()).filter(Boolean);
  const outDir = document.getElementById("png_out")?.value ?? "";
  const dpi    = parseInt(document.getElementById("png_dpi")?.value ?? "200", 10);
  const crop   = document.getElementById("png_crop")?.checked ?? true;

  if (!files.length) {
    appendLog("[AVISO] Nenhum arquivo Word informado.");
    return;
  }

  appendLog(`[PNG] Iniciando conversão de ${files.length} arquivo(s)…`);
  const res = await call("word_to_png", files, outDir, dpi, crop);
  showResult("png_result", res);

  if (res.ok) {
    appendLog(`[PNG] ${res.success}/${res.total} documentos convertidos → ${res.output_dir}`);
  }
});

/* ══════════════════════════════════════════════════════════════════════════
   Pequena shim para Path.basename (JS não tem Path nativo)
══════════════════════════════════════════════════════════════════════════ */
const Path = {
  basename: (p) => p.replace(/\\/g, "/").split("/").pop(),
};
