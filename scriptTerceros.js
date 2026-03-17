document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("cambiosForm");
  if (!form) return;

  const FLOW_URL = "";

  // =========================
  // ELEMENTOS BASE
  // =========================
  const modeGate = document.getElementById("modeGate");
  const formWrap = document.getElementById("formWrap");
  const btnIniciar = document.getElementById("btnIniciar");
  const btnCambiarTipo = document.getElementById("btnCambiarTipo");
  const modeOptions = document.getElementById("modeOptions");
  const selectedModeInput = document.getElementById("selectedMode");

  const certFlow = document.getElementById("certFlow");
  const implFlow = document.getElementById("implFlow");

  const sec0Details = document.getElementById("sec0Details");

  const nombresInput = document.getElementById("nombresCompletos");
  const numeroCambioInput = document.getElementById("numeroCambio");
  const numeroCambioWrap = document.getElementById("numeroCambioWrap");

  // Certificación
  const sec2Wrap = document.getElementById("sec2Wrap");
  const sec2Details = document.getElementById("sec2Details");
  const doneCertTareas = document.getElementById("doneCertTareas");

  // Implementación
  const sec3Details = document.getElementById("sec3Details");
  const sec4Wrap = document.getElementById("sec4Wrap");
  const sec4Details = document.getElementById("sec4Details");
  const doneImpl = document.getElementById("doneImpl");
  const donePost = document.getElementById("donePost");

  // Mensajes finales
  const flowEndMsg = document.getElementById("flowEndMsg");
  const finalNoteMsg = document.getElementById("finalNoteMsg");

  // Botones
  const btnReset = document.getElementById("btnReset");
  const btnSubmit = form.querySelector('button[type="submit"]');
  const btnPdf = document.getElementById("btnPdf");

  // Extras q11
  const q11Extra = document.getElementById("q11Extra");
  const q11Detalle = document.getElementById("q11Detalle");
  const btnQ11Continuar = document.getElementById("btnQ11Continuar");

  const titleText = (
    document.querySelector(".tcs-titlebar h1")?.textContent ||
    "CHECKLIST GESTIÓN DE CAMBIOS"
  ).trim();

  // =========================
  // IDS
  // =========================
  const GENERAL_IDS = ["q1", "q2"];
  const CERT_IDS = ["q8", "q9", "q10", "q11"];
  const IMPL_SEC1_IDS = ["q12"];
  const IMPL_SEC2_IDS = ["q15", "q16"];
  const IMPL_IDS = [...IMPL_SEC1_IDS, ...IMPL_SEC2_IDS];
  const ALL_Q_IDS = ["q0", ...GENERAL_IDS, ...CERT_IDS, ...IMPL_IDS];

  let selectedMode = null;
  let formStartedAt = "";
  let q11Continuado = false;

  // =========================
  // HELPERS
  // =========================
  function byId(id) {
    return document.getElementById(id);
  }

  function safeText(v) {
    return (v || "").toString().replace(/\s+/g, " ").trim();
  }

  function requiereNumeroCambio() {
    return true;
  }

  function setNumeroCambioVisibility() {
    if (!numeroCambioWrap) return;
    numeroCambioWrap.classList.remove("hidden");
    if (numeroCambioInput) numeroCambioInput.required = true;
  }

  function formatEcuadorDateTime(date = new Date()) {
    const formatter = new Intl.DateTimeFormat("es-EC", {
      timeZone: "America/Guayaquil",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false
    });

    const parts = {};
    formatter.formatToParts(date).forEach((p) => {
      parts[p.type] = p.value;
    });

    return `${parts.day}/${parts.month}/${parts.year} ${parts.hour}:${parts.minute}:${parts.second}`;
  }

  function ensureFormTimestamp() {
    if (!formStartedAt) {
      formStartedAt = formatEcuadorDateTime(new Date());
    }
  }

  function resetFormTimestamp() {
    formStartedAt = "";
  }

  function getAnswer(qid) {
    return form.querySelector(`input[name="${qid}_answer"]:checked`)?.value || "";
  }

  function clearAnswer(qid) {
    form.querySelectorAll(`input[name="${qid}_answer"]`).forEach((radio) => {
      radio.checked = false;
    });
  }

  function clearAnswers(ids) {
    ids.forEach(clearAnswer);
  }

  function setQuestionVisibility(qid, visible, clearIfHidden = true) {
    const block = byId(qid);
    if (!block) return;

    block.classList.toggle("hidden", !visible);

    if (!visible && clearIfHidden) {
      clearAnswer(qid);

      if (qid === "q11") {
        resetQ11Extra(true);
      }
    }
  }

  function hideAndClearQuestions(ids) {
    ids.forEach((id) => setQuestionVisibility(id, false, true));
  }

  function isDatosPersonalesValidos() {
    const okNombre = safeText(nombresInput?.value).length > 0;
    const okQChg =
      !requiereNumeroCambio() || safeText(numeroCambioInput?.value).length > 0;
    return !!selectedMode && okNombre && okQChg;
  }

  function getFlowIds() {
    if (selectedMode === "CERTIFICACION") {
      return [...GENERAL_IDS, ...CERT_IDS];
    }
    if (selectedMode === "IMPLEMENTACION") {
      return [...GENERAL_IDS, ...IMPL_IDS];
    }
    return [];
  }

  function resetQ11Extra(clearText = true) {
    q11Continuado = false;
    q11Extra?.classList.add("hidden");
    if (clearText && q11Detalle) {
      q11Detalle.value = "";
    }
  }

  function handleQ11Extra() {
    const q11Answer = getAnswer("q11");

    if (q11Answer === "SI") {
      q11Extra?.classList.remove("hidden");
    } else {
      resetQ11Extra(true);
    }
  }

  function getFlowState(ids) {
    const q0Answer = getAnswer("q0");

    if (!q0Answer) {
      return { finished: false, finishedByNo: false, fullYes: false, stopAt: "" };
    }

    if (q0Answer === "NO") {
      return { finished: true, finishedByNo: true, fullYes: false, stopAt: "q0" };
    }

    for (const qid of ids) {
      const answer = getAnswer(qid);

      if (!answer) {
        return { finished: false, finishedByNo: false, fullYes: false, stopAt: qid };
      }

      if (answer === "NO") {
        return { finished: true, finishedByNo: true, fullYes: false, stopAt: qid };
      }

      if (qid === "q11" && answer === "SI") {
        const detalle = safeText(q11Detalle?.value);
        if (!detalle || !q11Continuado) {
          return { finished: false, finishedByNo: false, fullYes: false, stopAt: qid };
        }
      }
    }

    return {
      finished: ids.length > 0,
      finishedByNo: false,
      fullYes: ids.length > 0,
      stopAt: ids[ids.length - 1] || ""
    };
  }

  function processSequential(ids) {
    let canContinue = true;

    for (const qid of ids) {
      const block = byId(qid);
      if (!block) continue;

      if (canContinue) {
        setQuestionVisibility(qid, true, false);
        const answer = getAnswer(qid);

        if (answer === "SI") {
          canContinue = true;
        } else if (answer === "NO") {
          canContinue = false;
        } else {
          canContinue = false;
        }
      } else {
        setQuestionVisibility(qid, false, true);
      }
    }
  }

  function resetMessages() {
    doneCertTareas?.classList.add("hidden");
    doneImpl?.classList.add("hidden");
    donePost?.classList.add("hidden");
    flowEndMsg?.classList.add("hidden");
    finalNoteMsg?.classList.add("hidden");
  }

  function showFinishedMessages(flowState) {
    if (flowEndMsg) {
      flowEndMsg.innerHTML = flowState.finishedByNo
        ? '✅ Se registró una respuesta <strong>NO</strong>. El formulario se dio por finalizado. Ya puede <strong>DESCARGAR PDF</strong>.'
        : '✅ Formulario completado. Ya puede <strong>DESCARGAR PDF</strong>.';

      flowEndMsg.classList.remove("hidden");
    }

    finalNoteMsg?.classList.remove("hidden");
  }

  function refreshIniciarState() {
    if (btnIniciar) btnIniciar.disabled = !selectedMode;
  }

  function refreshSubmitState() {
    const flowIds = getFlowIds();
    const flowState = getFlowState(flowIds);
    const ok = isDatosPersonalesValidos() && flowState.finished;

    if (btnSubmit) btnSubmit.disabled = !ok;
    if (btnPdf) btnPdf.disabled = !ok;
  }

  // =========================
  // UI PRINCIPAL
  // =========================
  function syncUI() {
    setNumeroCambioVisibility();
    refreshIniciarState();
    resetMessages();

    if (!formWrap || formWrap.classList.contains("hidden")) {
      refreshSubmitState();
      return;
    }

    const q0Answer = getAnswer("q0");

    certFlow?.classList.toggle("hidden", selectedMode !== "CERTIFICACION");
    implFlow?.classList.toggle("hidden", selectedMode !== "IMPLEMENTACION");

    // =========================
    // INFORMACIÓN GENERAL
    // =========================
    processSequential(GENERAL_IDS);

    if (sec0Details) {
      sec0Details.open = true;
    }

    if (q0Answer !== "SI") {
      hideAndClearQuestions(GENERAL_IDS);
      hideAndClearQuestions(CERT_IDS);
      hideAndClearQuestions(IMPL_IDS);
      resetQ11Extra(true);

      sec2Wrap?.classList.add("hidden");
      sec4Wrap?.classList.add("hidden");

      if (sec2Details) sec2Details.open = false;
      if (sec3Details) sec3Details.open = false;
      if (sec4Details) sec4Details.open = false;

      if (q0Answer === "NO") {
        showFinishedMessages(getFlowState(getFlowIds()));
      }

      refreshSubmitState();
      return;
    }

    const generalState = getFlowState(GENERAL_IDS);

    if (!generalState.finished) {
      hideAndClearQuestions(CERT_IDS);
      hideAndClearQuestions(IMPL_IDS);
      resetQ11Extra(true);

      sec2Wrap?.classList.add("hidden");
      sec4Wrap?.classList.add("hidden");

      if (sec2Details) sec2Details.open = false;
      if (sec3Details) sec3Details.open = false;
      if (sec4Details) sec4Details.open = false;

      refreshSubmitState();
      return;
    }

    if (generalState.finishedByNo) {
      hideAndClearQuestions(CERT_IDS);
      hideAndClearQuestions(IMPL_IDS);
      resetQ11Extra(true);

      sec2Wrap?.classList.add("hidden");
      sec4Wrap?.classList.add("hidden");

      if (sec2Details) sec2Details.open = false;
      if (sec3Details) sec3Details.open = false;
      if (sec4Details) sec4Details.open = false;

      showFinishedMessages(generalState);
      refreshSubmitState();
      return;
    }

    if (sec0Details) {
      sec0Details.open = false;
    }

    // =========================
    // CERTIFICACIÓN
    // =========================
    if (selectedMode === "CERTIFICACION") {
      hideAndClearQuestions(IMPL_IDS);
      sec4Wrap?.classList.add("hidden");
      if (sec3Details) sec3Details.open = false;
      if (sec4Details) sec4Details.open = false;

      sec2Wrap?.classList.remove("hidden");

      processSequential(CERT_IDS);
      handleQ11Extra();

      if (sec2Details) sec2Details.open = true;

      const flowState = getFlowState([...GENERAL_IDS, ...CERT_IDS]);

      if (flowState.fullYes) {
        doneCertTareas?.classList.remove("hidden");
      }

      if (flowState.finished) {
        showFinishedMessages(flowState);
        if (sec2Details) sec2Details.open = true;
      }
    }

    // =========================
    // IMPLEMENTACIÓN
    // =========================
    if (selectedMode === "IMPLEMENTACION") {
      hideAndClearQuestions(CERT_IDS);
      resetQ11Extra(true);
      sec2Wrap?.classList.add("hidden");
      if (sec2Details) sec2Details.open = false;

      processSequential(IMPL_IDS);

      if (sec3Details) sec3Details.open = true;

      const q12Answer = getAnswer("q12");

      if (q12Answer === "SI") {
        doneImpl?.classList.remove("hidden");
        sec4Wrap?.classList.remove("hidden");
        if (sec3Details) sec3Details.open = false;
        if (sec4Details) sec4Details.open = true;

        setQuestionVisibility("q15", true, false);

        if (getAnswer("q15") === "SI") {
          setQuestionVisibility("q16", true, false);
        } else if (getAnswer("q15") !== "NO") {
          setQuestionVisibility("q16", false, true);
        }
      } else if (q12Answer === "NO") {
        sec4Wrap?.classList.add("hidden");
        hideAndClearQuestions(IMPL_SEC2_IDS);
      } else {
        sec4Wrap?.classList.add("hidden");
        hideAndClearQuestions(IMPL_SEC2_IDS);
        if (sec3Details) sec3Details.open = true;
        if (sec4Details) sec4Details.open = false;
      }

      const flowState = getFlowState([...GENERAL_IDS, ...IMPL_IDS]);

      if (flowState.fullYes) {
        donePost?.classList.remove("hidden");
      }

      if (flowState.finished) {
        showFinishedMessages(flowState);
      }
    }

    refreshSubmitState();
  }

  // =========================
  // PAYLOAD
  // =========================
  function buildPayload() {
    ensureFormTimestamp();

    const payload = {
      selectedMode: selectedMode || "",
      nombresCompletos: safeText(nombresInput?.value),
      numeroCambio: safeText(numeroCambioInput?.value),
      fechaRealizacion: formStartedAt,
      formVersion: "3",
      source: window.location.href,
      q11_detalle: safeText(q11Detalle?.value),
      q11_continuado: q11Continuado
    };

    ALL_Q_IDS.forEach((qid) => {
      const answer = getAnswer(qid);
      payload[`${qid}_answer`] = answer;
      payload[`${qid}_check`] = answer === "SI";
    });

    return payload;
  }

  async function sendToPowerAutomate(payload) {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${txt}`);
    }

    return await res.text().catch(() => "");
  }

  // =========================
  // RESET TOTAL
  // =========================
  function resetAllToGate() {
    form.reset();
    clearAnswers(ALL_Q_IDS);

    modeGate?.classList.remove("hidden");
    formWrap?.classList.add("hidden");
    btnCambiarTipo?.classList.add("hidden");

    modeOptions?.querySelectorAll(".mode-card").forEach((b) => {
      b.setAttribute("aria-pressed", "false");
    });

    selectedMode = null;
    if (selectedModeInput) selectedModeInput.value = "";

    certFlow?.classList.add("hidden");
    implFlow?.classList.add("hidden");

    hideAndClearQuestions(GENERAL_IDS);
    hideAndClearQuestions(CERT_IDS);
    hideAndClearQuestions(IMPL_IDS);

    sec2Wrap?.classList.add("hidden");
    sec4Wrap?.classList.add("hidden");

    if (sec0Details) sec0Details.open = true;
    if (sec2Details) sec2Details.open = false;
    if (sec3Details) sec3Details.open = false;
    if (sec4Details) sec4Details.open = false;

    resetQ11Extra(true);
    resetMessages();
    resetFormTimestamp();
    setNumeroCambioVisibility();
    refreshIniciarState();
    refreshSubmitState();
  }

  // =========================
  // INIT
  // =========================
  if (btnPdf) btnPdf.disabled = true;
  if (btnSubmit) btnSubmit.disabled = true;
  resetAllToGate();

  // =========================
  // SELECCIÓN DE MODO
  // =========================
  modeOptions?.addEventListener("click", (e) => {
    const btn = e.target.closest(".mode-card");
    if (!btn) return;

    modeOptions.querySelectorAll(".mode-card").forEach((b) => {
      b.setAttribute("aria-pressed", "false");
    });

    btn.setAttribute("aria-pressed", "true");
    selectedMode = btn.dataset.mode || null;

    if (selectedModeInput) {
      selectedModeInput.value = selectedMode || "";
    }

    refreshIniciarState();
    refreshSubmitState();
  });

  // =========================
  // INICIAR
  // =========================
  btnIniciar?.addEventListener("click", () => {
    if (!selectedMode) {
      alert("Seleccione Certificación o Implementación.");
      return;
    }

    modeGate?.classList.add("hidden");
    formWrap?.classList.remove("hidden");
    btnCambiarTipo?.classList.remove("hidden");

    if (sec0Details) sec0Details.open = true;
    syncUI();
    sec0Details?.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  // =========================
  // CAMBIAR TIPO
  // =========================
  btnCambiarTipo?.addEventListener("click", () => {
    resetAllToGate();
  });

  // =========================
  // INPUTS CABECERA
  // =========================
  [nombresInput, numeroCambioInput].forEach((el) => {
    el?.addEventListener("input", () => {
      if (safeText(el.value)) {
        ensureFormTimestamp();
      }
      refreshSubmitState();
    });
  });

  // =========================
  // INPUT DETALLE Q11
  // =========================
  q11Detalle?.addEventListener("input", () => {
    q11Continuado = false;
    refreshSubmitState();
  });

  // =========================
  // BOTÓN CONTINUAR Q11
  // =========================
  btnQ11Continuar?.addEventListener("click", () => {
    const detalle = safeText(q11Detalle?.value);

    if (getAnswer("q11") !== "SI") return;

    if (!detalle) {
      alert("Debe ingresar el detalle del proceso automático antes de continuar.");
      q11Detalle?.focus();
      return;
    }

    q11Continuado = true;
    ensureFormTimestamp();
    syncUI();
  });

  // =========================
  // LÓGICA DE RADIOS
  // =========================
  form.addEventListener("change", (e) => {
    const target = e.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (target.type !== "radio") return;

    if (target.name === "q0_answer") {
      if (!safeText(nombresInput?.value)) {
        alert("Complete Nombres completos antes de continuar.");
        clearAnswer("q0");
        refreshSubmitState();
        return;
      }

      if (requiereNumeroCambio() && !safeText(numeroCambioInput?.value)) {
        alert("Complete el N° de Control de Cambios antes de continuar.");
        clearAnswer("q0");
        refreshSubmitState();
        return;
      }
    }

    if (target.name === "q11_answer") {
      q11Continuado = false;

      if (target.value === "SI") {
        q11Extra?.classList.remove("hidden");
      } else {
        resetQ11Extra(true);
      }
    }

    ensureFormTimestamp();
    syncUI();
  });

  // =========================
  // RESET
  // =========================
  btnReset?.addEventListener("click", () => {
    setTimeout(() => {
      resetAllToGate();
    }, 0);
  });

  // =========================
  // ENVIAR
  // =========================
  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    if (btnSubmit?.disabled) {
      alert("Complete el formulario para habilitar ENVIAR.");
      return;
    }

    const payload = buildPayload();

    const prevText = btnSubmit?.innerText || "ENVIAR";
    if (btnSubmit) {
      btnSubmit.disabled = true;
      btnSubmit.innerText = "ENVIANDO...";
    }

    try {
      await sendToPowerAutomate(payload);
      alert("Datos enviados correctamente.");
    } catch (err) {
      console.error(err);
      alert("No se pudo guardar en Excel. Revisa el Flow y posibles errores de CORS o permisos.");
    } finally {
      if (btnSubmit) btnSubmit.innerText = prevText;
      refreshSubmitState();
    }
  });

  // =========================
  // PDF
  // =========================
  function getItemText(qBlockId) {
    const block = byId(qBlockId);
    if (!block) return "";
    const textEl = block.querySelector(".q-text");
    return safeText(textEl?.textContent || "");
  }

  function rowsUntilStop(ids) {
    const rows = [];

    for (const qid of ids) {
      const block = byId(qid);
      if (!block) continue;

      const answer = getAnswer(qid);
      if (!answer) break;

      rows.push([getItemText(qid), answer]);

      if (qid === "q11" && answer === "SI") {
        const detalle = safeText(q11Detalle?.value);
        if (detalle) {
          rows.push(["Detalle del proceso automático", detalle]);
        }
      }

      if (answer === "NO") break;
    }

    return rows;
  }

  function buildPdfData() {
    ensureFormTimestamp();

    const nombre = safeText(nombresInput?.value);
    const numeroCambio = safeText(numeroCambioInput?.value);
    const fechaRealizacion = formStartedAt || formatEcuadorDateTime(new Date());

    const infoGeneral = [
      [getItemText("q0"), getAnswer("q0") || ""],
      ...rowsUntilStop(GENERAL_IDS)
    ];

    const certTareas = rowsUntilStop(CERT_IDS);
    const implementacion = rowsUntilStop(IMPL_SEC1_IDS);
    const postImplementacion = rowsUntilStop(IMPL_SEC2_IDS);

    return {
      nombre,
      numeroCambio,
      fechaRealizacion,
      infoGeneral,
      certTareas,
      implementacion,
      postImplementacion,
      selectedMode
    };
  }

  function generatePdf() {
    if (btnPdf?.disabled) {
      alert("Complete el formulario para habilitar la descarga del PDF.");
      return;
    }

    if (!window.jspdf?.jsPDF) {
      alert("No se cargó jsPDF. Verifique que los <script> CDN estén en el HTML.");
      return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });

    if (typeof doc.autoTable !== "function") {
      alert("No se cargó jsPDF AutoTable.");
      return;
    }

    const data = buildPdfData();

    if (!data.nombre) {
      alert("Complete 'Nombres completos' antes de generar el PDF.");
      return;
    }

    if (requiereNumeroCambio() && !data.numeroCambio) {
      alert("Complete 'N° de Control de Cambios' antes de generar el PDF.");
      return;
    }

    const pageW = doc.internal.pageSize.getWidth();
    const margin = 28;

    doc.setFillColor(30, 94, 230);
    doc.rect(margin, 22, pageW - margin * 2, 42, "F");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.setTextColor(255, 255, 255);
    doc.text(titleText, pageW / 2, 49, { align: "center" });

    doc.setTextColor(0, 0, 0);
    doc.setFontSize(10);

    let cursorY = 92;

    doc.setFont("helvetica", "normal");
    doc.text("Nombres completos:", margin, cursorY);
    doc.setFont("helvetica", "bold");
    doc.text(data.nombre, 170, cursorY);

    cursorY += 18;
    doc.setFont("helvetica", "normal");
    doc.text("N° de Control de Cambios:", margin, cursorY);
    doc.setFont("helvetica", "bold");
    doc.text(data.numeroCambio || "-", 170, cursorY);

    cursorY += 18;
    doc.setFont("helvetica", "normal");
    doc.text("Fecha de realización:", margin, cursorY);
    doc.setFont("helvetica", "bold");
    doc.text(data.fechaRealizacion, 170, cursorY);

    cursorY += 22;

    function drawSection(title, rows) {
      if (!rows || !rows.length) return;

      doc.autoTable({
        startY: cursorY,
        head: [[title, "Respuesta"]],
        body: rows,
        theme: "grid",
        margin: { left: margin, right: margin },
        styles: {
          font: "helvetica",
          fontSize: 9,
          cellPadding: 6,
          valign: "top",
          textColor: 0,
          overflow: "linebreak"
        },
        headStyles: {
          fillColor: [230, 238, 248],
          textColor: 0,
          fontStyle: "bold"
        },
        columnStyles: {
          0: {
            cellWidth: pageW - margin * 2 - 110,
            overflow: "linebreak"
          },
          1: {
            cellWidth: 110,
            halign: "center",
            valign: "middle"
          }
        }
      });

      cursorY = doc.lastAutoTable.finalY + 14;
    }

    drawSection("1. INFORMACIÓN GENERAL", data.infoGeneral);

    if (data.selectedMode === "CERTIFICACION") {
      drawSection("2. CERTIFICACIÓN DE TAREAS", data.certTareas);
    }

    if (data.selectedMode === "IMPLEMENTACION") {
      drawSection("2. IMPLEMENTACIÓN", data.implementacion);
      drawSection("3. POST IMPLEMENTACIÓN", data.postImplementacion);
    }

    const safeName = (data.nombre || "usuario")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^\w\s-]/g, "")
      .replace(/\s+/g, "_");

    const safeChg = (data.numeroCambio || "SIN_CHG").replace(/[^\w-]/g, "_");

    const fileName = `Checklist_${safeChg}_${safeName}.pdf`;
    doc.save(fileName);
  }

  btnPdf?.addEventListener("click", generatePdf);
});