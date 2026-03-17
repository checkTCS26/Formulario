document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("cambiosForm");

  const FLOW_URL = "";

  const modeGate = document.getElementById("modeGate");
  const formWrap = document.getElementById("formWrap");
  const btnIniciar = document.getElementById("btnIniciar");
  const btnCambiarTipo = document.getElementById("btnCambiarTipo");
  const modeOptions = document.getElementById("modeOptions");
  const selectedModeInput = document.getElementById("selectedMode");

  const certFlow = document.getElementById("certFlow");
  const implFlow = document.getElementById("implFlow");

  const secInfoDetails = document.getElementById("secInfoDetails");
  const nombresInput = document.getElementById("nombresCompletos");
  const numeroCambioInput = document.getElementById("numeroCambio");
  const numeroCambioWrap = document.getElementById("numeroCambioWrap");
  const q0 = document.getElementById("infoConfirm");
  const doneInfo = document.getElementById("doneInfo");

  const secPreDetails = document.getElementById("secPreDetails");
  const secRevWrap = document.getElementById("secRevWrap");
  const secRevDetails = document.getElementById("secRevDetails");
  const secCertWrap = document.getElementById("secCertWrap");
  const secCertDetails = document.getElementById("secCertDetails");
  const secPersWrap = document.getElementById("secPersWrap");
  const secPersDetails = document.getElementById("secPersDetails");

  const secImplDetails = document.getElementById("secImplDetails");
  const secPostWrap = document.getElementById("secPostWrap");
  const secPostDetails = document.getElementById("secPostDetails");

  const donePre = document.getElementById("donePre");
  const doneRev = document.getElementById("doneRev");
  const doneCert = document.getElementById("doneCert");
  const donePers = document.getElementById("donePers");
  const doneImpl = document.getElementById("doneImpl");
  const donePost = document.getElementById("donePost");
  const doneFinal = document.getElementById("doneFinal");

  const btnReset = document.getElementById("btnReset");
  const btnSubmit = form?.querySelector('button[type="submit"]');
  const btnPdf = document.getElementById("btnPdf");
  const finalNote = document.getElementById("finalNote");

  if (btnPdf) btnPdf.disabled = true;

  const preBlocks = [
    document.getElementById("pre1"),
    document.getElementById("pre2"),
    document.getElementById("pre3"),
    document.getElementById("pre4"),
  ].filter(Boolean);

  const pre_aut_coord = document.getElementById("pre_aut_coord");
  const pre_aut_banco = document.getElementById("pre_aut_banco");
  const pre_matriz_infra = document.getElementById("pre_matriz_infra");

  const cmdbReasonWrap = document.getElementById("cmdbReasonWrap");
  const cmdbReason = document.getElementById("cmdbReason");
  const btnCmdbNext = document.getElementById("btnCmdbNext");

  const revBlocks = Array.from(secRevDetails?.querySelectorAll(".q-block") || []);
  const certBlocks = Array.from(secCertDetails?.querySelectorAll(".q-block") || []);
  const persBlocks = Array.from(secPersDetails?.querySelectorAll(".q-block") || []);
  const implBlocks = Array.from(secImplDetails?.querySelectorAll(".q-block") || []);
  const postBlocks = Array.from(secPostDetails?.querySelectorAll(".q-block") || []);

  // =========================
  // Certificación pregunta 9
  // =========================
  const cert9 = document.getElementById("cert9");
  const cert9Done = document.getElementById("cert9_done");
  const cert9Radios = Array.from(document.querySelectorAll('input[name="cert9_resp"]'));
  const cert9SiWrap = document.getElementById("cert9_si_wrap");
  const cert9Detalle = document.getElementById("cert9_detalle");
  const btnCert9Next = document.getElementById("btnCert9Next");

  // =========================
  // Personal asignado dinámico
  // =========================
  const pers4 = document.getElementById("pers4");
  const pers4Done = document.getElementById("pers4_done");
  const certificacionRows = document.getElementById("certificacionRows");
  const btnPers4Next = document.getElementById("btnPers4Next");

  const pers5 = document.getElementById("pers5");
  const pers5Done = document.getElementById("pers5_done");
  const ejecucionRows = document.getElementById("ejecucionRows");
  const btnPers5Next = document.getElementById("btnPers5Next");

  // =========================
  // Post implementación
  // =========================
  const post2 = document.getElementById("post2");
  const post2Done = document.getElementById("post2_done");
  const post2Radios = Array.from(document.querySelectorAll('input[name="post2_resp"]'));
  const post2Extra = document.getElementById("post2_extra");
  const post2Comment = document.getElementById("post2_comment");
  const btnPost2Finish = document.getElementById("btnPost2Finish");

  const post3 = document.getElementById("post3");
  const post3Done = document.getElementById("post3_done");
  const post3SiWrap = document.getElementById("post3_si");
  const post3NoWrap = document.getElementById("post3_no");
  const post3Detalle = document.getElementById("post3_detalle");
  const post3Just = document.getElementById("post3_justificacion");
  const btnPost3Si = document.getElementById("btnPost3Si");
  const btnPost3No = document.getElementById("btnPost3No");

  let selectedMode = null;
  let terminatedByNo = false;
  let formCompletedAt = "";

  const show = (el) => el && el.classList.remove("hidden");
  const hide = (el) => el && el.classList.add("hidden");

  function getEcuadorDateTime() {
    const now = new Date();

    const fecha = new Intl.DateTimeFormat("es-EC", {
      timeZone: "America/Guayaquil",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(now);

    const hora = new Intl.DateTimeFormat("es-EC", {
      timeZone: "America/Guayaquil",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false,
    }).format(now);

    return `${fecha} ${hora}`;
  }

  function markFormCompleted() {
    if (!formCompletedAt) {
      formCompletedAt = getEcuadorDateTime();
    }
  }

  function clearFormCompleted() {
    formCompletedAt = "";
  }

  function refreshFinalMessage() {
    const ok = isDatosGeneralesValidos() && isFlujoCompleto();

    if (ok) {
      show(finalNote);
      show(doneFinal);
    } else {
      hide(finalNote);
      hide(doneFinal);
    }
  }

  function getCheckbox(block) {
    return block?.querySelector('input[type="checkbox"]') || null;
  }

  function getSelectedValueByName(name) {
    const checked = document.querySelector(`input[name="${name}"]:checked`);
    return checked ? checked.value : "";
  }

  function getYesNoFromBlock(block) {
    const radio = block?.querySelector('input[type="radio"]:checked');
    return radio ? radio.value : "";
  }

  function initYesNoBlocks() {
    const ynBlocks = document.querySelectorAll(".q-yn");

    ynBlocks.forEach((block) => {
      const radios = block.querySelectorAll(".yn-radio");
      const hiddenCheck = block.querySelector(".yn-done");

      radios.forEach((radio) => {
        radio.addEventListener("change", () => {
          const checked = block.querySelector(".yn-radio:checked");
          if (hiddenCheck) {
            hiddenCheck.checked = !!checked;
            hiddenCheck.dispatchEvent(new Event("change", { bubbles: true }));
          }
        });
      });
    });
  }

  // =========================
  // Repetibles
  // =========================
  function getRepeatableRows(container) {
    if (!container) return [];
    return Array.from(container.querySelectorAll(".row-grid.row"));
  }

  function getAddRowButton(container) {
    return container?.querySelector(".add-row-btn") || null;
  }

  function createRepeatableRow(type) {
    const row = document.createElement("div");
    row.className = "row-grid row";

    if (type === "cert") {
      row.innerHTML = `
        <input type="text" name="cert_grupo[]" placeholder="Ej: TEC_TCS_WINTEL">
        <button type="button" class="btn-danger remove-row" aria-label="Eliminar fila">✕</button>
      `;
    } else {
      row.innerHTML = `
        <input type="text" name="ejec_grupo[]" placeholder="Ej: TCS_DLVY_STORAGE">
        <input type="text" name="ejec_personal[]" placeholder="Ej: Nestor Morales">
        <button type="button" class="btn-danger remove-row" aria-label="Eliminar fila">✕</button>
      `;
    }

    row.querySelectorAll("input").forEach((input) => {
      input.addEventListener("input", refreshSubmitState);
    });

    const removeBtn = row.querySelector(".remove-row");
    removeBtn?.addEventListener("click", () => {
      const container = row.parentElement;
      const rows = getRepeatableRows(container);

      if (rows.length <= 1) {
        row.querySelectorAll("input").forEach((i) => {
          i.value = "";
        });
      } else {
        row.remove();
      }

      if (container?.id === "certificacionRows" && pers4Done) {
        pers4Done.checked = false;
      }
      if (container?.id === "ejecucionRows" && pers5Done) {
        pers5Done.checked = false;
      }

      refreshSubmitState();
    });

    return row;
  }

  function insertNewRepeatableRow(container, type) {
    if (!container) return;
    const addBtn = getAddRowButton(container);
    const newRow = createRepeatableRow(type);

    if (addBtn) {
      container.insertBefore(newRow, addBtn);
    } else {
      container.appendChild(newRow);
    }
  }

  function setupRepeatableContainer(container, type) {
    if (!container) return;

    const addBtn = getAddRowButton(container);

    getRepeatableRows(container).forEach((row) => {
      row.querySelectorAll("input").forEach((input) => {
        input.addEventListener("input", refreshSubmitState);
      });

      const removeBtn = row.querySelector(".remove-row");
      removeBtn?.addEventListener("click", () => {
        const rows = getRepeatableRows(container);

        if (rows.length <= 1) {
          row.querySelectorAll("input").forEach((i) => {
            i.value = "";
          });
        } else {
          row.remove();
        }

        if (container.id === "certificacionRows" && pers4Done) {
          pers4Done.checked = false;
        }
        if (container.id === "ejecucionRows" && pers5Done) {
          pers5Done.checked = false;
        }

        refreshSubmitState();
      });
    });

    addBtn?.addEventListener("click", () => {
      insertNewRepeatableRow(container, type);
      if (container.id === "certificacionRows" && pers4Done) pers4Done.checked = false;
      if (container.id === "ejecucionRows" && pers5Done) pers5Done.checked = false;
      refreshSubmitState();
    });
  }

  function clearRepeatableRows(container) {
    if (!container) return;
    const rows = getRepeatableRows(container);

    rows.forEach((row, idx) => {
      if (idx === 0) {
        row.querySelectorAll("input").forEach((i) => {
          i.value = "";
        });
      } else {
        row.remove();
      }
    });
  }

  function getRepeatableData(container, type) {
    const rows = getRepeatableRows(container);

    return rows
      .map((row) => {
        const inputs = row.querySelectorAll("input");

        if (type === "cert") {
          const grupo = (inputs[0]?.value || "").trim();
          return { grupo, type };
        }

        const grupo = (inputs[0]?.value || "").trim();
        const personal = (inputs[1]?.value || "").trim();

        return { grupo, personal, type };
      })
      .filter((item) => {
        if (type === "cert") return item.grupo;
        return item.grupo || item.personal;
      });
  }

  function validateRepeatableBlock(container) {
    const rows = getRepeatableRows(container);
    const isPers4 = container?.id === "certificacionRows";
    let validCount = 0;

    for (const row of rows) {
      const inputs = row.querySelectorAll("input");

      if (isPers4) {
        const grupo = (inputs[0]?.value || "").trim();

        if (!grupo) continue;

        validCount++;
        continue;
      }

      const grupo = (inputs[0]?.value || "").trim();
      const personal = (inputs[1]?.value || "").trim();

      if (!grupo && !personal) continue;

      if (!grupo || !personal) {
        return {
          ok: false,
          message: "Complete Grupo resolutor y Personal asignado en todas las filas utilizadas."
        };
      }

      validCount++;
    }

    if (validCount === 0) {
      return {
        ok: false,
        message: isPers4
          ? "Debe completar al menos un Grupo resolutor."
          : "Debe completar al menos una fila."
      };
    }

    return { ok: true };
  }

  function clearAllInputs(block) {
    if (!block) return;

    block.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
      cb.checked = false;
    });

    block.querySelectorAll('input[type="radio"]').forEach((r) => {
      r.checked = false;
    });

    block.querySelectorAll('input[type="text"]').forEach((i) => {
      i.value = "";
    });

    block.querySelectorAll("textarea").forEach((t) => {
      t.value = "";
    });

    block.querySelectorAll("#post3_si, #post3_no, #cert9_si_wrap, #post2_extra").forEach((el) => {
      el.classList.add("hidden");
    });

    if (block.id === "pers4") {
      clearRepeatableRows(certificacionRows);
    }

    if (block.id === "pers5") {
      clearRepeatableRows(ejecucionRows);
    }
  }

  function hideBlockAndClear(block) {
    if (!block) return;
    hide(block);
    clearAllInputs(block);

    if (block.id === "pre4") {
      form.querySelectorAll('input[name="cmdb_aplica"]').forEach((r) => {
        r.checked = false;
      });
      if (cmdbReason) cmdbReason.value = "";
      hide(cmdbReasonWrap);
      if (btnCmdbNext) btnCmdbNext.disabled = true;
    }

    if (block.id === "pers4" && pers4Done) {
      pers4Done.checked = false;
    }

    if (block.id === "pers5" && pers5Done) {
      pers5Done.checked = false;
    }

    if (block.id === "post2") {
      if (post2Comment) post2Comment.value = "";
      hide(post2Extra);
      if (post2Done) post2Done.checked = false;
    }

    if (block.id === "post3") {
      if (post3Detalle) post3Detalle.value = "";
      if (post3Just) post3Just.value = "";
      hide(post3SiWrap);
      hide(post3NoWrap);
      if (post3Done) post3Done.checked = false;
    }
  }

  function resetBlocks(list) {
    list.forEach((b) => hideBlockAndClear(b));
  }

  function showFirst(list) {
    if (list[0]) show(list[0]);
  }

  function showNext(list, currentBlock) {
    const idx = list.indexOf(currentBlock);
    if (idx < 0) return;
    const next = list[idx + 1];
    if (next) show(next);
  }

  function clearFrom(list, startIndex) {
    for (let i = startIndex; i < list.length; i++) {
      hideBlockAndClear(list[i]);
    }
  }

  function getCmdbValue() {
    const r = form.querySelector('input[name="cmdb_aplica"]:checked');
    return r ? r.value : "";
  }

  function cmdbOk() {
    const v = getCmdbValue();
    if (v === "SI") return true;
    if (v === "NA") return !!cmdbReason?.value.trim();
    return false;
  }

  function requiereNumeroCambio() {
    return true;
  }

  function setNumeroCambioVisibility() {
    if (!numeroCambioWrap) return;
    show(numeroCambioWrap);
    if (numeroCambioInput) numeroCambioInput.required = true;
  }

  function isDatosGeneralesValidos() {
    const okNombre = !!nombresInput?.value.trim();
    const okQ0 = !!q0?.checked;
    const okChg = !requiereNumeroCambio() || !!numeroCambioInput?.value.trim();
    return okNombre && okQ0 && okChg;
  }

  function allCheckedByBlocks(blocks, opts = {}) {
    for (const b of blocks) {
      if (!b) continue;

      if (opts.cmdbInPre && b.id === "pre4") {
        if (!cmdbOk()) return false;
        continue;
      }

      const cb = getCheckbox(b);
      if (!cb || !cb.checked) return false;
    }
    return true;
  }

  function isFlujoCompleto() {
    if (!selectedMode) return false;
    if (terminatedByNo) return true;

    if (selectedMode === "CERTIFICACION") {
      return (
        allCheckedByBlocks(preBlocks, { cmdbInPre: true }) &&
        allCheckedByBlocks(revBlocks) &&
        allCheckedByBlocks(certBlocks) &&
        allCheckedByBlocks(persBlocks)
      );
    }

    return allCheckedByBlocks(implBlocks) && allCheckedByBlocks(postBlocks);
  }

  function refreshSubmitState() {
    const ok = isDatosGeneralesValidos() && isFlujoCompleto();

    if (ok) {
      markFormCompleted();
    }

    if (btnSubmit) btnSubmit.disabled = !ok;
    if (btnPdf) btnPdf.disabled = !ok;

    refreshFinalMessage();
  }

  function closeAllOpenFlowSections() {
    if (secPreDetails) secPreDetails.open = false;
    if (secRevDetails) secRevDetails.open = false;
    if (secCertDetails) secCertDetails.open = false;
    if (secPersDetails) secPersDetails.open = false;
    if (secImplDetails) secImplDetails.open = false;
    if (secPostDetails) secPostDetails.open = false;
  }

  function stopFlowFromBlock(list, block) {
    const idx = list.indexOf(block);
    if (idx >= 0) {
      clearFrom(list, idx + 1);
    }
  }

  function terminateFlowBecauseNo(originBlock = null) {
    terminatedByNo = true;

    if (selectedMode === "CERTIFICACION") {
      if (originBlock) {
        if (preBlocks.includes(originBlock)) {
          stopFlowFromBlock(preBlocks, originBlock);
          hide(secRevWrap);
          hide(secCertWrap);
          hide(secPersWrap);
          resetBlocks(revBlocks);
          resetBlocks(certBlocks);
          resetBlocks(persBlocks);
          hide(donePre);
          hide(doneRev);
          hide(doneCert);
          hide(donePers);
        } else if (revBlocks.includes(originBlock)) {
          stopFlowFromBlock(revBlocks, originBlock);
          hide(secCertWrap);
          hide(secPersWrap);
          resetBlocks(certBlocks);
          resetBlocks(persBlocks);
          hide(doneRev);
          hide(doneCert);
          hide(donePers);
        } else if (certBlocks.includes(originBlock)) {
          stopFlowFromBlock(certBlocks, originBlock);
          hide(secPersWrap);
          resetBlocks(persBlocks);
          hide(doneCert);
          hide(donePers);
        } else if (persBlocks.includes(originBlock)) {
          stopFlowFromBlock(persBlocks, originBlock);
          hide(donePers);
        }
      } else {
        hide(secRevWrap);
        hide(secCertWrap);
        hide(secPersWrap);
      }
    }

    if (selectedMode === "IMPLEMENTACION") {
      if (originBlock) {
        if (implBlocks.includes(originBlock)) {
          stopFlowFromBlock(implBlocks, originBlock);
          hide(secPostWrap);
          resetBlocks(postBlocks);
          hide(doneImpl);
          hide(donePost);
        } else if (postBlocks.includes(originBlock)) {
          stopFlowFromBlock(postBlocks, originBlock);
          hide(donePost);
        }
      } else {
        hide(secPostWrap);
      }
    }

    closeAllOpenFlowSections();
    refreshSubmitState();
  }

  function resetTermination() {
    terminatedByNo = false;
  }

  function resetAllToGate() {
    resetTermination();
    clearFormCompleted();
    hide(finalNote);
    hide(doneFinal);

    hide(formWrap);
    show(modeGate);
    hide(btnCambiarTipo);

    modeOptions?.querySelectorAll(".mode-card").forEach((b) => {
      b.setAttribute("aria-pressed", "false");
    });

    selectedMode = null;
    if (selectedModeInput) selectedModeInput.value = "";

    hide(certFlow);
    hide(implFlow);

    form?.reset();

    resetBlocks(preBlocks);
    resetBlocks(revBlocks);
    resetBlocks(certBlocks);
    resetBlocks(persBlocks);
    resetBlocks(implBlocks);
    resetBlocks(postBlocks);

    hide(secRevWrap);
    hide(secCertWrap);
    hide(secPersWrap);
    hide(secPostWrap);

    if (secInfoDetails) secInfoDetails.open = true;
    if (secPreDetails) secPreDetails.open = false;
    if (secRevDetails) secRevDetails.open = false;
    if (secCertDetails) secCertDetails.open = false;
    if (secPersDetails) secPersDetails.open = false;
    if (secImplDetails) secImplDetails.open = false;
    if (secPostDetails) secPostDetails.open = false;

    hide(doneInfo);
    hide(donePre);
    hide(doneRev);
    hide(doneCert);
    hide(donePers);
    hide(doneImpl);
    hide(donePost);
    hide(doneFinal);

    hide(cmdbReasonWrap);
    if (btnCmdbNext) btnCmdbNext.disabled = true;

    if (pers4Done) pers4Done.checked = false;
    if (pers5Done) pers5Done.checked = false;
    clearRepeatableRows(certificacionRows);
    clearRepeatableRows(ejecucionRows);

    if (post2Comment) post2Comment.value = "";
    hide(post2Extra);

    setNumeroCambioVisibility();
    if (btnIniciar) btnIniciar.disabled = true;
    refreshSubmitState();
  }

  function openFirstFlowSection() {
    if (selectedMode === "CERTIFICACION") {
      if (secInfoDetails) secInfoDetails.open = false;
      if (secPreDetails) secPreDetails.open = true;
      showFirst(preBlocks);
      secPreDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    } else if (selectedMode === "IMPLEMENTACION") {
      if (secInfoDetails) secInfoDetails.open = false;
      if (secImplDetails) secImplDetails.open = true;
      showFirst(implBlocks);
      secImplDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  initYesNoBlocks();
  setupRepeatableContainer(certificacionRows, "cert");
  setupRepeatableContainer(ejecucionRows, "ejec");
  resetAllToGate();

  modeOptions?.addEventListener("click", (e) => {
    const btn = e.target.closest(".mode-card");
    if (!btn) return;

    modeOptions.querySelectorAll(".mode-card").forEach((b) => {
      b.setAttribute("aria-pressed", "false");
    });

    btn.setAttribute("aria-pressed", "true");

    selectedMode = btn.dataset.mode;
    if (selectedModeInput) selectedModeInput.value = selectedMode;

    setNumeroCambioVisibility();
    if (btnIniciar) btnIniciar.disabled = false;
    refreshSubmitState();
  });

  btnIniciar?.addEventListener("click", () => {
    if (!selectedMode) {
      alert("Seleccione Certificación o Implementación.");
      return;
    }

    resetTermination();
    clearFormCompleted();
    hide(finalNote);
    hide(doneFinal);

    hide(modeGate);
    show(formWrap);
    show(btnCambiarTipo);

    resetBlocks(preBlocks);
    resetBlocks(revBlocks);
    resetBlocks(certBlocks);
    resetBlocks(persBlocks);
    resetBlocks(implBlocks);
    resetBlocks(postBlocks);

    hide(secRevWrap);
    hide(secCertWrap);
    hide(secPersWrap);
    hide(secPostWrap);

    hide(donePre);
    hide(doneRev);
    hide(doneCert);
    hide(donePers);
    hide(doneImpl);
    hide(donePost);
    hide(doneFinal);

    hide(cmdbReasonWrap);
    if (btnCmdbNext) btnCmdbNext.disabled = true;

    if (pers4Done) pers4Done.checked = false;
    if (pers5Done) pers5Done.checked = false;
    clearRepeatableRows(certificacionRows);
    clearRepeatableRows(ejecucionRows);

    if (post2Comment) post2Comment.value = "";
    hide(post2Extra);

    setNumeroCambioVisibility();

    if (selectedMode === "CERTIFICACION") {
      show(certFlow);
      hide(implFlow);
      if (secPreDetails) secPreDetails.open = false;
    } else {
      show(implFlow);
      hide(certFlow);
      if (secImplDetails) secImplDetails.open = false;
    }

    if (secInfoDetails) secInfoDetails.open = true;
    secInfoDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    refreshSubmitState();
  });

  btnCambiarTipo?.addEventListener("click", resetAllToGate);

  q0?.addEventListener("change", () => {
    const checked = !!q0.checked;

    if (!checked) {
      clearFormCompleted();
      hide(finalNote);
      hide(doneFinal);
      resetTermination();

      hide(doneInfo);

      resetBlocks(preBlocks);
      resetBlocks(revBlocks);
      resetBlocks(certBlocks);
      resetBlocks(persBlocks);
      resetBlocks(implBlocks);
      resetBlocks(postBlocks);

      hide(secRevWrap);
      hide(secCertWrap);
      hide(secPersWrap);
      hide(secPostWrap);

      hide(donePre);
      hide(doneRev);
      hide(doneCert);
      hide(donePers);
      hide(doneImpl);
      hide(donePost);
      hide(doneFinal);

      if (secInfoDetails) secInfoDetails.open = true;
      if (secPreDetails) secPreDetails.open = false;
      if (secRevDetails) secRevDetails.open = false;
      if (secCertDetails) secCertDetails.open = false;
      if (secPersDetails) secPersDetails.open = false;
      if (secImplDetails) secImplDetails.open = false;
      if (secPostDetails) secPostDetails.open = false;

      hide(cmdbReasonWrap);
      if (btnCmdbNext) btnCmdbNext.disabled = true;

      if (pers4Done) pers4Done.checked = false;
      if (pers5Done) pers5Done.checked = false;
      clearRepeatableRows(certificacionRows);
      clearRepeatableRows(ejecucionRows);

      if (post2Comment) post2Comment.value = "";
      hide(post2Extra);

      refreshSubmitState();
      return;
    }

    if (!nombresInput?.value.trim()) {
      alert("Complete Nombres completos antes de continuar.");
      q0.checked = false;
      form.querySelectorAll('input[name="infoConfirm"]').forEach((r) => {
        r.checked = false;
      });
      refreshSubmitState();
      return;
    }

    if (requiereNumeroCambio() && !numeroCambioInput?.value.trim()) {
      alert("Complete el N° de Control de Cambios antes de continuar.");
      q0.checked = false;
      form.querySelectorAll('input[name="infoConfirm"]').forEach((r) => {
        r.checked = false;
      });
      refreshSubmitState();
      return;
    }

    show(doneInfo);

    const infoResp = getSelectedValueByName("infoConfirm");
    if (secInfoDetails) secInfoDetails.open = false;

    if (infoResp === "NO") {
      terminateFlowBecauseNo();
      return;
    }

    openFirstFlowSection();
    refreshSubmitState();
  });

  ["nombresCompletos", "numeroCambio"].forEach((id) => {
    document.getElementById(id)?.addEventListener("input", () => {
      refreshSubmitState();
    });
  });

  form?.addEventListener("change", (e) => {
    const t = e.target;

    if (t && t.name === "cmdb_aplica") {
      const val = getCmdbValue();
      const pre4 = document.getElementById("pre4");

      if (val === "NA") {
        show(cmdbReasonWrap);
        if (btnCmdbNext) btnCmdbNext.disabled = !(cmdbReason?.value.trim());
      } else if (val === "SI") {
        hide(cmdbReasonWrap);
        if (cmdbReason) cmdbReason.value = "";
        if (btnCmdbNext) btnCmdbNext.disabled = true;

        if (pre4 && !pre4.classList.contains("hidden")) {
          showNext(preBlocks, pre4);
        }
      } else {
        hide(cmdbReasonWrap);
        if (btnCmdbNext) btnCmdbNext.disabled = true;
      }

      if (pre4 && !pre4.classList.contains("hidden")) {
        const idx = preBlocks.indexOf(pre4);
        if (!cmdbOk()) {
          clearFrom(preBlocks, idx + 1);
          hide(donePre);
          hide(secRevWrap);
          hide(secCertWrap);
          hide(secPersWrap);
          hide(doneRev);
          hide(doneCert);
          hide(donePers);
        }
      }

      refreshSubmitState();
    }
  });

  cmdbReason?.addEventListener("input", () => {
    if (getCmdbValue() === "NA" && btnCmdbNext) {
      btnCmdbNext.disabled = !(cmdbReason?.value.trim());
    }
    refreshSubmitState();
  });

  btnCmdbNext?.addEventListener("click", () => {
    if (getCmdbValue() !== "NA") return;

    if (!cmdbReason?.value.trim()) {
      alert("Especifique el motivo para N/A antes de continuar.");
      cmdbReason?.focus();
      return;
    }

    const pre4 = document.getElementById("pre4");
    if (!pre4) return;

    showNext(preBlocks, pre4);
    refreshPreCompletionGate();

    const idx = preBlocks.indexOf(pre4);
    const next = preBlocks[idx + 1];
    (next || secRevDetails)?.scrollIntoView({ behavior: "smooth", block: "start" });

    refreshSubmitState();
  });

  [pre_aut_coord, pre_aut_banco, pre_matriz_infra].forEach((el) => {
    el?.addEventListener("change", () => {
      if (selectedMode !== "CERTIFICACION" || !q0.checked) return;

      const block = el.closest(".q-block");
      if (!block) return;

      const value = getYesNoFromBlock(block);

      if (value === "NO") {
        if (secPreDetails) secPreDetails.open = false;
        terminateFlowBecauseNo(block);
        return;
      }

      if (!el.checked) {
        const idx = preBlocks.indexOf(block);
        clearFrom(preBlocks, idx + 1);

        hide(donePre);
        hide(secRevWrap);
        hide(secCertWrap);
        hide(secPersWrap);
        hide(doneRev);
        hide(doneCert);
        hide(donePers);
        hide(doneFinal);

        if (secPreDetails) secPreDetails.open = true;
        if (secRevDetails) secRevDetails.open = false;
        if (secCertDetails) secCertDetails.open = false;
        if (secPersDetails) secPersDetails.open = false;

        refreshSubmitState();
        return;
      }

      showNext(preBlocks, block);
      refreshSubmitState();
    });
  });

  function wireSequentialList(blocks, onComplete, onUncomplete) {
    blocks.forEach((block) => {
      const cb = getCheckbox(block);
      if (!cb) return;

      cb.addEventListener("change", () => {
        if (!q0.checked) {
          cb.checked = false;
          return;
        }

        const value = getYesNoFromBlock(block);

        if (value === "NO") {
          terminateFlowBecauseNo(block);
          return;
        }

        const idx = blocks.indexOf(block);

        if (!cb.checked) {
          clearFrom(blocks, idx + 1);
          onUncomplete?.(block, cb);
          refreshSubmitState();
          return;
        }

        showNext(blocks, block);

        if (allCheckedByBlocks(blocks)) {
          onComplete?.(block, cb);
        }

        refreshSubmitState();
      });
    });
  }

  function tryOpenRevisionIfPreComplete() {
    if (!allCheckedByBlocks(preBlocks, { cmdbInPre: true })) return false;

    show(donePre);
    if (secPreDetails) secPreDetails.open = false;

    show(secRevWrap);
    if (secRevDetails) secRevDetails.open = true;

    resetBlocks(revBlocks);
    showFirst(revBlocks);

    secRevDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    return true;
  }

  function refreshPreCompletionGate() {
    if (selectedMode !== "CERTIFICACION" || !q0.checked || terminatedByNo) return;

    const p1 = pre_aut_coord?.checked;
    const p2 = pre_aut_banco?.checked;
    const p3 = pre_matriz_infra?.checked;

    if (p1 && p2 && p3 && cmdbOk()) {
      tryOpenRevisionIfPreComplete();
    } else {
      hide(donePre);
      hide(secRevWrap);
      hide(secCertWrap);
      hide(secPersWrap);
      hide(doneRev);
      hide(doneCert);
      hide(donePers);
      hide(doneFinal);

      resetBlocks(revBlocks);
      resetBlocks(certBlocks);
      resetBlocks(persBlocks);

      if (secRevDetails) secRevDetails.open = false;
      if (secCertDetails) secCertDetails.open = false;
      if (secPersDetails) secPersDetails.open = false;
      if (secPreDetails) secPreDetails.open = true;
    }
  }

  form.addEventListener("change", (e) => {
    const t = e.target;
    if (
      t?.id === "pre_aut_coord" ||
      t?.id === "pre_aut_banco" ||
      t?.id === "pre_matriz_infra" ||
      t?.name === "cmdb_aplica"
    ) {
      refreshPreCompletionGate();
    }
  });

  wireSequentialList(
    revBlocks,
    () => {
      show(doneRev);
      if (secRevDetails) secRevDetails.open = false;

      show(secCertWrap);
      if (secCertDetails) secCertDetails.open = true;

      resetBlocks(certBlocks);
      showFirst(certBlocks);

      secCertDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    },
    () => {
      hide(doneRev);
      hide(secCertWrap);
      hide(secPersWrap);
      hide(doneCert);
      hide(donePers);
      hide(doneFinal);

      resetBlocks(certBlocks);
      resetBlocks(persBlocks);

      if (secCertDetails) secCertDetails.open = false;
      if (secPersDetails) secPersDetails.open = false;
      if (secRevDetails) secRevDetails.open = true;
    }
  );

  wireSequentialList(
    certBlocks,
    () => {
      show(doneCert);
      if (secCertDetails) secCertDetails.open = false;

      show(secPersWrap);
      if (secPersDetails) secPersDetails.open = true;

      resetBlocks(persBlocks);
      showFirst(persBlocks);

      secPersDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    },
    (block, cb) => {
      hide(doneCert);
      hide(secPersWrap);
      hide(donePers);
      hide(doneFinal);

      resetBlocks(persBlocks);

      if (cb?.id === "cert9_done") {
        resetCert9ConditionalUI();
      }

      if (secPersDetails) secPersDetails.open = false;
      if (secCertDetails) secCertDetails.open = true;
    }
  );

  wireSequentialList(
    persBlocks,
    () => {
      show(donePers);
      if (secPersDetails) secPersDetails.open = false;
    },
    (block, cb) => {
      hide(donePers);
      hide(doneFinal);

      if (cb?.id === "pers4_done") {
        clearRepeatableRows(certificacionRows);
      }

      if (cb?.id === "pers5_done") {
        clearRepeatableRows(ejecucionRows);
      }

      if (secPersDetails) secPersDetails.open = true;
    }
  );

  wireSequentialList(
    implBlocks,
    () => {
      show(doneImpl);
      if (secImplDetails) secImplDetails.open = false;

      show(secPostWrap);
      if (secPostDetails) secPostDetails.open = true;

      resetBlocks(postBlocks);
      showFirst(postBlocks);

      secPostDetails?.scrollIntoView({ behavior: "smooth", block: "start" });
    },
    () => {
      hide(doneImpl);
      hide(secPostWrap);
      hide(donePost);
      hide(doneFinal);

      resetBlocks(postBlocks);

      if (secPostDetails) secPostDetails.open = false;
      if (secImplDetails) secImplDetails.open = true;
    }
  );

  // =========================
  // Certificación pregunta 9
  // =========================
  cert9Radios.forEach((r) => {
    r.addEventListener("change", () => {
      if (selectedMode !== "CERTIFICACION") return;
      if (!q0.checked) return;

      const val = r.value;

      if (val === "SI") {
        if (cert9Done) cert9Done.checked = false;
        show(cert9SiWrap);
        refreshSubmitState();
        return;
      }

      hide(cert9SiWrap);
      if (cert9Detalle) cert9Detalle.value = "";

      if (cert9Done) {
        cert9Done.checked = true;
        cert9Done.dispatchEvent(new Event("change", { bubbles: true }));
      }

      refreshSubmitState();
    });
  });

  btnCert9Next?.addEventListener("click", () => {
    if (selectedMode !== "CERTIFICACION") return;
    if (!q0.checked) return;

    if (getCert9Value() !== "SI") return;

    const txt = (cert9Detalle?.value || "").trim();
    if (!txt) {
      alert("Por favor detalle el proceso automático antes de continuar.");
      cert9Detalle?.focus();
      return;
    }

    if (cert9Done) {
      cert9Done.checked = true;
      cert9Done.dispatchEvent(new Event("change", { bubbles: true }));
    }

    refreshSubmitState();
  });

  function resetCert9ConditionalUI() {
    cert9Radios.forEach((r) => {
      r.checked = false;
    });

    if (cert9Detalle) cert9Detalle.value = "";
    if (cert9Done) cert9Done.checked = false;
    hide(cert9SiWrap);
  }

  function getCert9Value() {
    return document.querySelector('input[name="cert9_resp"]:checked')?.value || "";
  }

  // =========================
  // Personal asignado dinámico
  // =========================
  btnPers4Next?.addEventListener("click", () => {
    if (selectedMode !== "CERTIFICACION") return;
    if (!q0.checked) return;

    const validation = validateRepeatableBlock(certificacionRows);
    if (!validation.ok) {
      alert(validation.message);
      return;
    }

    if (pers4Done) {
      pers4Done.checked = true;
      pers4Done.dispatchEvent(new Event("change", { bubbles: true }));
    }

    refreshSubmitState();
  });

  btnPers5Next?.addEventListener("click", () => {
    if (selectedMode !== "CERTIFICACION") return;
    if (!q0.checked) return;

    const validation = validateRepeatableBlock(ejecucionRows);
    if (!validation.ok) {
      alert(validation.message);
      return;
    }

    if (pers5Done) {
      pers5Done.checked = true;
      pers5Done.dispatchEvent(new Event("change", { bubbles: true }));
    }

    refreshSubmitState();
  });

  function resetPostConditionalUI() {
    post2Radios.forEach((r) => {
      r.checked = false;
    });

    if (post2Comment) post2Comment.value = "";
    hide(post2Extra);

    if (post3Detalle) post3Detalle.value = "";
    if (post3Just) post3Just.value = "";
    hide(post3SiWrap);
    hide(post3NoWrap);

    if (post2Done) post2Done.checked = false;
    if (post3Done) post3Done.checked = false;
  }

  function showPost3Branch(val) {
    hide(post3SiWrap);
    hide(post3NoWrap);

    if (post3Detalle) post3Detalle.value = "";
    if (post3Just) post3Just.value = "";
    if (post3Done) post3Done.checked = false;

    if (val === "SI") show(post3SiWrap);
    if (val === "NO") show(post3NoWrap);
  }

  // =========================
  // POST2: SI sigue, NO pide comentario y finaliza
  // =========================
  post2Radios.forEach((r) => {
    r.addEventListener("change", () => {
      if (selectedMode !== "IMPLEMENTACION") return;
      if (!q0.checked) return;

      const val = r.value;

      if (post2Done) {
        post2Done.checked = false;
      }

      const idxPost2 = postBlocks.indexOf(post2);
      if (idxPost2 >= 0) {
        clearFrom(postBlocks, idxPost2 + 1);
      }

      if (val === "NO") {
        show(post2Extra);
        if (post3Done) post3Done.checked = false;
        refreshSubmitState();
        return;
      }

      hide(post2Extra);
      if (post2Comment) post2Comment.value = "";

      if (post2Done) {
        post2Done.checked = true;
        post2Done.dispatchEvent(new Event("change", { bubbles: true }));
      }

      show(post3);
      showPost3Branch(val);

      refreshSubmitState();
    });
  });

  btnPost2Finish?.addEventListener("click", () => {
    if (selectedMode !== "IMPLEMENTACION") return;
    if (!q0.checked) return;

    const selected = document.querySelector('input[name="post2_resp"]:checked')?.value || "";
    if (selected !== "NO") return;

    const txt = (post2Comment?.value || "").trim();
    if (!txt) {
      alert("Debe escribir el motivo antes de finalizar.");
      post2Comment?.focus();
      return;
    }

    if (post2Done) {
      post2Done.checked = true;
    }

    if (secPostDetails) secPostDetails.open = false;
    terminateFlowBecauseNo(post2);
    refreshSubmitState();
  });

  function continuePost3(isSi) {
    if (selectedMode !== "IMPLEMENTACION") return;
    if (!q0.checked) return;

    const txt = isSi ? (post3Detalle?.value || "").trim() : (post3Just?.value || "").trim();
    if (!txt) {
      alert("Por favor complete el campo antes de continuar.");
      (isSi ? post3Detalle : post3Just)?.focus();
      return;
    }

    if (post3Done) {
      post3Done.checked = true;
      post3Done.dispatchEvent(new Event("change", { bubbles: true }));
    }

    refreshSubmitState();
  }

  btnPost3Si?.addEventListener("click", () => continuePost3(true));
  btnPost3No?.addEventListener("click", () => continuePost3(false));

  wireSequentialList(
    postBlocks,
    () => {
      show(donePost);
      if (secPostDetails) secPostDetails.open = false;
    },
    (block, cb) => {
      hide(donePost);
      hide(doneFinal);

      if (cb?.id === "post2_done" || cb?.id === "post3_done") {
        resetPostConditionalUI();
      }

      if (secPostDetails) secPostDetails.open = true;
      refreshSubmitState();
    }
  );

  btnReset?.addEventListener("click", () => {
    setTimeout(() => resetAllToGate(), 0);
  });

  async function sendToPowerAutomate(payload) {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${txt}`);
    }

    return await res.text().catch(() => "");
  }

  function buildPayload() {
    const post2Resp = document.querySelector('input[name="post2_resp"]:checked')?.value || "";

    return {
      selectedMode: selectedMode || "",
      nombresCompletos: (nombresInput?.value || "").trim(),
      numeroCambio: (numeroCambioInput?.value || "").trim(),
      fechaRealizacion: formCompletedAt || getEcuadorDateTime(),
      infoConfirm: getSelectedValueByName("infoConfirm"),

      pre_aut_coord: getSelectedValueByName("pre_aut_coord_resp"),
      pre_aut_banco: getSelectedValueByName("pre_aut_banco_resp"),
      pre_matriz_infra: getSelectedValueByName("pre_matriz_infra_resp"),

      cmdb_aplica: getCmdbValue(),
      cmdb_reason: (cmdbReason?.value || "").trim(),

      rev: revBlocks.map((b) => getYesNoFromBlock(b)),
      cert: certBlocks.map((b) => {
        if (b?.id === "cert9") return getCert9Value();
        return getYesNoFromBlock(b);
      }),
      cert9_resp: getCert9Value(),
      cert9_detalle: (cert9Detalle?.value || "").trim(),

      pers: persBlocks.map((b) => {
        if (b?.id === "pers4") return pers4Done?.checked ? "COMPLETADO" : "";
        if (b?.id === "pers5") return pers5Done?.checked ? "COMPLETADO" : "";
        return getYesNoFromBlock(b);
      }),
      pers4_detalle: getRepeatableData(certificacionRows, "cert"),
      pers5_detalle: getRepeatableData(ejecucionRows, "ejec"),

      impl: implBlocks.map((b) => getYesNoFromBlock(b)),

      post: postBlocks.map((b) => getYesNoFromBlock(b)),
      post2_resp: post2Resp,
      post2_comment: (post2Comment?.value || "").trim(),
      post3_detalle: (post3Detalle?.value || "").trim(),
      post3_justificacion: (post3Just?.value || "").trim(),

      terminatedByNo,
      formVersion: "TCS-SI-NO-STOP-ON-NO",
      source: window.location.href,
    };
  }

  form?.addEventListener("submit", async (e) => {
    e.preventDefault();

    if (btnSubmit?.disabled) {
      alert("Complete el formulario para habilitar ENVIAR.");
      return;
    }

    const payload = buildPayload();

    const prevHtml = btnSubmit?.innerHTML;
    if (btnSubmit) btnSubmit.disabled = true;
    if (btnSubmit) btnSubmit.innerHTML = "<span>ENVIANDO...</span>";

    try {
      await sendToPowerAutomate(payload);
      alert("Datos enviados correctamente");
    } catch (err) {
      console.error(err);
      alert("No se pudo guardar. Revisa el Flow y permisos.");
    } finally {
      if (btnSubmit) btnSubmit.innerHTML = prevHtml || "<span>ENVIAR</span>";
      refreshSubmitState();
    }
  });

  function safeText(v) {
    return (v || "").toString().replace(/\s+/g, " ").trim();
  }

  function getQuestionText(block) {
    if (!block) return "";
    const qText = block.querySelector(".q-text");
    if (qText) return safeText(qText.textContent);

    const label = block.querySelector("label");
    if (!label) return "";
    const clone = label.cloneNode(true);
    clone.querySelectorAll("input, button, textarea, .repeatable").forEach((i) => i.remove());
    return safeText(clone.textContent);
  }

  function rowForYesNoBlock(block) {
    const txt = getQuestionText(block);
    const val = getYesNoFromBlock(block);
    return [txt, val || ""];
  }

  function hasRows(rows) {
    return Array.isArray(rows) && rows.length > 0;
  }

  function buildPdfData() {
    const nombre = safeText(nombresInput?.value);
    const numeroCambio = safeText(numeroCambioInput?.value);
    const fechaRealizacion = formCompletedAt || getEcuadorDateTime();

    const infoGeneral = [[
      "Confirmo que el nombre y número de control de cambio son correctos",
      getSelectedValueByName("infoConfirm")
    ]];

    const preRows = [];
    const pre1 = document.getElementById("pre1");
    const pre2 = document.getElementById("pre2");
    const pre3 = document.getElementById("pre3");

    if (pre1 && getYesNoFromBlock(pre1)) preRows.push(rowForYesNoBlock(pre1));
    if (pre2 && getYesNoFromBlock(pre2)) preRows.push(rowForYesNoBlock(pre2));
    if (pre3 && getYesNoFromBlock(pre3)) preRows.push(rowForYesNoBlock(pre3));

    const cmdbVal = getCmdbValue();
    const cmdbLabel = "4. ¿Aplica Matriz de CMDB?";
    if (cmdbVal === "SI") {
      preRows.push([cmdbLabel, "SI"]);
    } else if (cmdbVal === "NA") {
      const reason = safeText(cmdbReason?.value);
      preRows.push([cmdbLabel, "N/A"]);
      if (reason) preRows.push([`Motivo N/A: ${reason}`, ""]);
    }

    const revRows = revBlocks
      .map((b) => rowForYesNoBlock(b))
      .filter((r) => r[0] && r[1]);

    const certRows = [];

    certBlocks.forEach((b) => {
      if (!b) return;

      if (b.id === "cert9") {
        const val = getCert9Value();
        if (val) {
          certRows.push([
            "9. ¿Dentro de las actividades a implementarse, se empleará un proceso automático?",
            val
          ]);

          const detalle = safeText(cert9Detalle?.value);
          if (val === "SI" && detalle) {
            certRows.push([`Detalle del proceso automático: ${detalle}`, ""]);
          }
        }
        return;
      }

      const row = rowForYesNoBlock(b);
      if (row[0] && row[1]) certRows.push(row);
    });

    const persRows = [];

    persBlocks.forEach((b) => {
      if (!b) return;

      if (b.id === "pers4") {
        if (pers4Done?.checked) {
          persRows.push([
            "4. Detalle nombre del grupo resolutor asignado para Certificación de tareas TCS",
            "COMPLETADO"
          ]);

          getRepeatableData(certificacionRows, "cert").forEach((item, index) => {
            persRows.push([
              `Item ${index + 1}: Grupo resolutor: ${item.grupo}`,
              ""
            ]);
          });
        }
        return;
      }

      if (b.id === "pers5") {
        if (pers5Done?.checked) {
          persRows.push([
            "5. Detalle nombre del grupo resolutor y personal asignado para Ejecución de tareas TCS",
            "COMPLETADO"
          ]);

          getRepeatableData(ejecucionRows, "ejec").forEach((item, index) => {
            persRows.push([
              `Item ${index + 1}: Grupo resolutor: ${item.grupo} | Personal asignado: ${item.personal}`,
              ""
            ]);
          });
        }
        return;
      }

      const row = rowForYesNoBlock(b);
      if (row[0] && row[1]) persRows.push(row);
    });

    const implRows = implBlocks
      .map((b) => rowForYesNoBlock(b))
      .filter((r) => r[0] && r[1]);

    const resp = document.querySelector('input[name="post2_resp"]:checked')?.value || "";
    const post2Motivo = safeText(post2Comment?.value);
    const det = safeText(post3Detalle?.value);
    const jus = safeText(post3Just?.value);

    const postRows = [];
    const b1 = document.getElementById("post1");
    const b4 = document.getElementById("post4");
    const b5 = document.getElementById("post5");

    if (b1 && getYesNoFromBlock(b1)) postRows.push(rowForYesNoBlock(b1));
    if (resp) postRows.push(["2. ¿Se realizaron pruebas funcionales como parte del plan de pruebas post implementación?", resp]);
    if (resp === "NO" && post2Motivo) postRows.push([`Motivo de finalización: ${post2Motivo}`, ""]);
    if (resp === "SI" && det) postRows.push([`3. Detalle pruebas: ${det}`, ""]);
    if (resp === "NO" && jus) postRows.push([`3. Justificación: ${jus}`, ""]);
    if (b4 && getYesNoFromBlock(b4)) postRows.push(rowForYesNoBlock(b4));
    if (b5 && getYesNoFromBlock(b5)) postRows.push(rowForYesNoBlock(b5));

    return {
      nombre,
      numeroCambio,
      fechaRealizacion,
      selectedMode,
      infoGeneral,
      preRows,
      revRows,
      certRows,
      persRows,
      implRows,
      postRows
    };
  }

  function generatePdf() {
    if (btnPdf?.disabled) {
      alert("Complete el formulario para habilitar la descarga del PDF.");
      return;
    }

    if (!window.jspdf?.jsPDF) {
      alert("No se cargó jsPDF. Verifique que los script CDN estén en el HTML.");
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

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });

    const pageW = doc.internal.pageSize.getWidth();
    const margin = 28;

    doc.setFillColor(11, 95, 255);
    doc.rect(margin, 22, pageW - margin * 2, 34, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(255);
    doc.text("CHECKLIST GESTIÓN DE CAMBIOS TCS", pageW / 2, 44, { align: "center" });
    doc.setFontSize(10);
    doc.setTextColor(0);

    const labelX = margin;
    const valueX = 210;
    const y1 = 78;
    const y2 = 94;
    const y3 = 110;

    doc.setFont("helvetica", "normal");
    doc.text("Nombres completos:", labelX, y1);
    doc.text("N° de Control de Cambios:", labelX, y2);
    doc.text("Fecha de realización:", labelX, y3);

    doc.setFont("helvetica", "bold");
    doc.text(`${data.nombre}`, valueX, y1);
    doc.text(`${data.numeroCambio}`, valueX, y2);
    doc.text(`${data.fechaRealizacion}`, valueX, y3);

    let cursorY = 126;

    function drawSection(title, rows) {
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
          fillColor: [231, 241, 255],
          textColor: 0,
          fontStyle: "bold"
        },
        columnStyles: {
          0: { cellWidth: pageW - margin * 2 - 90, overflow: "linebreak" },
          1: { cellWidth: 90, halign: "center", valign: "middle" }
        }
      });

      cursorY = doc.lastAutoTable.finalY + 10;
    }

    drawSection("1. INFORMACIÓN GENERAL", data.infoGeneral);

    if (data.selectedMode === "CERTIFICACION") {
      if (hasRows(data.preRows)) drawSection("2. PRE REQUISITOS GESTIÓN DE CAMBIOS TCS", data.preRows);
      if (hasRows(data.revRows)) drawSection("3. REVISIÓN GENERAL DEL CAMBIO", data.revRows);
      if (hasRows(data.certRows)) drawSection("4. CERTIFICACIÓN DE TAREAS", data.certRows);
      if (hasRows(data.persRows)) drawSection("5. PERSONAL ASIGNADO", data.persRows);
    } else {
      if (hasRows(data.implRows)) drawSection("2. IMPLEMENTACIÓN", data.implRows);
      if (hasRows(data.postRows)) drawSection("3. POST IMPLEMENTACIÓN", data.postRows);
    }

    const fileName = `Checklist_${data.numeroCambio}_${data.nombre.replace(/\s+/g, "_")}.pdf`;
    doc.save(fileName);
  }

  btnPdf?.addEventListener("click", generatePdf);

  refreshSubmitState();
});