(function () {
  "use strict";

  // ===== State =====
  const state = {
    hrFile: null, // { name, data: ArrayBuffer }
    impresos: [], // [{ name, data: ArrayBuffer }]
    ceFile: null, // { name, data: ArrayBuffer, isJson: bool }
    lastResult: null, // { rows: ZebraRow[], xlsxBlob: Blob }
  };

  // ===== Utils =====
  function normalizeStr(s) {
    if (!s) return "";
    return String(s)
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  }

  function stripStr(val) {
    if (val === null || val === undefined) return "";
    return String(val).trim();
  }

  function normalizeColNames(rows) {
    if (!rows.length) return rows;
    const keys = Object.keys(rows[0]);
    const mapping = {};
    for (const k of keys) {
      mapping[k] = normalizeStr(k);
    }
    return rows.map((row) => {
      const newRow = {};
      for (const k of keys) {
        newRow[mapping[k]] = row[k];
      }
      return newRow;
    });
  }

  function extraerNumeroFactura(remito) {
    const r = stripStr(remito);
    let parte;
    if (r.includes("-")) {
      parte = r.split("-").slice(1).join("-");
    } else {
      parte = r;
    }
    const numero = parte.replace(/^0+/, "");
    return numero || "0";
  }

  // ===== XLSX Parsing (client-side via SheetJS) =====
  function parseXlsx(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    return wb;
  }

  function sheetToJson(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  }

  function loadHR(arrayBuffer) {
    const wb = parseXlsx(arrayBuffer);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
    return normalizeColNames(raw);
  }

  function loadImpresos(arrayBuffers) {
    const allRows = [];
    for (const ab of arrayBuffers) {
      try {
        const wb = parseXlsx(ab);
        // Try 'Hoja1' first, then first sheet
        const sheetName = wb.SheetNames.includes("Hoja1") ? "Hoja1" : wb.SheetNames[0];
        const raw = sheetToJson(wb, sheetName);
        const normalized = normalizeColNames(raw);

        // Detectar columnas una sola vez por archivo (los encabezados no cambian por fila).
        // Detección por prioridad: coincidencia exacta primero, evitando falsos positivos
        // como "TranspUnit" o "Unit of Weight" que contienen "unit" pero NO son el HU.
        const sampleKeys = normalized.length ? Object.keys(normalized[0]) : [];

        const remitoKey =
          sampleKeys.find((k) => k === "remito") ||
          sampleKeys.find((k) => k === "referencia") ||
          sampleKeys.find(
            (k) =>
              k.includes("remito") &&
              !k.includes("date") &&
              !k.includes("fecha") &&
              !k.includes("pdf")
          ) ||
          sampleKeys.find((k) => k.includes("referencia"));

        const huKey =
          sampleKeys.find((k) => k === "handling unit") ||
          sampleKeys.find(
            (k) => k.includes("handling") && k.includes("unit") && !k.includes("uom")
          ) ||
          sampleKeys.find((k) => k.includes("handling")) ||
          sampleKeys.find((k) => k === "hu" || k === "unidad");

        for (const row of normalized) {
          if (remitoKey && huKey) {
            const remito = stripStr(row[remitoKey]);
            const hu = stripStr(row[huKey]);
            if (remito && hu) {
              allRows.push({ remito, "handling unit": hu });
            }
          }
        }
      } catch (e) {
        console.error("Error parsing impreso:", e);
      }
    }
    return allRows;
  }

  function loadCe(data, isJson, filename) {
    if (isJson || (filename && filename.endsWith(".json"))) {
      try {
        const text = new TextDecoder().decode(data);
        const raw = JSON.parse(text);
        if (Array.isArray(raw)) {
          return raw
            .map((r) => ({
              localidad: normalizeStr(stripStr(r.localidad || r.Localidad || "")),
              ruta: stripStr(r.ruta || r.Ruta || ""),
            }))
            .filter((r) => r.localidad);
        }
      } catch (e) {
        return [];
      }
      return [];
    }
    // Excel
    const wb = parseXlsx(data);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
    const normalized = normalizeColNames(raw);

    // El Ce real trae la ciudad limpia dentro de "CONCA_CE_LOC" (formato
    // "CLIENTE - CIUDAD"); la columna "LOCALIDAD" es una clave interna sucia
    // con prefijos de cliente (p. ej. "CASTRO RAFAELA"), que casi nunca matchea
    // con la "Ciudad Destinatario" del HR. Si existe CONCA_CE_LOC, derivamos la
    // ciudad de ahí; si no, caemos al formato simple documentado (Localidad/Ruta).
    const hasConca = normalized.length && "conca_ce_loc" in normalized[0];

    const cityFromConca = (val) => {
      // Algunos separadores usan espacio no separable (\xa0) en vez de espacio.
      const s = stripStr(val).replace(/ /g, " ");
      // La ciudad es el texto tras el último "-" (las ciudades no llevan "-";
      // los nombres de cliente sí pueden tener espacios/puntos).
      const idx = s.lastIndexOf("-");
      return idx >= 0 ? s.slice(idx + 1) : s;
    };

    return normalized
      .map((row) => ({
        localidad: normalizeStr(
          hasConca
            ? cityFromConca(row["conca_ce_loc"])
            : stripStr(row["localidad"] || row["ciudad"] || "")
        ),
        ruta: stripStr(row["ruta"] || ""),
      }))
      .filter((r) => r.localidad);
  }

  // ===== Core Processing =====
  function processData(hrRows, impresos, ce) {
    const alerts = [];

    try {
      const sampleHR = hrRows[0] || {};
      const hrKeys = Object.keys(sampleHR);

      const findCol = (patterns) => {
        for (const p of patterns) {
          const found = hrKeys.find((k) => k.includes(p));
          if (found) return found;
        }
        return undefined;
      };

      const colFactura = findCol([
        "factura",
        "n factura",
        "nfactura",
        "no factura",
        "nro factura",
        "num factura",
        "numero factura",
      ]);
      const colDestinatario = findCol(["destinatario"]);
      const colCiudadDest = findCol([
        "ciudad destinatario",
        "ciudad dest",
        "ciudad_dest",
        "ciudaddestinatario",
      ]);

      if (!colFactura) {
        return {
          success: false,
          alerts,
          error: `No se encontró columna "Nº Factura" en el HR. Columnas detectadas: ${hrKeys.join(", ")}`,
        };
      }
      if (!colDestinatario) {
        return {
          success: false,
          alerts,
          error: `No se encontró columna "Destinatario" en el HR. Columnas detectadas: ${hrKeys.join(", ")}`,
        };
      }
      if (!colCiudadDest) {
        return {
          success: false,
          alerts,
          error: `No se encontró columna "Ciudad Destinatario" en el HR. Columnas detectadas: ${hrKeys.join(", ")}`,
        };
      }

      // Derive factura from impresos remito
      const impresosWithFac = impresos.map((row) => ({
        ...row,
        factura_derivada: extraerNumeroFactura(row.remito),
      }));

      // Build sets for cross-check
      const facturasHR = new Set(
        hrRows
          .filter((r) => stripStr(r[colDestinatario]) !== "Total Gral.")
          .map((r) => stripStr(r[colFactura]))
          .filter(Boolean)
      );
      const facturasImpresos = new Set(
        impresosWithFac.map((r) => r.factura_derivada).filter(Boolean)
      );

      // Alerts: in HR but not in impresos
      for (const fac of [...facturasHR].sort()) {
        if (!facturasImpresos.has(fac)) {
          alerts.push({
            type: "warning",
            msg: `Factura "${fac}" presente en HR pero sin Remito correspondiente en Impresos`,
          });
        }
      }
      // Alerts: in impresos but not in HR
      for (const fac of [...facturasImpresos].sort()) {
        if (!facturasHR.has(fac)) {
          const remitos = [
            ...new Set(
              impresosWithFac.filter((r) => r.factura_derivada === fac).map((r) => r.remito)
            ),
          ];
          for (const rem of remitos) {
            alerts.push({
              type: "warning",
              msg: `Factura "${fac}" de Impresos (Remito: ${rem}) no encontrada en HR`,
            });
          }
        }
      }

      // Build HR lookup
      const hrByFactura = new Map();
      for (const row of hrRows) {
        const fac = stripStr(row[colFactura]);
        if (!fac || stripStr(row[colDestinatario]) === "Total Gral.") continue;
        if (!hrByFactura.has(fac)) {
          hrByFactura.set(fac, {
            destinatario: stripStr(row[colDestinatario]),
            ciudad: stripStr(row[colCiudadDest]),
          });
        }
      }

      // Build Ce lookup
      const ceByLocalidad = new Map();
      for (const row of ce) {
        ceByLocalidad.set(row.localidad, row.ruta);
      }

      // Merge + dedup by HU
      const zebraRows = [];
      const seenHU = new Set();

      for (const imp of impresosWithFac) {
        const hrEntry = hrByFactura.get(imp.factura_derivada);
        if (!hrEntry) continue;

        const hu = stripStr(imp["handling unit"]);
        if (seenHU.has(hu)) continue;
        seenHU.add(hu);

        const ciudadNorm = normalizeStr(hrEntry.ciudad);
        const ruta = ceByLocalidad.get(ciudadNorm) || "";

        zebraRows.push({
          Referencia: imp.remito,
          "Handling Unit": hu,
          Destino: hrEntry.destinatario,
          Ciudad: hrEntry.ciudad,
          Ruta: ruta,
        });
      }

      return { success: true, rows: zebraRows, alerts };
    } catch (e) {
      return { success: false, alerts, error: `Error en procesamiento: ${e.message || e}` };
    }
  }

  // ===== XLSX Export =====
  function generateZebraXLSX(rows) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, {
      header: ["Referencia", "Handling Unit", "Destino", "Ciudad", "Ruta"],
    });
    // Style header row (column widths)
    ws["!cols"] = [
      { wch: 18 }, // Referencia
      { wch: 18 }, // Handling Unit
      { wch: 30 }, // Destino
      { wch: 20 }, // Ciudad
      { wch: 10 }, // Ruta
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Etiquetas");
    const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    return new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 500);
  }

  function getTimestamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  }

  // ===== UI Helpers =====
  function showToast(msg, type = "info", duration = 3500) {
    const icons = {
      success:
        '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg>',
      error:
        '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
      info: '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>',
    };
    const el = document.createElement("div");
    el.className = `toast ${type}`;
    el.innerHTML = `${icons[type] || ""}<span>${msg}</span>`;
    const container = document.getElementById("toast-container");
    container.appendChild(el);
    setTimeout(() => {
      el.style.opacity = "0";
      el.style.transform = "translateX(20px)";
      el.style.transition = "opacity 0.3s, transform 0.3s";
      setTimeout(() => el.remove(), 300);
    }, duration);
  }

  function updateStatusCard(id, text, cls) {
    const el = document.getElementById(id);
    el.textContent = text;
    el.className = "status-value " + (cls || "");
  }

  function renderFileTags(containerId, files, removeCallback) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";
    files.forEach((f, i) => {
      const tag = document.createElement("div");
      tag.className = "file-tag";
      tag.innerHTML = `
          <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
          <span title="${f.name}">${f.name}</span>
          <button class="file-tag-remove" aria-label="Eliminar ${f.name}">✕</button>
        `;
      tag.querySelector(".file-tag-remove").addEventListener("click", () => removeCallback(i));
      container.appendChild(tag);
    });
  }

  function updateProcessBtn() {
    const btn = document.getElementById("btn-process");
    const ready = state.hrFile && state.impresos.length > 0;
    btn.disabled = !ready;
  }

  function showAlerts(alerts) {
    const container = document.getElementById("alerts-container");
    container.innerHTML = "";
    if (!alerts.length) {
      container.classList.add("hidden");
      return;
    }
    container.classList.remove("hidden");
    alerts.forEach((a) => {
      const div = document.createElement("div");
      const isWarn = a.type === "warning";
      div.className = `alert alert-${isWarn ? "warning" : "error"}`;
      div.innerHTML = `
          <div class="alert-icon">
            ${
              isWarn
                ? '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'
                : '<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>'
            }
          </div>
          <span>${a.msg || a}</span>
        `;
      container.appendChild(div);
    });
  }

  function showResult(rows, blob) {
    state.lastResult = { rows, blob };
    document.getElementById("result-count").textContent = rows.length;
    const hasNoRuta = rows.filter((r) => !r["Ruta"]).length;
    document.getElementById("result-desc").textContent =
      `XLSX descargado · ${rows.length} Handling Units` +
      (hasNoRuta > 0 ? ` · ⚠ ${hasNoRuta} sin Ruta asignada (Ce sin match)` : "");

    // Populate preview table
    const tbody = document.getElementById("preview-tbody");
    tbody.innerHTML = "";
    const preview = rows.slice(0, 50);
    for (const row of preview) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
          <td>${row["Referencia"]}</td>
          <td>${row["Handling Unit"]}</td>
          <td>${row["Destino"]}</td>
          <td>${row["Ciudad"]}</td>
          <td>${row["Ruta"] || '<span style="color:var(--color-text-faint)">—</span>'}</td>
        `;
      tbody.appendChild(tr);
    }

    document.getElementById("result-section").classList.remove("hidden");
  }

  // ===== Event Handlers =====

  // Theme toggle
  const themeBtn = document.getElementById("btn-theme");
  let currentTheme = "dark";
  document.documentElement.setAttribute("data-theme", currentTheme);

  themeBtn.addEventListener("click", () => {
    currentTheme = currentTheme === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", currentTheme);
    themeBtn.innerHTML =
      currentTheme === "dark"
        ? '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>'
        : '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>';
  });

  // Setup file drop zone
  function setupDropZone(dzId, inputId, multiple, onFiles) {
    const dz = document.getElementById(dzId);
    const input = document.getElementById(inputId);

    ["dragover", "dragenter"].forEach((evt) => {
      dz.addEventListener(evt, (e) => {
        e.preventDefault();
        dz.classList.add("drag-over");
      });
    });
    ["dragleave", "drop"].forEach((evt) => {
      dz.addEventListener(evt, () => dz.classList.remove("drag-over"));
    });
    dz.addEventListener("drop", (e) => {
      e.preventDefault();
      const files = [...(e.dataTransfer.files || [])];
      if (files.length) onFiles(files);
    });
    input.addEventListener("change", () => {
      const files = [...(input.files || [])];
      if (files.length) onFiles(files);
      input.value = ""; // reset so same file can be re-selected
    });
    dz.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        input.click();
      }
    });
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // HR
  setupDropZone("dz-hr", "file-hr", false, async (files) => {
    const f = files[0];
    if (!f.name.match(/\.(xlsx|xls)$/i)) {
      showToast("Solo se aceptan archivos XLSX o XLS para el HR", "error");
      return;
    }
    const data = await readFileAsArrayBuffer(f);
    state.hrFile = { name: f.name, data };
    renderFileTags("tags-hr", [f], () => {
      state.hrFile = null;
      renderFileTags("tags-hr", [], () => {});
      updateStatusCard("status-hr", "Sin archivo", "required-missing");
      updateProcessBtn();
    });
    updateStatusCard("status-hr", f.name, "has-files");
    updateProcessBtn();
    showToast(`HR cargado: ${f.name}`, "success");
  });

  // Impresos (multiple)
  setupDropZone("dz-impresos", "file-impresos", true, async (files) => {
    const xlsxFiles = files.filter((f) => f.name.match(/\.(xlsx|xls)$/i));
    if (!xlsxFiles.length) {
      showToast("No se encontraron archivos XLSX válidos", "error");
      return;
    }
    for (const f of xlsxFiles) {
      const data = await readFileAsArrayBuffer(f);
      // Avoid duplicates
      if (!state.impresos.find((x) => x.name === f.name)) {
        state.impresos.push({ name: f.name, data });
      }
    }
    const fileObjs = state.impresos.map((x) => ({ name: x.name }));
    renderFileTags("tags-impresos", fileObjs, (i) => {
      state.impresos.splice(i, 1);
      const fo = state.impresos.map((x) => ({ name: x.name }));
      renderFileTags("tags-impresos", fo, (j) => {
        state.impresos.splice(j, 1);
        const fo2 = state.impresos.map((x) => ({ name: x.name }));
        renderFileTags("tags-impresos", fo2, () => {});
        updateStatusCard(
          "status-impresos",
          state.impresos.length ? `${state.impresos.length} archivos` : "Sin archivos",
          state.impresos.length ? "has-files" : "required-missing"
        );
        updateProcessBtn();
      });
      updateStatusCard(
        "status-impresos",
        state.impresos.length ? `${state.impresos.length} archivos` : "Sin archivos",
        state.impresos.length ? "has-files" : "required-missing"
      );
      updateProcessBtn();
    });
    updateStatusCard(
      "status-impresos",
      `${state.impresos.length} archivo${state.impresos.length > 1 ? "s" : ""}`,
      "has-files"
    );
    updateProcessBtn();
    showToast(
      `${xlsxFiles.length} Impreso${xlsxFiles.length > 1 ? "s" : ""} cargado${xlsxFiles.length > 1 ? "s" : ""}`,
      "success"
    );
  });

  // Ce
  setupDropZone("dz-ce", "file-ce", false, async (files) => {
    const f = files[0];
    if (!f.name.match(/\.(xlsx|xls|json)$/i)) {
      showToast("Ce solo acepta XLSX, XLS o JSON", "error");
      return;
    }
    const isJson = f.name.endsWith(".json");
    const data = await readFileAsArrayBuffer(f);
    state.ceFile = { name: f.name, data, isJson };
    renderFileTags("tags-ce", [f], () => {
      state.ceFile = null;
      renderFileTags("tags-ce", [], () => {});
      updateStatusCard("status-ce", "Opcional", "");
      document.getElementById("status-ce").style.color = "var(--color-warning)";
    });
    updateStatusCard("status-ce", f.name, "has-files");
    document.getElementById("status-ce").style.color = "";
    showToast(`Ce cargado: ${f.name}`, "success");
  });

  // Process
  document.getElementById("btn-process").addEventListener("click", async () => {
    if (!state.hrFile || !state.impresos.length) return;

    // Show processing
    const indicator = document.getElementById("processing-indicator");
    const processMsg = document.getElementById("processing-msg");
    indicator.classList.remove("hidden");
    document.getElementById("result-section").classList.add("hidden");
    document.getElementById("alerts-container").classList.add("hidden");
    document.getElementById("btn-process").disabled = true;

    // Give the UI a chance to render before heavy work
    await new Promise((r) => setTimeout(r, 50));

    try {
      processMsg.textContent = "Leyendo HR...";
      const hrRows = loadHR(state.hrFile.data);

      processMsg.textContent = "Leyendo Impresos...";
      const impresos = loadImpresos(state.impresos.map((x) => x.data));

      processMsg.textContent = "Leyendo Ce...";
      let ce = [];
      if (state.ceFile) {
        ce = loadCe(state.ceFile.data, state.ceFile.isJson, state.ceFile.name);
      }

      processMsg.textContent = "Cruzando datos...";
      await new Promise((r) => setTimeout(r, 30));

      const result = processData(hrRows, impresos, ce);

      indicator.classList.add("hidden");
      document.getElementById("btn-process").disabled = false;
      updateProcessBtn();

      showAlerts(result.alerts || []);

      if (!result.success) {
        const errAlert = [{ type: "error", msg: result.error }];
        showAlerts([...errAlert, ...(result.alerts || [])]);
        showToast("Error en el procesamiento", "error", 5000);
        return;
      }

      if (!result.rows.length) {
        showAlerts([
          {
            type: "error",
            msg: "No se encontraron filas con match entre HR e Impresos. Verificá que las facturas coincidan.",
          },
          ...(result.alerts || []),
        ]);
        showToast("Sin resultados — verificá los archivos", "error", 5000);
        return;
      }

      const blob = generateZebraXLSX(result.rows);
      const filename = `CEVA_Zebra_Etiquetas_${getTimestamp()}.xlsx`;
      downloadBlob(blob, filename);
      showResult(result.rows, blob);

      const warnCount = (result.alerts || []).filter((a) => a.type === "warning").length;
      const msg = `✓ ${result.rows.length} etiquetas generadas${warnCount ? ` · ${warnCount} alertas` : ""}`;
      showToast(msg, "success", 4000);
    } catch (e) {
      indicator.classList.add("hidden");
      document.getElementById("btn-process").disabled = false;
      updateProcessBtn();
      showAlerts([{ type: "error", msg: `Error inesperado: ${e.message}` }]);
      showToast("Error inesperado — verificá la consola", "error", 5000);
      console.error(e);
    }
  });

  // Download again
  document.getElementById("btn-download-again").addEventListener("click", () => {
    if (!state.lastResult) return;
    const filename = `CEVA_Zebra_Etiquetas_${getTimestamp()}.xlsx`;
    downloadBlob(state.lastResult.blob, filename);
    showToast("Descargando...", "info");
  });

  // Clear
  document.getElementById("btn-clear").addEventListener("click", () => {
    state.hrFile = null;
    state.impresos = [];
    state.ceFile = null;
    state.lastResult = null;

    renderFileTags("tags-hr", [], () => {});
    renderFileTags("tags-impresos", [], () => {});
    renderFileTags("tags-ce", [], () => {});

    updateStatusCard("status-hr", "Sin archivo", "required-missing");
    updateStatusCard("status-impresos", "Sin archivos", "required-missing");
    updateStatusCard("status-ce", "Opcional", "");
    document.getElementById("status-ce").style.color = "var(--color-warning)";

    document.getElementById("alerts-container").classList.add("hidden");
    document.getElementById("result-section").classList.add("hidden");
    document.getElementById("processing-indicator").classList.add("hidden");
    document.getElementById("alerts-container").innerHTML = "";

    updateProcessBtn();
    showToast("Todo limpiado", "info");
  });

  // Instructions toggle
  document.getElementById("btn-instructions-toggle").addEventListener("click", () => {
    const panel = document.getElementById("instructions-panel");
    panel.classList.toggle("hidden");
  });

  // Preview table toggle
  document.getElementById("toggle-preview").addEventListener("click", () => {
    const btn = document.getElementById("toggle-preview");
    const content = document.getElementById("preview-content");
    const expanded = btn.getAttribute("aria-expanded") === "true";
    btn.setAttribute("aria-expanded", String(!expanded));
    btn.querySelector("span").textContent = expanded ? "Mostrar" : "Ocultar";
    content.style.maxHeight = expanded ? "0" : "400px";
    if (!expanded) content.style.maxHeight = "";
  });

  // ===== PWA Install Prompt =====
  let deferredPrompt = null;
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredPrompt = e;
    document.getElementById("install-bar").classList.add("visible");
  });

  document.getElementById("btn-install").addEventListener("click", async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    deferredPrompt = null;
    document.getElementById("install-bar").classList.remove("visible");
    if (outcome === "accepted") showToast("App instalada correctamente", "success");
  });

  document.getElementById("btn-dismiss-install").addEventListener("click", () => {
    document.getElementById("install-bar").classList.remove("visible");
  });

  // ===== Service Worker Registration =====
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker
        .register("./sw.js", { scope: "./" })
        .then((reg) => {
          console.log("[PWA] Service Worker registrado:", reg.scope);
        })
        .catch((err) => {
          console.warn("[PWA] Service Worker falló:", err);
        });
    });
  }

  // ===== Init =====
  updateProcessBtn();
})();
