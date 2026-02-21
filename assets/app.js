/* global XLSX */

function formatMoney(amount) {
  const n = Number.isFinite(amount) ? amount : 0;
  return n.toLocaleString("es-ES", { style: "currency", currency: "USD" });
}

function parseAmount(raw) {
  if (raw == null) return 0;
  const n = Number(String(raw).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function todayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function fileStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}${mm}${dd}`;
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function getEl(selector) {
  const el = document.querySelector(selector);
  if (!el) throw new Error(`No se encontró el elemento: ${selector}`);
  return el;
}

function setNotice(message, { tone = "info" } = {}) {
  const el = getEl('[data-role="notice"]');
  el.textContent = message;
  el.classList.toggle("notice--error", tone === "error");
}

function clearNotice() {
  setNotice("");
}

function rowTemplate(data) {
  const {
    id,
    fecha = "",
    categoria = "",
    descripcion = "",
    metodo = "Efectivo",
    monto = "",
  } = data;

  const tr = document.createElement("tr");
  tr.dataset.rowId = id;

  tr.innerHTML = `
    <td>
      <input class="table__row-input" type="date" value="${escapeAttr(fecha)}" data-field="fecha" aria-label="Fecha" />
    </td>
    <td>
      <input class="table__row-input" type="text" value="${escapeAttr(
        categoria
      )}" data-field="categoria" aria-label="Categoría" placeholder="Categoría" />
    </td>
    <td>
      <input class="table__row-input" type="text" value="${escapeAttr(
        descripcion
      )}" data-field="descripcion" aria-label="Descripción" placeholder="Descripción" />
    </td>
    <td>
      <select class="table__row-input" data-field="metodo" aria-label="Método">
        ${["Efectivo", "Débito", "Crédito", "Transferencia", "Otro"]
          .map(
            (opt) =>
              `<option value="${escapeAttr(opt)}"${
                opt === metodo ? " selected" : ""
              }>${escapeHtml(opt)}</option>`
          )
          .join("")}
      </select>
    </td>
    <td class="table__num">
      <input class="table__row-input" type="number" step="0.01" min="0" inputmode="decimal" value="${escapeAttr(
        monto
      )}" data-field="monto" aria-label="Monto" placeholder="0.00" />
    </td>
    <td class="table__actions">
      <button type="button" class="icon-button icon-button--danger" data-action="delete-row" aria-label="Eliminar fila">
        Eliminar
      </button>
    </td>
  `.trim();

  return tr;
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(str) {
  return escapeHtml(str).replaceAll("\n", " ");
}

function uid() {
  if ("randomUUID" in crypto) return crypto.randomUUID();
  return `row_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

function readRowsFromDOM() {
  const tbody = getEl('[data-role="rows"]');
  const trs = Array.from(tbody.querySelectorAll("tr"));
  return trs
    .map((tr) => {
      const get = (field) =>
        tr.querySelector(`[data-field="${field}"]`)?.value ?? "";
      return {
        id: tr.dataset.rowId || uid(),
        fecha: normalizeText(get("fecha")),
        categoria: normalizeText(get("categoria")),
        descripcion: normalizeText(get("descripcion")),
        metodo: normalizeText(get("metodo")),
        monto: normalizeText(get("monto")),
      };
    })
    .filter((r) => Object.values(r).some((v) => String(v).trim() !== ""));
}

function computeTotal(rows) {
  return rows.reduce((sum, r) => sum + parseAmount(r.monto), 0);
}

function updateTotalUI() {
  const rows = readRowsFromDOM();
  const total = computeTotal(rows);
  getEl('[data-role="total"]').textContent = formatMoney(total);
}

function addRow(data) {
  const tbody = getEl('[data-role="rows"]');
  const row = rowTemplate({ id: uid(), ...data });
  tbody.prepend(row);
  updateTotalUI();
}

function resetAll() {
  getEl('[data-role="rows"]').innerHTML = "";
  updateTotalUI();
  clearNotice();
  const form = getEl('[data-form="expense"]');
  form.reset();
  getEl("#fecha").value = todayISO();
  getEl("#categoria").focus();
}

function validateBeforeExport(rows) {
  if (!rows.length) {
    return "Agrega al menos un gasto antes de descargar el Excel.";
  }

  const hasMissing = rows.some((r) => !r.fecha || !r.categoria || !r.monto);
  if (hasMissing) {
    return "Completa Fecha, Categoría y Monto en todas las filas antes de exportar.";
  }

  const hasInvalidAmount = rows.some((r) => parseAmount(r.monto) <= 0);
  if (hasInvalidAmount) {
    return "El Monto debe ser mayor a 0 en todas las filas.";
  }

  return null;
}

function buildGastosSheet(rows) {
  const header = ["Fecha", "Categoría", "Descripción", "Método", "Monto"];
  const data = rows
    .slice()
    .sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)))
    .map((r) => [
      r.fecha,
      r.categoria,
      r.descripcion,
      r.metodo,
      Number(parseAmount(r.monto).toFixed(2)),
    ]);
  return [header, ...data];
}

function groupSum(rows, keyFn) {
  const map = new Map();
  for (const r of rows) {
    const key = keyFn(r);
    const prev = map.get(key) ?? 0;
    map.set(key, prev + parseAmount(r.monto));
  }
  return map;
}

function buildResumenSheet(rows) {
  const total = computeTotal(rows);
  const porCategoria = groupSum(rows, (r) => r.categoria || "Sin categoría");
  const porDia = groupSum(rows, (r) => r.fecha || "Sin fecha");

  const section = (title, entries) => {
    const out = [[title, "Total"]];
    for (const [k, v] of entries) out.push([k, Number(v.toFixed(2))]);
    out.push(["", ""]);
    return out;
  };

  const catEntries = Array.from(porCategoria.entries()).sort((a, b) =>
    String(a[0]).localeCompare(String(b[0]), "es")
  );
  const dayEntries = Array.from(porDia.entries()).sort((a, b) =>
    String(a[0]).localeCompare(String(b[0]))
  );

  return [
    ["Resumen", ""],
    ["Total general", Number(total.toFixed(2))],
    ["", ""],
    ...section("Totales por categoría", catEntries),
    ...section("Totales por día", dayEntries),
  ];
}

function exportXlsx() {
  if (typeof XLSX === "undefined") {
    setNotice(
      "No se pudo cargar la librería de Excel. Revisa tu conexión o vuelve a intentar.",
      { tone: "error" }
    );
    return;
  }

  const rows = readRowsFromDOM();
  const error = validateBeforeExport(rows);
  if (error) {
    setNotice(error, { tone: "error" });
    return;
  }

  const wb = XLSX.utils.book_new();

  const gastosAoA = buildGastosSheet(rows);
  const wsGastos = XLSX.utils.aoa_to_sheet(gastosAoA);
  wsGastos["!cols"] = [
    { wch: 12 },
    { wch: 18 },
    { wch: 28 },
    { wch: 16 },
    { wch: 12 },
  ];
  XLSX.utils.book_append_sheet(wb, wsGastos, "Gastos");

  const resumenAoA = buildResumenSheet(rows);
  const wsResumen = XLSX.utils.aoa_to_sheet(resumenAoA);
  wsResumen["!cols"] = [{ wch: 28 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");

  const name = `gastos-diarios_${fileStamp()}.xlsx`;
  clearNotice();
  XLSX.writeFile(wb, name, { compression: true });
  setNotice(`Descarga lista: ${name}`);
}

function init() {
  const form = getEl('[data-form="expense"]');
  const fecha = getEl("#fecha");
  fecha.value = todayISO();

  form.addEventListener("submit", (ev) => {
    ev.preventDefault();
    clearNotice();

    const fd = new FormData(form);
    const data = {
      fecha: normalizeText(fd.get("fecha")),
      categoria: normalizeText(fd.get("categoria")),
      descripcion: normalizeText(fd.get("descripcion")),
      metodo: normalizeText(fd.get("metodo")) || "Efectivo",
      monto: normalizeText(fd.get("monto")),
    };

    if (!data.fecha || !data.categoria || !data.monto) {
      setNotice("Completa Fecha, Categoría y Monto para agregar el gasto.", {
        tone: "error",
      });
      return;
    }

    if (parseAmount(data.monto) <= 0) {
      setNotice("El Monto debe ser mayor a 0.", { tone: "error" });
      return;
    }

    addRow(data);
    form.reset();
    fecha.value = todayISO();
    getEl("#categoria").focus();
    setNotice("Gasto agregado.");
  });

  document.addEventListener("click", (ev) => {
    const target = ev.target instanceof Element ? ev.target : null;
    if (!target) return;

    const actionEl = target.closest("[data-action]");
    if (!actionEl) return;

    const action = actionEl.getAttribute("data-action");
    if (!action) return;

    if (action === "delete-row") {
      const tr = actionEl.closest("tr");
      if (tr) tr.remove();
      updateTotalUI();
      setNotice("Fila eliminada.");
      return;
    }

    if (action === "export-xlsx") {
      exportXlsx();
      return;
    }

    if (action === "reset-all") {
      resetAll();
      setNotice("Se reinició la tabla.");
      return;
    }

    if (action === "add-empty-row") {
      addRow({ fecha: todayISO() });
      setNotice("Fila agregada. Completa los datos y exporta cuando quieras.");
    }
  });

  document.addEventListener("input", (ev) => {
    const target = ev.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement))
      return;
    if (!target.closest("tbody")) return;
    updateTotalUI();
  });

  updateTotalUI();
  clearNotice();
}

document.addEventListener("DOMContentLoaded", init);

