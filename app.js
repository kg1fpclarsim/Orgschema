/* global Papa, d3, d3OrgChart, OrgChart, htmlToImage, jspdf */
(() => {
  const d3Api = typeof d3 !== "undefined" ? d3 : null;
  const PapaApi = typeof Papa !== "undefined" ? Papa : null;
  const htmlToImageApi = typeof htmlToImage !== "undefined" ? htmlToImage : null;
  const jsPdfApi = typeof jspdf !== "undefined" ? jspdf : null;

  const OrgChartCtor =
    (typeof d3OrgChart !== "undefined" && d3OrgChart && d3OrgChart.OrgChart) ||
    (d3Api && d3Api.OrgChart) ||
    (typeof OrgChart !== "undefined" ? OrgChart : null);

  const els = {
    fileInput: document.getElementById("fileInput"),
    searchInput: document.getElementById("searchInput"),
    filterBolag: document.getElementById("filterBolag"),
    filterPlats: document.getElementById("filterPlats"),
    filterTrafik: document.getElementById("filterTrafik"),
    colorBy: document.getElementById("colorBy"),
    palette: document.getElementById("palette"),
    expandAllBtn: document.getElementById("expandAllBtn"),
    collapseAllBtn: document.getElementById("collapseAllBtn"),
    fitBtn: document.getElementById("fitBtn"),
    pngBtn: document.getElementById("pngBtn"),
    pdfBtn: document.getElementById("pdfBtn"),
    hint: document.getElementById("hint"),
    chart: document.getElementById("chart"),
    overlay: document.getElementById("detailsOverlay"),
    detailsBody: document.getElementById("detailsBody"),
    detailsIdPill: document.getElementById("detailsIdPill"),
    closeDetails: document.getElementById("closeDetails"),
    debugPanel: document.getElementById("debugPanel"),
    debugOutput: document.getElementById("debugOutput"),
  };

  const FALLBACK_COLORS = ["#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948"];

  const COLOR_PRESETS = {
    tableau: d3Api?.schemeTableau10 || FALLBACK_COLORS,
    set3: d3Api?.schemeSet3 || FALLBACK_COLORS,
    pastel1: d3Api?.schemePastel1 || FALLBACK_COLORS,
  };

  const state = {
    data: [],
    filtered: [],
    finalData: [],
    selected: null,
    chart: null,
    q: "",
    bolag: "all",
    plats: "all",
    trafik: "all",
    colorBy: "branch",
    palette: "tableau",
    showAllSam: false,
    showFullDesc: false,
    debug: null,
    connections: [],
  };

  function norm(v) {
    return (v ?? "").toString().trim();
  }

  function sanitizeIdentifier(value) {
    return norm(value)
      .normalize("NFC")
      .replace(/[\u200B-\u200D\uFEFF]/g, "");
  }

  function escapeHtml(s) {
    return (s ?? "")
      .toString()
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function parseMaybeList(value) {
    const v = norm(value);
    if (!v) return [];
    try {
      const parsed = JSON.parse(v);
      if (Array.isArray(parsed)) return parsed.map((x) => norm(x)).filter(Boolean);
    } catch {}
    if (v.includes(";")) return v.split(";").map((x) => x.trim()).filter(Boolean);
    if (v.includes(",")) return v.split(",").map((x) => x.trim()).filter(Boolean);
    return [v];
  }

  function normalizeHeaderKey(key) {
    return norm(key).replace(/^\uFEFF/, "").toLowerCase();
  }

  function getField(row, ...candidates) {
    if (!row || typeof row !== "object") return "";

    const entries = Object.entries(row);
    for (const candidate of candidates) {
      const direct = row[candidate];
      if (direct !== undefined && direct !== null && norm(direct)) return direct;

      const wanted = normalizeHeaderKey(candidate);
      const hit = entries.find(([k]) => normalizeHeaderKey(k) === wanted);
      if (!hit) continue;
      const [, value] = hit;
      if (value !== undefined && value !== null && norm(value)) return value;
    }

    return "";
  }

  function callIfFn(obj, name, ...args) {
    const fn = obj && obj[name];
    if (typeof fn === "function") {
      fn.apply(obj, args);
      return true;
    }
    return false;
  }

  function ts() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(
      d.getMinutes()
    )}${pad(d.getSeconds())}`;
  }

  function setControlsEnabled(enabled) {
    [
      "searchInput",
      "filterBolag",
      "filterPlats",
      "filterTrafik",
      "colorBy",
      "palette",
      "expandAllBtn",
      "collapseAllBtn",
      "fitBtn",
      "pngBtn",
      "pdfBtn",
    ].forEach(
      (k) => (els[k].disabled = !enabled)
    );
    els.palette.disabled = !enabled || state.colorBy === "none";
  }

  function uniqSorted(arr) {
    return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "sv"));
  }

  function fillSelect(sel, allOpt, values) {
    sel.innerHTML = "";
    const o0 = document.createElement("option");
    o0.value = allOpt[0];
    o0.textContent = allOpt[1];
    sel.appendChild(o0);
    values.forEach((v) => {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      sel.appendChild(o);
    });
  }

  function rebuildOptions() {
    fillSelect(els.filterBolag, ["all", "Alla bolag"], uniqSorted(state.data.map((d) => d.bolag)));
    fillSelect(els.filterPlats, ["all", "Alla (Plats / Nivå)"], uniqSorted(state.data.map((d) => d.plats)));
    fillSelect(els.filterTrafik, ["all", "Alla (Trafikområde)"], uniqSorted(state.data.map((d) => d.trafik)));
  }

  function applyFilters() {
    const qq = state.q.toLowerCase();
    state.filtered = state.data.filter((d) => {
      if (state.bolag !== "all" && d.bolag !== state.bolag) return false;
      if (state.plats !== "all" && d.plats !== state.plats) return false;
      if (state.trafik !== "all" && d.trafik !== state.trafik) return false;
      if (!qq) return true;

      const hay = [
        d.id,
        d.parentId,
        d.name,
        d.title,
        d.bolag,
        d.plats,
        d.trafik,
        ...(d.ansvar || []),
        ...(d.sam || []),
        d.arbetsbeskrivning,
      ]
        .join(" ")
        .toLowerCase();
      return hay.includes(qq);
    });
  }

  function keepAncestors() {
    if (!state.data.length) return [];
    if (!state.filtered.length) return [];
    const byId = new Map(state.data.map((d) => [d.id, d]));
    const keep = new Set();

    state.filtered.forEach((d) => {
      let cur = d;
      while (cur) {
        if (keep.has(cur.id)) break;
        keep.add(cur.id);
        cur = cur.parentId ? byId.get(cur.parentId) || null : null;
      }
    });

    return state.data.filter((d) => keep.has(d.id));
  }

  function makeColorState() {
    const rows = state.finalData;
    const byId = new Map(rows.map((d) => [d.id, d]));
    const roots = rows.filter((d) => !d.parentId || !byId.has(d.parentId));
    const root = roots[0] || null;

    function depth1AncestorId(id) {
      if (!root) return id;
      let cur = byId.get(id);
      if (!cur) return id;
      while (cur.parentId && byId.has(cur.parentId)) {
        const p = byId.get(cur.parentId);
        if (!p) break;
        if (p.id === root.id) return cur.id;
        cur = p;
      }
      return cur.id;
    }

    const keyFn = (d) => {
      if (state.colorBy === "none") return "";
      if (state.colorBy === "bolag") return d.bolag || "(saknar bolag)";
      if (state.colorBy === "plats") return d.plats || "(saknar plats)";
      if (state.colorBy === "trafik") return d.trafik || "(saknar trafik)";
      const bId = depth1AncestorId(d.id);
      const b = byId.get(bId);
      return (b?.name || b?.title || bId || "gren").toString();
    };

    const colors = COLOR_PRESETS[state.palette] || COLOR_PRESETS.tableau;
    const uniqKeys = Array.from(new Set(rows.map(keyFn).filter(Boolean)));
    const map = new Map();
    uniqKeys.forEach((k, i) => map.set(k, colors[i % colors.length]));

    const colorFor = (d) => {
      if (state.colorBy === "none") return { stripe: "#CBD5E1", soft: "rgba(203,213,225,.20)", key: "" };
      const key = keyFn(d);
      const stripe = map.get(key) || "#94A3B8";
      const c = d3Api?.color ? d3Api.color(stripe) : null;
      const soft = c ? c.copy({ opacity: 0.12 }).formatRgb() : "rgba(148,163,184,.12)";
      return { stripe, soft, key };
    };

    return { colorFor };
  }

  function unwrapNodeData(nodeLike) {
    return nodeLike?.data?.data || nodeLike?.data || nodeLike || {};
  }

  function renderChart() {
    els.chart.innerHTML = "";
    if (!state.finalData.length) {
      state.chart = null;
      return;
    }

    const colorState = makeColorState();

    if (!OrgChartCtor) {
      showHint("Visualiseringsbiblioteket kunde inte laddas. Kontrollera nätverk/brandvägg och ladda om sidan.");
      return;
    }

    let c;
    try {
      c = new OrgChartCtor();
    } catch {
      c = OrgChartCtor();
    }

    if (!c) {
      showHint("Kunde inte initiera visualiseringsbiblioteket för organisationsschemat.");
      return;
    }

    c.container(`#${els.chart.id}`);
    c.data(state.finalData);
    c.nodeId((d) => d.id);
    c.parentNodeId((d) => d.parentId);
    callIfFn(c, "connections", state.connections || []);

    callIfFn(c, "svgHeight", 640);
    callIfFn(c, "nodeWidth", () => 260);
    callIfFn(c, "nodeHeight", () => 150);
    callIfFn(c, "childrenMargin", () => 60);
    callIfFn(c, "compactMarginBetween", () => 45);
    callIfFn(c, "compactMarginPair", () => 30);

    callIfFn(c, "linkUpdate", function (d) {
      const child = d.data?.data || d.data;
      const cc = colorState.colorFor(child);
      if (!d3Api?.select) return;
      d3Api.select(this).attr("stroke", cc?.stripe || "#CBD5E1").attr("stroke-width", 2).attr("stroke-opacity", 0.55);
    });

    callIfFn(c, "buttonContent", ({ node }) => {
      const row = unwrapNodeData(node);
      const isExpanded = !!node.children;
      const cnt = node.data?._directSubordinates ?? node.data?._totalSubordinates ?? "";
      const cc = colorState.colorFor(row);
      return `
        <div style="
          display:flex;align-items:center;gap:8px;
          padding:6px 10px;border-radius:999px;
          border:1px solid ${cc.stripe};background:white;
          box-shadow:0 10px 24px rgba(15,23,42,.10);
          color:#0f172a;font-size:12px;font-weight:800;
        ">
          <span style="
            width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;
            border-radius:999px;background:${cc.soft};border:1px solid ${cc.stripe};
          ">${isExpanded ? "−" : "+"}</span>
          <span style="opacity:.9">${cnt || ""}</span>
        </div>
      `;
    });

    callIfFn(c, "nodeContent", (d) => {
      const row = unwrapNodeData(d);
      const cc = colorState.colorFor(row);

      const platsLine = row.plats
        ? `<div style="font-size:12px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(
            row.plats
          )}</div>`
        : "";

      const trafikLine = row.trafik
        ? `<div style="font-size:12px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(
            row.trafik
          )}</div>`
        : "";

      const bolagFooter = row.bolag
        ? `<div style="
              margin-top:auto;font-size:11px;color:${cc.stripe};
              display:flex;justify-content:flex-end;
              white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
            ">${escapeHtml(row.bolag)}</div>`
        : `<div style="margin-top:auto;font-size:11px;">&nbsp;</div>`;

      return `
        <div style="
          width:${d.width}px;height:${d.height}px;
          position:relative;border-radius:18px;
          border:1px solid rgba(15,23,42,.12);
          background:white;box-shadow:0 16px 40px rgba(15,23,42,.10);
          overflow:hidden;
        ">
          <div style="position:absolute;inset:0 auto 0 0;width:10px;background:${cc.stripe};"></div>

          <div style="position:absolute;left:12px;right:12px;top:12px;bottom:12px;display:flex;flex-direction:column;gap:6px;">
            <div style="font-weight:900;font-size:14px;color:#0f172a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
              ${escapeHtml(row.name)}
            </div>

            <div style="font-size:12px;color:#475569;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
              ${escapeHtml(row.title)}
            </div>

            ${platsLine}
            ${trafikLine}

            ${bolagFooter}
          </div>

          <div style="position:absolute;top:-6px;left:50%;transform:translateX(-50%);width:12px;height:12px;border-radius:999px;background:white;border:2px solid rgba(15,23,42,.16);"></div>
          <div style="position:absolute;bottom:-6px;left:50%;transform:translateX(-50%);width:12px;height:12px;border-radius:999px;background:white;border:2px solid rgba(15,23,42,.16);"></div>
        </div>
      `;
    });

    callIfFn(c, "onNodeClick", (d) => openDetails(unwrapNodeData(d)));

    try {
      c.render();
      state.chart = c;
      setTimeout(() => {
        try {
          c.fit();
        } catch {}
      }, 0);
    } catch (error) {
      state.chart = null;
      setRenderError(error);
      showHint("Kunde inte rita organisationsschemat med den inlästa datan.");
    }
  }

  function openDetails(row) {
    state.selected = row;
    state.showAllSam = false;
    state.showFullDesc = false;
    els.detailsIdPill.textContent = row?.id || "";
    renderDetails();
    els.overlay.setAttribute("aria-hidden", "false");
  }

  function closeDetails() {
    state.selected = null;
    els.overlay.setAttribute("aria-hidden", "true");
    els.detailsBody.innerHTML = "";
    els.detailsIdPill.textContent = "";
  }

  function clampText(text, lines) {
    const safe = escapeHtml(text || "-");
    return `<div class="v" style="
      white-space:pre-wrap;
      overflow:hidden;
      display:-webkit-box;
      -webkit-line-clamp:${lines};
      -webkit-box-orient:vertical;
    ">${safe}</div>`;
  }

  function renderDetails() {
    const r = state.selected;
    if (!r) return;

    const sam = Array.isArray(r.sam) ? r.sam : [];
    const ansvar = Array.isArray(r.ansvar) ? r.ansvar : [];

    const desc = r.arbetsbeskrivning || "";
    const descNeedsToggle = desc.length > 220;
    const samNeedsToggle = sam.length > 4;

    const descHtml = state.showFullDesc
      ? `<div class="v" style="white-space:pre-wrap;">${escapeHtml(desc || "-")}</div>`
      : clampText(desc || "-", 5);

    const samList = state.showAllSam || !samNeedsToggle ? sam : sam.slice(0, 4);

    els.detailsBody.innerHTML = `
      <div class="kv"><div class="k">Namn</div><div class="v" style="font-weight:800;">${escapeHtml(r.name)}</div></div>
      <div class="kv"><div class="k">Titel</div><div class="v">${escapeHtml(r.title)}</div></div>
      <div class="kv"><div class="k">Bolag</div><div class="v">${escapeHtml(r.bolag || "-")}</div></div>
      <div class="kv"><div class="k">Plats / Nivå</div><div class="v">${escapeHtml(r.plats || "-")}</div></div>
      <div class="kv"><div class="k">Trafikområde</div><div class="v">${escapeHtml(r.trafik || "-")}</div></div>

      <div class="kv">
        <div class="k">Ansvarsområde</div>
        <div class="badges">
          ${
            ansvar.length ? ansvar.map((a) => `<span class="badge">${escapeHtml(a)}</span>`).join("") : `<span class="small">-</span>`
          }
        </div>
      </div>

      <div class="kv">
        <div class="toggleRow">
          <div class="k">Arbetsbeskrivning</div>
          ${descNeedsToggle ? `<button id="toggleDesc" class="toggleBtn">${state.showFullDesc ? "Visa mindre" : "Visa mer"}</button>` : ""}
        </div>
        ${descHtml}
      </div>

      <div class="kv">
        <div class="toggleRow">
          <div class="k">Fördelning SAM</div>
          ${
            samNeedsToggle
              ? `<button id="toggleSam" class="toggleBtn">${state.showAllSam ? "Visa mindre" : `Visa alla (${sam.length})`}</button>`
              : ""
          }
        </div>
        <ul class="list">
          ${
            samList.length
              ? samList.map((s) => `<li>${escapeHtml(s)}</li>`).join("")
              : `<li class="small" style="list-style:none;">-</li>`
          }
        </ul>
        ${!state.showAllSam && samNeedsToggle ? `<div class="small">Visar 4 av ${sam.length}</div>` : ""}
      </div>

      <div class="small">ID: <code>${escapeHtml(r.id)}</code></div>
    `;

    const tDesc = document.getElementById("toggleDesc");
    if (tDesc) tDesc.addEventListener("click", () => ((state.showFullDesc = !state.showFullDesc), renderDetails()));

    const tSam = document.getElementById("toggleSam");
    if (tSam) tSam.addEventListener("click", () => ((state.showAllSam = !state.showAllSam), renderDetails()));
  }

  async function exportPNG() {
    if (!htmlToImageApi?.toPng) {
      showHint("Exportbiblioteket kunde inte laddas. Kontrollera nätverk/brandvägg och ladda om sidan.");
      return;
    }
    const node = els.chart;
    const dataUrl = await htmlToImageApi.toPng(node, { pixelRatio: 2, cacheBust: true });
    const a = document.createElement("a");
    a.download = `organisationsschema_${ts()}.png`;
    a.href = dataUrl;
    a.click();
  }

  async function exportPDF() {
    if (!htmlToImageApi?.toPng || !jsPdfApi?.jsPDF) {
      showHint("PDF-export kunde inte initieras eftersom nödvändiga bibliotek saknas.");
      return;
    }
    const node = els.chart;
    const dataUrl = await htmlToImageApi.toPng(node, { pixelRatio: 2, cacheBust: true });

    const img = new Image();
    img.src = dataUrl;
    await img.decode();

    const { jsPDF } = jsPdfApi;
    const pdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    const margin = 18;
    const maxW = pageW - margin * 2;
    const maxH = pageH - margin * 2;

    const ratio = Math.min(maxW / img.width, maxH / img.height);
    const w = img.width * ratio;
    const h = img.height * ratio;
    const x = (pageW - w) / 2;
    const y = (pageH - h) / 2;

    pdf.addImage(dataUrl, "PNG", x, y, w, h, undefined, "FAST");
    pdf.save(`organisationsschema_${ts()}.pdf`);
  }

  function refresh() {
    applyFilters();
    state.finalData = keepAncestors();
    renderChart();
  }

  function showHint(message) {
    els.hint.textContent = message;
    els.hint.style.display = "block";
  }

  function parseRows(rawRows) {
    return (rawRows || [])
      .map((r) => {
        const parentRaw = getField(
          r,

          "Rapporterar till (Roll)",
          "Rapporterar till (Roll-ID)",
          "Rapporterar till Roll",
          "Rapporterar till Roll-ID",
          "Rapporterar till (Roll ID)",
          "Rapporterar till",
          "Rapporterar till ID",
          "Parent",
          "ParentID",
          "ParentId",
          "Parent Role ID",
          "Parent ID"
        );

        const parentIds = Array.from(
          new Set(parseMaybeList(parentRaw).map((x) => sanitizeIdentifier(x)).filter(Boolean))
        );

        return {
          id: sanitizeIdentifier(
            getField(r, "Roll-ID", "Roll ID", "RollID", "ID", "Role-ID", "role-id", "Role ID")
          ),
          parentId: parentIds[0] || null,
          parentIds,

          name:
            norm(getField(r, "Namn", "Name", "Person", "Personnamn")) ||
            norm(getField(r, "Manuellt namn", "Manual Name")) ||
            "(saknar namn)",
          title: norm(getField(r, "Roll - Titel", "Rolltitel", "Titel", "Role Title", "Title")) || "(saknar titel)",
          bolag: norm(getField(r, "Bolag")),
          plats: norm(getField(r, "Plats / Nivå")),
          trafik: norm(getField(r, "Trafikområde")),
          ansvar: parseMaybeList(getField(r, "Ansvarsområde")),
          sam: parseMaybeList(getField(r, "Fördelning SAM")),
          arbetsbeskrivning: norm(getField(r, "Arbetsbeskrivning")),
        };
      })
      .filter((d) => d.id);
  }

  function sanitizeHierarchyRows(rows) {
    const issues = [];
    const uniqueRows = [];
    const byIdInitial = new Map();
    const connections = [];

    function makeFingerprint(row) {
      const ansvar = Array.isArray(row.ansvar) ? row.ansvar.join("|") : "";
      const sam = Array.isArray(row.sam) ? row.sam.join("|") : "";
      return [row.name, row.title, row.bolag, row.plats, row.trafik, ansvar, sam, row.arbetsbeskrivning]
        .map((v) => norm(v).toLowerCase())
        .join("::");
    }

    function mergeUniqueValues(...lists) {
      return Array.from(new Set(lists.flat().filter(Boolean)));
    }

    (rows || []).forEach((row) => {
      if (!row || !row.id) return;
      const parentIds = Array.isArray(row.parentIds) ? row.parentIds : row.parentId ? [row.parentId] : [];
      const existing = byIdInitial.get(row.id);
      if (existing) {
        existing.parentIds = mergeUniqueValues(existing.parentIds || [], parentIds);
        if (!existing.name && row.name) existing.name = row.name;
        if (!existing.title && row.title) existing.title = row.title;
        if (!existing.bolag && row.bolag) existing.bolag = row.bolag;
        if (!existing.plats && row.plats) existing.plats = row.plats;
        if (!existing.trafik && row.trafik) existing.trafik = row.trafik;
        if (!existing.arbetsbeskrivning && row.arbetsbeskrivning) existing.arbetsbeskrivning = row.arbetsbeskrivning;
        existing.ansvar = mergeUniqueValues(existing.ansvar || [], row.ansvar || []);
        existing.sam = mergeUniqueValues(existing.sam || [], row.sam || []);
        issues.push(`Dubblett av Roll-ID '${row.id}' slogs ihop till en nod.`);
        return;
      }

      const merged = {
        ...row,
        parentIds: mergeUniqueValues(parentIds),
      };
      byIdInitial.set(row.id, merged);
      uniqueRows.push(merged);
    });

    const aliasById = new Map();
    const canonicalByFingerprint = new Map();
    const mergedRows = [];

    uniqueRows.forEach((row) => {
      const key = makeFingerprint(row);
      if (!key || key === ":::::::") {
        aliasById.set(row.id, row.id);
        mergedRows.push(row);
        return;
      }

      const existing = canonicalByFingerprint.get(key);
      if (!existing) {
        canonicalByFingerprint.set(key, row);
        aliasById.set(row.id, row.id);
        mergedRows.push(row);
        return;
      }

      aliasById.set(row.id, existing.id);
      existing.parentIds = mergeUniqueValues(existing.parentIds || [], row.parentIds || []);
      existing.ansvar = mergeUniqueValues(existing.ansvar || [], row.ansvar || []);
      existing.sam = mergeUniqueValues(existing.sam || [], row.sam || []);
      issues.push(
        `Roll-ID '${row.id}' hade samma innehåll som '${existing.id}' och slogs ihop för att undvika visuell duplicering.`
      );
    });

    mergedRows.forEach((row) => {
      row.parentIds = (row.parentIds || []).map((parentId) => aliasById.get(parentId) || parentId);
    });

    const byId = new Map(mergedRows.map((row) => [row.id, row]));

    mergedRows.forEach((row) => {
      const parentIds = Array.isArray(row.parentIds) ? row.parentIds.slice() : row.parentId ? [row.parentId] : [];

      if (row.parentId && row.parentId === row.id) {
        issues.push(`Roll-ID '${row.id}' rapporterade till sig själv. Länken togs bort.`);
      }

      parentIds.forEach((parentId) => {
        if (parentId !== row.id && !byId.has(parentId)) {
          issues.push(`Roll-ID '${row.id}' rapporterade till okänd chef ('${parentId}'). Länken togs bort.`);
        }
      });

      const validParents = parentIds.filter((parentId) => parentId !== row.id && byId.has(parentId));
      row.parentIds = Array.from(new Set(validParents));
      row.parentId = row.parentIds[0] || null;

      if (row.parentIds.length > 1) {
        row.parentIds.slice(1).forEach((parentId) => {
          connections.push({ from: parentId, to: row.id });
        });

        issues.push(
          `Roll-ID '${row.id}' rapporterade till flera överordnade (${row.parentIds.join(", ")}). Extra överordnade ritas som kopplingslinjer utan att noden dupliceras.`
        );
      }
    });

    const visitState = new Map();
    const pathStack = [];

    function walk(id) {
      const stateValue = visitState.get(id) || 0;
      if (stateValue === 2) return;
      if (stateValue === 1) {
        const cycleStart = pathStack.indexOf(id);
        const cycle = cycleStart >= 0 ? pathStack.slice(cycleStart).concat(id) : [id, id];
        const breakNodeId = cycle[0];
        const breakNode = byId.get(breakNodeId);
        if (breakNode && breakNode.parentId) {
          issues.push(`Cykel hittades (${cycle.join(" → ")}). Länken för '${breakNodeId}' togs bort.`);
          breakNode.parentId = null;
          breakNode.parentIds = [];
        }
        return;
      }

      visitState.set(id, 1);
      pathStack.push(id);

      const node = byId.get(id);
      if (node?.parentId && byId.has(node.parentId)) walk(node.parentId);

      pathStack.pop();
      visitState.set(id, 2);
    }

    Array.from(byId.keys()).forEach((id) => walk(id));

    const hasRoot = mergedRows.some((row) => !row.parentId);
    if (!hasRoot && mergedRows.length) {
      const forcedRoot = mergedRows[0];
      issues.push(`Ingen rotnod hittades i hierarkin. Länken för '${forcedRoot.id}' togs bort så att schemat kan ritas.`);
      forcedRoot.parentId = null;
      forcedRoot.parentIds = [];
    }

    const roots = mergedRows.filter((row) => !row.parentId);
    if (roots.length > 1) {
      let syntheticId = "__virtual_root__";
      while (byId.has(syntheticId)) syntheticId += "_x";

      roots.forEach((row) => {
        row.parentId = syntheticId;
      });

      mergedRows.unshift({
        id: syntheticId,
        parentId: null,
        parentIds: [],
        name: "Organisation",
        title: "Automatisk rotnod",
        bolag: "",
        plats: "",
        trafik: "",
        ansvar: [],
        sam: [],
        arbetsbeskrivning: "",
      });

      issues.push(`Flera rotnoder hittades (${roots.length}). En virtuell rotnod lades till så att schemat kan ritas.`);
    }

    return {
      rows: mergedRows,
      issues,
      connections,
    };
  }

  function parseCsvWithFallback(file, onDone) {
    if (!PapaApi?.parse) {
      onDone({ data: [], errors: [{ code: "ParserLibraryMissing" }], parsedRows: [] });
      return;
    }

    const delimiters = ["", ";", ",", "\t"];
    let index = 0;

    function run() {
      const delimiter = delimiters[index++];
      PapaApi.parse(file, {
        header: true,
        delimiter,
        skipEmptyLines: true,
        complete: (res) => {
          const rows = parseRows(res.data);
          const hasFatalErrors = (res.errors || []).some((e) => e.code === "UndetectableDelimiter");

          if (rows.length || index >= delimiters.length || !hasFatalErrors) {
            onDone({ ...res, parsedRows: rows });
            return;
          }

          run();
        },
        error: () => {
          if (index < delimiters.length) return run();
          onDone({ data: [], errors: [{ code: "ReadError" }], parsedRows: [] });
        },
      });
    }

    run();
  }

  function setDebugPayload(payload) {
    state.debug = payload;
    if (!els.debugOutput) return;
    els.debugOutput.textContent = JSON.stringify(payload, null, 2);
  }

  function buildDebugInfo({
    fileName = "",
    parseResult = {},
    parsedRows = [],
    sanitizedRows = [],
    issues = [],
    renderError = null,
  }) {
    const errorList = (parseResult.errors || []).map((e) => ({
      code: e.code || "Unknown",
      message: e.message || "",
      row: e.row,
    }));

    const byId = new Set(sanitizedRows.map((r) => r.id));
    const unresolvedParents = sanitizedRows
      .filter((r) => r.parentId && !byId.has(r.parentId))
      .map((r) => ({ id: r.id, parentId: r.parentId }));

    return {
      timestamp: new Date().toISOString(),
      fileName,
      counts: {
        rawRows: parseResult.data?.length || 0,
        parsedRows: parsedRows.length,
        sanitizedRows: sanitizedRows.length,
        issues: issues.length,
        parserErrors: errorList.length,
      },
      roots: sanitizedRows.filter((r) => !r.parentId).map((r) => r.id),
      parserErrors: errorList,
      issues,
      unresolvedParents,
      sampleIds: sanitizedRows.slice(0, 10).map((r) => r.id),
      renderError,
    };
  }

  function setRenderError(error) {
    if (!state.debug) return;
    setDebugPayload({
      ...state.debug,
      renderError: {
        message: error?.message || "Okänt renderingsfel",
        stack: error?.stack || null,
      },
    });
  }

  // Events
  els.fileInput.addEventListener("click", () => {
    // Allow selecting the same file again to re-trigger parsing.
    els.fileInput.value = "";
  });

  els.fileInput.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    setDebugPayload({
      timestamp: new Date().toISOString(),
      fileName: file.name,
      info: "CSV vald. Läser in…",
    });

    try {
      parseCsvWithFallback(file, (res) => {
        const parsedRows = res.parsedRows || [];
        const { rows, issues, connections } = sanitizeHierarchyRows(parsedRows);

        if (!rows.length) {
          setDebugPayload(
            buildDebugInfo({
              fileName: file.name,
              parseResult: res,
              parsedRows,
              sanitizedRows: rows,
              issues,
            })
          );
          showHint(
            "Inga giltiga rader hittades i CSV-filen. Kontrollera att kolumnen 'Roll-ID' finns och att separatorn är korrekt (komma eller semikolon)."
          );
          setControlsEnabled(false);
          state.data = [];
          state.filtered = [];
          state.finalData = [];
          state.connections = [];
          renderChart();
          closeDetails();
          return;
        }

        state.data = rows;
        state.connections = connections;
        state.q = "";
        state.bolag = "all";
        state.plats = "all";
        state.trafik = "all";

        els.searchInput.value = "";
        els.filterBolag.value = "all";
        els.filterPlats.value = "all";
        els.filterTrafik.value = "all";

        els.hint.style.display = "none";
        if ((res.errors && res.errors.length) || issues.length) {
          const details = issues.slice(0, 2).join(" ");
          const more = issues.length > 2 ? ` (+${issues.length - 2} till)` : "";
          showHint(
            `CSV lästes in, men vissa rader behövde justeras för att kunna ritas. ${details}${more}`.trim()
          );
        }

        setDebugPayload(
          buildDebugInfo({
            fileName: file.name,
            parseResult: res,
            parsedRows,
            sanitizedRows: rows,
            issues,
          })
        );

        setControlsEnabled(true);
        rebuildOptions();
        refresh();
        closeDetails();
      });
    } catch (error) {
      setDebugPayload({
        timestamp: new Date().toISOString(),
        fileName: file.name,
        info: "CSV kunde inte läsas in.",
        error: {
          message: error?.message || "Okänt fel",
          stack: error?.stack || null,
        },
      });
      showHint("CSV kunde inte läsas in. Kontrollera filen och ladda om sidan.");
    }
  });

  els.searchInput.addEventListener("input", (e) => {
    state.q = e.target.value || "";
    refresh();
  });
  els.filterBolag.addEventListener("change", (e) => {
    state.bolag = e.target.value;
    refresh();
  });
  els.filterPlats.addEventListener("change", (e) => {
    state.plats = e.target.value;
    refresh();
  });
  els.filterTrafik.addEventListener("change", (e) => {
    state.trafik = e.target.value;
    refresh();
  });

  els.colorBy.addEventListener("change", (e) => {
    state.colorBy = e.target.value;
    els.palette.disabled = state.colorBy === "none";
    refresh();
  });

  els.palette.addEventListener("change", (e) => {
    state.palette = e.target.value;
    refresh();
  });

  els.fitBtn.addEventListener("click", () => {
    try {
      state.chart && state.chart.fit();
    } catch {}
  });

  els.expandAllBtn.addEventListener("click", () => {
    const c = state.chart;
    if (!c) return;
    if (callIfFn(c, "expandAll")) {
      try {
        c.fit();
      } catch {}
    }
  });

  els.collapseAllBtn.addEventListener("click", () => {
    const c = state.chart;
    if (!c) return;
    if (callIfFn(c, "collapseAll")) {
      try {
        c.fit();
      } catch {}
    }
  });

  els.pngBtn.addEventListener("click", () => exportPNG());
  els.pdfBtn.addEventListener("click", () => exportPDF());

  els.closeDetails.addEventListener("click", () => closeDetails());
  els.overlay.addEventListener("click", (e) => {
    if (e.target === els.overlay) closeDetails();
  });

  // Start disabled until CSV loaded
  setControlsEnabled(false);
  setDebugPayload({ info: "Ingen CSV laddad ännu." });

  if (!PapaApi?.parse) {
    showHint("CSV-biblioteket kunde inte laddas. Kontrollera nätverk/brandvägg och ladda om sidan.");
  }
})();
