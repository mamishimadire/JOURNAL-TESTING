const state = {
	databases: [],
	tables: [],
	glColumns: [],
	tbColumns: [],
};

const PREF_THEME_KEY = "caats.theme";
const PREF_SOUND_KEY = "caats.clickSound";
let clickSoundEnabled = localStorage.getItem(PREF_SOUND_KEY) !== "off";
let audioCtx = null;

const ensureAudioContext = () => {
	if (!audioCtx) {
		const Ctx = window.AudioContext || window.webkitAudioContext;
		if (!Ctx) {
			return null;
		}
		audioCtx = new Ctx();
	}
	if (audioCtx.state === "suspended") {
		audioCtx.resume().catch(() => {});
	}
	return audioCtx;
};

const playClickSound = () => {
	if (!clickSoundEnabled) {
		return;
	}
	const ctx = ensureAudioContext();
	if (!ctx) {
		return;
	}
	const now = ctx.currentTime;
	const osc = ctx.createOscillator();
	const gain = ctx.createGain();
	osc.type = "triangle";
	osc.frequency.setValueAtTime(920, now);
	osc.frequency.exponentialRampToValueAtTime(620, now + 0.035);
	gain.gain.setValueAtTime(0.0001, now);
	gain.gain.linearRampToValueAtTime(0.06, now + 0.005);
	gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.05);
	osc.connect(gain);
	gain.connect(ctx.destination);
	osc.start(now);
	osc.stop(now + 0.055);
};

const $ = (id) => document.getElementById(id);
const NA_OPTION = "-- Not Available --";

const applyTheme = (theme) => {
	const isDark = theme === "dark";
	document.body.classList.toggle("theme-dark", isDark);
	const themeBtn = $("btnThemeToggle");
	if (themeBtn) {
		themeBtn.textContent = isDark ? "Switch to Light" : "Switch to Dark";
	}
	localStorage.setItem(PREF_THEME_KEY, isDark ? "dark" : "light");
};

const applySoundLabel = () => {
	const soundBtn = $("btnSoundToggle");
	if (soundBtn) {
		soundBtn.textContent = `Click Sound: ${clickSoundEnabled ? "On" : "Off"}`;
	}
};

const initUiPreferences = () => {
	const savedTheme = localStorage.getItem(PREF_THEME_KEY) || "light";
	applyTheme(savedTheme);
	applySoundLabel();

	$("btnThemeToggle")?.addEventListener("click", () => {
		const next = document.body.classList.contains("theme-dark") ? "light" : "dark";
		applyTheme(next);
	});

	$("btnSoundToggle")?.addEventListener("click", () => {
		clickSoundEnabled = !clickSoundEnabled;
		localStorage.setItem(PREF_SOUND_KEY, clickSoundEnabled ? "on" : "off");
		applySoundLabel();
	});

	document.addEventListener("click", (ev) => {
		const target = ev.target;
		if (!(target instanceof Element)) {
			return;
		}
		if (target.closest("button")) {
			playClickSound();
		}
	}, true);
};

const status = (msg) => {
	$("status").textContent = msg;
};

const initDefaultOutputFolder = async () => {
	const out = $("outFolder");
	if (!out) {
		return;
	}

	try {
		const res = await api("/api/caats/default-output-folder");
		if (res && typeof res.outputFolder === "string" && res.outputFolder.trim()) {
			out.value = res.outputFolder;
		}
	} catch {
		// Keep the existing fallback value in the input when endpoint lookup fails.
	}
};

const api = async (url, method = "GET", body = null) => {
	const res = await fetch(url, {
		method,
		headers: { "Content-Type": "application/json" },
		body: body ? JSON.stringify(body) : null,
	});

	const text = await res.text();
	let data = {};
	if (text) {
		try {
			data = JSON.parse(text);
		} catch {
			data = { message: text };
		}
	}
	if (!res.ok) {
		throw new Error(data.message || `Request failed (${res.status})`);
	}

	return data;
};

const fillSelect = (id, list) => {
	const sel = $(id);
	if (!sel) {
		return;
	}
	sel.innerHTML = "";
	list.forEach((x) => {
		const opt = document.createElement("option");
		opt.value = x;
		opt.textContent = x;
		sel.appendChild(opt);
	});
};

const fillSelectWithNa = (id, list) => {
	const sel = $(id);
	if (!sel) {
		return;
	}

	sel.innerHTML = "";
	const na = document.createElement("option");
	na.value = NA_OPTION;
	na.textContent = NA_OPTION;
	sel.appendChild(na);

	(list || []).forEach((x) => {
		const opt = document.createElement("option");
		opt.value = x;
		opt.textContent = x;
		sel.appendChild(opt);
	});

	sel.value = NA_OPTION;
};

const applyColumnOptions = (glColumns, tbColumns) => {
	const glIds = [
		"m_gl_account_number", "m_gl_posting_date", "m_gl_creation_date", "m_gl_journal_id", "m_gl_description",
		"m_gl_debit", "m_gl_credit", "m_gl_amount_signed", "m_gl_abs_amount", "m_gl_dc_indicator",
		"m_gl_user_id", "m_gl_period", "m_gl_year", "m_gl_line_id", "m_gl_jnl_origin",
		"e_group_col", "e_weekend_date_col", "e_gl_recon",
	];
	const tbIds = ["m_tb_account_number", "m_tb_account_name", "m_tb_opening_bal", "m_tb_closing_bal"];

	glIds.forEach((id) => fillSelectWithNa(id, glColumns));
	tbIds.forEach((id) => fillSelectWithNa(id, tbColumns));
};

const pickTableByHints = (tables, hints, fallbackIndex = 0) => {
	const arr = tables || [];
	const hit = arr.find((t) => hints.some((h) => String(t).toLowerCase().includes(h)));
	return hit || arr[Math.max(0, Math.min(fallbackIndex, arr.length - 1))] || "";
};

const pickByHints = (columns, hints) => {
	const cols = columns || [];
	for (const h of hints) {
		const found = cols.find((c) => String(c).toLowerCase() === h.toLowerCase());
		if (found) {
			return found;
		}
	}

	for (const h of hints) {
		const found = cols.find((c) => String(c).toLowerCase().includes(h.toLowerCase()));
		if (found) {
			return found;
		}
	}

	return "";
};

const autoMapFromCurrentColumns = () => {
	if (!state.glColumns.length || !state.tbColumns.length) {
		throw new Error("Preview tables first to load GL/TB columns.");
	}

	const gl = state.glColumns;
	const tb = state.tbColumns;

	const glMap = {
		account_number: pickByHints(gl, ["account_number", "account_as_tb", "hkont", "acctno", "gl_account"]),
		posting_date: pickByHints(gl, ["posted_dt", "doc_dt", "budat", "posting_date", "txndate"]),
		creation_date: pickByHints(gl, ["cpudt", "creation_date", "entrydate", "doc_dt"]),
		journal_id: pickByHints(gl, ["doc", "belnr", "journal_id", "jnl", "docno", "txn_no"]),
		description: pickByHints(gl, ["memo_description", "memo/description", "description", "memo", "bktxt"]),
		debit: pickByHints(gl, ["debit", "dr", "debit_amount"]),
		credit: pickByHints(gl, ["credit", "cr", "credit_amount"]),
		amount_signed: pickByHints(gl, ["signed_amount", "amount_signed", "balance", "amount", "dmbtr"]),
		abs_amount: pickByHints(gl, ["abs_amount", "absolute_amount", "amount_abs"]),
		dc_indicator: pickByHints(gl, ["shkzg", "dc", "d/c", "drcr"]),
		user_id: pickByHints(gl, ["usnam", "creator_id", "enteredby", "user", "posted by"]),
		period: pickByHints(gl, ["financial_period", "monat", "period", "prd"]),
		year: pickByHints(gl, ["financial_year", "gjahr", "year"]),
		line_id: pickByHints(gl, ["line_item_number", "line_id", "line", "item_no"]),
		jnl_origin: pickByHints(gl, ["jnl", "journal_origin", "jnl_origin", "origin"]),
	};

	const tbMap = {
		account_number: pickByHints(tb, ["account_number", "account", "gl_account"]),
		account_name: pickByHints(tb, ["account_name", "account", "description"]),
		opening_bal: pickByHints(tb, ["opening_balance", "opening_bal"]),
		closing_bal: pickByHints(tb, ["closing_balance", "closing_bal"]),
	};

	Object.entries(glMap).forEach(([k, v]) => {
		if (v) {
			const el = $(`m_gl_${k}`);
			if (el) {
				el.value = v;
			}
		}
	});

	Object.entries(tbMap).forEach(([k, v]) => {
		if (v) {
			const el = $(`m_tb_${k}`);
			if (el) {
				el.value = v;
			}
		}
	});

	if (glMap.journal_id) {
		$("e_group_col").value = glMap.journal_id;
	}

	if (glMap.posting_date) {
		$("e_weekend_date_col").value = glMap.posting_date;
	}

	const reconCol = pickByHints(gl, ["balance", "amount", "signed_amount", "dmbtr"]);
	if (reconCol) {
		$("e_gl_recon").value = reconCol;
	}

	const required = {
		glAccount: !!glMap.account_number,
		glPosting: !!glMap.posting_date,
		tbAccount: !!tbMap.account_number,
		tbClosing: !!tbMap.closing_bal,
	};

	const allRequiredMapped = required.glAccount && required.glPosting && required.tbAccount && required.tbClosing;
	const summary = $("mapSummary");
	if (summary) {
		summary.className = `summary-strip ${allRequiredMapped ? "success" : "warn"}`;
		summary.title = `GL Account: ${glMap.account_number || "Not Found"} | GL Posting Date: ${glMap.posting_date || "Not Found"} | TB Account: ${tbMap.account_number || "Not Found"} | TB Closing: ${tbMap.closing_bal || "Not Found"}`;
		summary.innerHTML = allRequiredMapped
			? `GL: <b>${Number(gl.length).toLocaleString()}</b> cols | TB: <b>${Number(tb.length).toLocaleString()}</b> cols -- auto-mapped ✅`
			: `GL: <b>${Number(gl.length).toLocaleString()}</b> cols | TB: <b>${Number(tb.length).toLocaleString()}</b> cols -- review required fields ⚠️`;
	}

	return required;
};

const fillSelectKeepFirst = (id, list) => {
	const sel = $(id);
	if (!sel) {
		return;
	}
	const first = sel.options.length > 0 ? sel.options[0].textContent : "All";
	sel.innerHTML = "";
	const allOpt = document.createElement("option");
	allOpt.value = first;
	allOpt.textContent = first;
	sel.appendChild(allOpt);
	list.filter((x) => x !== first).forEach((x) => {
		const opt = document.createElement("option");
		opt.value = x;
		opt.textContent = x;
		sel.appendChild(opt);
	});
};

const readMap = (prefix, keys) => {
	const map = {};
	keys.forEach((k) => {
		const el = $(`${prefix}_${k}`);
		if (!el) {
			return;
		}
		const v = String(el.value || "").trim();
		if (v && v !== NA_OPTION) {
			map[k] = v;
		}
	});
	return map;
};

const run = async (fn, trigger = null) => {
	if (trigger && trigger.dataset.busy === "1") {
		return;
	}

	if (trigger) {
		trigger.dataset.busy = "1";
		trigger.disabled = true;
	}

	try {
		await fn();
	} catch (err) {
		status(`Error: ${err.message}`);
	} finally {
		if (trigger) {
			trigger.dataset.busy = "0";
			trigger.disabled = false;
		}
	}
};

const esc = (v) => String(v ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const statusChip = (txt) => {
	const isWarn = String(txt).includes("⚠") || String(txt).toLowerCase().includes("variance");
	return `<span class="chip ${isWarn ? "warn" : "ok"}">${esc(txt)}</span>`;
};

const toCurrency = (v) => {
	const n = Number(v || 0);
	return `R${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
};

const renderTable = (containerId, headers, rows) => {
	const wrap = $(containerId);
	if (!rows || rows.length === 0) {
		wrap.innerHTML = "<div class='summary-strip'>No rows to display.</div>";
		return;
	}

	const headHtml = headers.map((h) => `<th>${esc(h)}</th>`).join("");
	const bodyHtml = rows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join("")}</tr>`).join("");
	wrap.innerHTML = `<table class="tbl"><thead><tr>${headHtml}</tr></thead><tbody>${bodyHtml}</tbody></table>`;
};

const previewCell = (value) => {
	if (value === null || value === undefined || String(value).trim() === "") {
		return "<span class='null-pill'>NULL</span>";
	}

	const raw = String(value);
	const clipped = raw.length > 90 ? `${raw.slice(0, 90)}...` : raw;
	return `<span title="${esc(raw)}">${esc(clipped)}</span>`;
};

const renderSqlPreview = (tableWrapId, tableTitleId, titleText, columns, rows) => {
	$(tableTitleId).textContent = titleText;
	const wrap = $(tableWrapId);
	if (!rows || rows.length === 0) {
		wrap.innerHTML = "<div class='summary-strip'>No sample rows returned.</div>";
		return;
	}

	const headers = (columns || []).map((c) => `<th>${esc(c)}</th>`).join("");
	const body = rows.map((row) => {
		const cells = (columns || []).map((c) => `<td>${previewCell(row[c])}</td>`).join("");
		return `<tr>${cells}</tr>`;
	}).join("");

	wrap.innerHTML = `<table class="tbl sql-preview"><thead><tr>${headers}</tr></thead><tbody>${body}</tbody></table>`;
};

const applyIndustryRules = () => {
	const industry = $("e_industry")?.value || "Standard (office hours)";
	const normalIndustries = ["Retail", "Hospitality", "Healthcare", "Manufacturing"];
	const isNormal = normalIndustries.some((k) => industry.includes(k));

	$("e_weekend_normal").checked = isNormal;
	$("e_holiday_normal").checked = isNormal;

	const note = $("industryNote");
	if (note) {
		note.innerHTML = isNormal
			? "Industry detected as 24/7 or shift-based: weekend and holiday journals are treated as NORMAL (informational)."
			: "Industry set to office-hours context: weekend and holiday journals are treated as EXCEPTIONS.";
	}
};

const toggleSngOrigins = () => {
	const wrap = $("sngOriginsWrap");
	if (!wrap) {
		return;
	}
	wrap.style.display = $("e_sng_rule").checked ? "flex" : "none";
};

const showAutoIndicatorGuide = () => {
	const host = $("autoIndicatorGuide");
	if (!host) {
		return;
	}

	const jnl = $("m_gl_jnl_origin")?.value || NA_OPTION;
	const sng = $("e_sng_rule")?.checked;

	if (sng) {
		host.className = "summary-strip success";
		host.innerHTML = jnl !== NA_OPTION
			? `SNG GT rule is ON: classification uses <b>JNL Origin</b> (<b>${esc(jnl)}</b>). Manual = origin in Manual Origins list; Automated = all other origins.`
			: "SNG GT rule is ON, but JNL Origin is not mapped yet. Map JNL Origin to classify Manual vs Automated by origin list.";
		return;
	}

	host.className = "summary-strip info";
	host.innerHTML = "SNG GT rule is OFF: all entries are treated as Manual (conservative mode).";
};

const isProcedureApplicable = (newId, legacyId = "") => {
	const sel = $(newId);
	if (sel && typeof sel.value === "string") {
		return sel.value === "Applicable";
	}

	const legacy = legacyId ? $(legacyId) : null;
	if (legacy && typeof legacy.checked === "boolean") {
		return legacy.checked;
	}

	return true;
};

const readProcedureScope = (name) => {
	const manual = $(`proc_mode_${name}_manual`);
	const semi = $(`proc_mode_${name}_semi`);
	const auto = $(`proc_mode_${name}_auto`);
	const manualVal = manual ? !!manual.checked : true;
	return {
		manual: manualVal,
		// If Semi-Auto control is not rendered, keep it aligned with Manual selection.
		semiAuto: semi ? !!semi.checked : manualVal,
		auto: auto ? !!auto.checked : true,
	};
};

const setProcedureScope = (name, manual, semiAuto, auto) => {
	const m = $(`proc_mode_${name}_manual`);
	const s = $(`proc_mode_${name}_semi`);
	const a = $(`proc_mode_${name}_auto`);
	if (m) {
		m.checked = manual;
	}
	if (s) {
		s.checked = semiAuto;
	}
	if (a) {
		a.checked = auto;
	}
};

const syncProcedureModesBySng = () => {
	const enabled = !!($("e_sng_rule")?.checked);
	const manualOnly = ["backdated", "holiday", "weekend", "round", "adj_desc", "low_fsli", "user"];
	const allTypes = ["completeness", "unbalanced", "above_mat", "duplicate", "benford"];

	if (enabled) {
		manualOnly.forEach((name) => setProcedureScope(name, true, false, false));
		allTypes.forEach((name) => setProcedureScope(name, true, true, true));
		return;
	}

	[...manualOnly, ...allTypes].forEach((name) => setProcedureScope(name, true, true, true));
};

const PROC_IMPACT = {
	proc_completeness: {
		app: "Will include full GL/TB reconciliation table and variance status in Step 5 + report Section 1.",
		not: "Will skip GL/TB reconciliation outputs in Step 5 metrics."
	},
	proc_backdated: {
		app: "Will include Year/Month backdated flags and months difference in results/report.",
		not: "Will skip backdated testing and hide related counts from summary."
	},
	proc_holiday: {
		app: "Will include manual journals on SA public holidays (or informational by industry mode).",
		not: "Will skip holiday journal testing and related outputs."
	},
	proc_round: {
		app: "Will include v14 round categories: 10,000 (.00) and 9,999 any decimal.",
		not: "Will skip round amount analysis outputs."
	},
	proc_weekend: {
		app: "Will include manual weekend journals (exception or informational by industry mode).",
		not: "Will skip weekend journal testing and related outputs."
	},
	proc_unbalanced: {
		app: "Will include journal-group unbalanced detection (Dr != Cr).",
		not: "Will skip unbalanced journal outputs."
	},
	proc_above_mat: {
		app: "Will include journals above performance materiality threshold.",
		not: "Will skip above-materiality outputs."
	},
	proc_adj_desc: {
		app: "Will include journals with adjustment/correction/reversal keywords.",
		not: "Will skip adjustment-description outputs."
	},
	proc_low_fsli: {
		app: "Will include low-frequency account/period postings (Low FSLI).",
		not: "Will skip low FSLI outputs."
	},
	proc_duplicate: {
		app: "Will include duplicate journal line detection outputs.",
		not: "Will skip duplicate detection outputs."
	},
	proc_user: {
		app: "Will include user analysis summaries in output context.",
		not: "Will keep user analysis out of Step 5 applicability summary."
	},
	proc_benford: {
		app: "Will include Benford frequency table and outlier digit flags.",
		not: "Will skip Benford analysis outputs."
	}
};

let procImpactTimer;
const showProcedureImpact = (procId, applicable) => {
	const host = $("procImpact");
	if (!host) {
		return;
	}
	const cfg = PROC_IMPACT[procId];
	if (!cfg) {
		return;
	}
	clearTimeout(procImpactTimer);
	host.style.display = "block";
	host.className = `summary-strip ${applicable ? "success" : "warn"}`;
	host.innerHTML = applicable
		? `<b>Applicable:</b> ${esc(cfg.app)}`
		: `<b>Not Applicable:</b> ${esc(cfg.not)}`;
	procImpactTimer = setTimeout(() => {
		host.style.display = "none";
	}, 8000);
};

const wireProcedureImpact = () => {
	Object.keys(PROC_IMPACT).forEach((id) => {
		const el = $(id);
		if (!el) {
			return;
		}
		el.addEventListener("change", () => {
			showProcedureImpact(id, el.value === "Applicable");
		});
	});
};

$("btnConnect")?.addEventListener("click", () => run(async () => {
	status("Connecting...");
	const data = await api("/api/caats/connect", "POST", {
		server: $("server").value,
		driver: $("driver").value,
		trustedConnection: $("trusted").checked,
	});

	state.databases = data.databases || [];
	fillSelect("database", state.databases);
	status(data.message);
}, $("btnConnect")));

$("btnUseDb")?.addEventListener("click", () => run(async () => {
	const db = $("database").value;
	if (!db) {
		throw new Error("Select a database first.");
	}
	await api("/api/caats/database", "POST", { database: db });
	const tables = await api("/api/caats/tables");
	state.tables = tables;
	fillSelect("glTable", tables);
	fillSelect("tbTable", tables);
	$("glTable").value = pickTableByHints(tables, ["gl", "journal", "jnl", "ledger", "posting", "trans"], 0);
	$("tbTable").value = pickTableByHints(tables, ["tb", "trial", "balance", "coa", "account", "chart"], tables.length - 1);
	status(`Database ${db} selected. ${tables.length} tables loaded.`);
}, $("btnUseDb")));

$("btnPreview")?.addEventListener("click", () => run(async () => {
	if (!$("glTable").value || !$("tbTable").value) {
		throw new Error("Select GL and TB tables first.");
	}

	const started = performance.now();
	status("Loading table preview (GL + TB in parallel)...");
	const res = await api("/api/caats/preview", "POST", {
		glTable: $("glTable").value,
		tbTable: $("tbTable").value,
	});

	renderSqlPreview(
		"glPreviewTable",
		"glPreviewTitle",
		`GL: ${$("glTable").value} (${(res.glColumns || []).length} cols)` ,
		res.glColumns || [],
		res.glSample || []);

	renderSqlPreview(
		"tbPreviewTable",
		"tbPreviewTitle",
		`TB: ${$("tbTable").value} (${(res.tbColumns || []).length} cols)`,
		res.tbColumns || [],
		res.tbSample || []);

	state.glColumns = res.glColumns || [];
	state.tbColumns = res.tbColumns || [];
	applyColumnOptions(state.glColumns, state.tbColumns);
	autoMapFromCurrentColumns();

	const elapsedMs = Math.round(performance.now() - started);
	status(`Preview loaded in ${(elapsedMs / 1000).toFixed(1)}s.`);
}, $("btnPreview")));

$("btnAutoMap")?.addEventListener("click", () => run(async () => {
	status("Detecting column mapping from table schemas...");
	const required = autoMapFromCurrentColumns();
	if (!required.glAccount || !required.glPosting || !required.tbAccount || !required.tbClosing) {
		status("Auto-detect finished with missing required fields. Please review highlighted mapping summary.");
		return;
	}

	status("Column auto-detect complete.");
}, $("btnAutoMap")));

$("btnSaveMap")?.addEventListener("click", () => run(async () => {
	const glKeys = ["account_number", "posting_date", "creation_date", "journal_id", "description", "debit", "credit", "amount_signed", "abs_amount", "dc_indicator", "user_id", "period", "year", "line_id", "jnl_origin"];
	const tbKeys = ["account_number", "account_name", "opening_bal", "closing_bal"];
	await api("/api/caats/mapping", "POST", {
		gl: readMap("m_gl", glKeys),
		tb: readMap("m_tb", tbKeys),
	});
	status("Column mapping saved.");
}, $("btnSaveMap")));

$("btnLoadData")?.addEventListener("click", () => run(async () => {
	const started = performance.now();
	status("Loading full GL/TB data (this may take a while for large tables)...");
	const res = await api("/api/caats/load-data", "POST", {});
	const elapsedMs = Math.round(performance.now() - started);
	status(`${res.message} GL rows: ${res.glRows.toLocaleString()} | TB rows: ${res.tbRows.toLocaleString()} | Time: ${(elapsedMs / 1000).toFixed(1)}s`);
}, $("btnLoadData")));

$("btnSaveEng")?.addEventListener("click", () => run(async () => {
	await api("/api/caats/engagement", "POST", {
		parent: $("e_parent").value,
		client: $("e_client").value,
		engagement: $("e_engagement").value,
		glSystem: $("e_gl_system").value,
		industry: $("e_industry").value,
		sngOriginsRaw: $("e_sng_origins").value,
		journalGroupColumn: $("e_group_col").value === NA_OPTION ? "" : $("e_group_col").value,
		weekendDateColumn: $("e_weekend_date_col")?.value === NA_OPTION ? "" : $("e_weekend_date_col")?.value || "",
		glReconAmountColumn: $("e_gl_recon").value === NA_OPTION ? "" : $("e_gl_recon").value,
		fyStart: $("e_fy_start").value || null,
		fyEnd: $("e_fy_end").value || null,
		periodStart: $("e_period_start").value || null,
		periodEnd: $("e_period_end").value || null,
		signingDate: $("e_signing_date").value || null,
		materiality: Number($("e_materiality").value || "0"),
		performanceMateriality: Number($("e_perf_mat").value || "0"),
		minJournalValue: Number($("e_min_jv").value || "0"),
		lowFsliThreshold: Number($("e_low_fsli").value || "10"),
		auditor: $("e_auditor").value,
		manager: $("e_manager").value,
		countryCode: $("e_country").value,
		weekendNormal: $("e_weekend_normal").checked,
		holidayNormal: $("e_holiday_normal").checked,
		sngRule: $("e_sng_rule").checked,
	});
	status("Engagement settings saved.");
}, $("btnSaveEng")));

$("btnAutoDetect")?.addEventListener("click", () => run(async () => {
	status("Detecting GL profile...");
	const res = await api("/api/caats/profile", "POST", {});

	if (res.minDate) {
		const minDate = String(res.minDate).slice(0, 10);
		$("e_fy_start").value = minDate;
		$("e_period_start").value = minDate;
	}

	if (res.maxDate) {
		const maxDate = String(res.maxDate).slice(0, 10);
		$("e_fy_end").value = maxDate;
		$("e_period_end").value = maxDate;
	}

	if (res.dateColumnUsed && (!$("e_weekend_date_col").value || $("e_weekend_date_col").value === "CPUDT" || $("e_weekend_date_col").value === NA_OPTION)) {
		$("e_weekend_date_col").value = res.dateColumnUsed;
	}

	$("e_weekend_normal").checked = !!res.weekendNormalSuggested;
	$("e_holiday_normal").checked = !!res.weekendNormalSuggested;
	if (res.suggestedIndustry) {
		$("e_industry").value = res.suggestedIndustry;
	}
	applyIndustryRules();

	$("profileSummary").innerHTML = `FY: <b>${res.fyYear ?? "N/A"}</b> | Date Column: <b>${esc(res.dateColumnUsed || "N/A")}</b> | Range: <b>${res.minDate ? String(res.minDate).slice(0, 10) : "N/A"}</b> to <b>${res.maxDate ? String(res.maxDate).slice(0, 10) : "N/A"}</b> | Weekend %: <b>${Number(res.weekendPercent || 0).toFixed(1)}%</b> | Structure: <b>${esc(res.amountStructure || "unknown")}</b> | Rows: <b>${Number(res.totalRows || 0).toLocaleString()}</b>`;
	status("Auto-detect complete.");
}, $("btnAutoDetect")));

$("e_industry")?.addEventListener("change", applyIndustryRules);
$("e_sng_rule")?.addEventListener("change", toggleSngOrigins);
$("e_sng_rule")?.addEventListener("change", showAutoIndicatorGuide);
$("e_sng_rule")?.addEventListener("change", syncProcedureModesBySng);
$("m_gl_jnl_origin")?.addEventListener("change", showAutoIndicatorGuide);
applyIndustryRules();
toggleSngOrigins();
showAutoIndicatorGuide();
syncProcedureModesBySng();
wireProcedureImpact();
initUiPreferences();
initDefaultOutputFolder();

$("btnRun")?.addEventListener("click", () => run(async () => {
	status("Running CAATS tests...");
	const res = await api("/api/caats/run", "POST", {
		procedures: {
			completeness: isProcedureApplicable("proc_completeness", "p_completeness"),
			backdated: isProcedureApplicable("proc_backdated", "p_backdated"),
			holiday: isProcedureApplicable("proc_holiday", "p_holiday"),
			round: isProcedureApplicable("proc_round", "p_round"),
			weekend: isProcedureApplicable("proc_weekend", "p_weekend"),
			unbalanced: isProcedureApplicable("proc_unbalanced", "p_unbalanced"),
			above_mat: isProcedureApplicable("proc_above_mat", "p_above_mat"),
			adj_desc: isProcedureApplicable("proc_adj_desc", "p_adj_desc"),
			low_fsli: isProcedureApplicable("proc_low_fsli", "p_low_fsli"),
			duplicate: isProcedureApplicable("proc_duplicate", "p_duplicate"),
			benford: isProcedureApplicable("proc_benford", "p_benford"),
			user: isProcedureApplicable("proc_user", "p_user"),
		},
		procedureScopes: {
			completeness: readProcedureScope("completeness"),
			backdated: readProcedureScope("backdated"),
			holiday: readProcedureScope("holiday"),
			round: readProcedureScope("round"),
			weekend: readProcedureScope("weekend"),
			unbalanced: readProcedureScope("unbalanced"),
			above_mat: readProcedureScope("above_mat"),
			adj_desc: readProcedureScope("adj_desc"),
			low_fsli: readProcedureScope("low_fsli"),
			duplicate: readProcedureScope("duplicate"),
			user: readProcedureScope("user"),
			benford: readProcedureScope("benford"),
		},
	});

	const cards = $("cards");
	cards.innerHTML = "";
	Object.entries(res.counts).forEach(([k, v]) => {
		const el = document.createElement("div");
		el.className = "metric";
		const value = v < 0 ? "N/A" : Number(v).toLocaleString();
		el.innerHTML = `<div class="name">${k}</div><div class="value">${value}</div>`;
		cards.appendChild(el);
	});

	const reconRows = Number(res.reconTotalRows || 0);
	$("runSummary").innerHTML = `Recon Rows: <b>${Number(reconRows).toLocaleString()}</b> | Recon Agrees: <b>${res.reconAgree < 0 ? "N/A" : Number(res.reconAgree).toLocaleString()}</b> | Recon Variance: <b>${res.reconVariance < 0 ? "N/A" : Number(res.reconVariance).toLocaleString()}</b> | Total Debit: <b>${toCurrency(res.totalDebit)}</b> | Total Credit: <b>${toCurrency(res.totalCredit)}</b> | Benford Flagged Digits: <b>${res.benfordFlaggedDigits < 0 ? "N/A" : Number(res.benfordFlaggedDigits).toLocaleString()}</b>`;

	const holidayRows = res.holidayVerification || [];
	const holidaySummary = $("holidayVerificationSummary");
	if (holidaySummary) {
		const applicable = !!res.holidayProcedureApplicable;
		if (!holidayRows.length) {
			holidaySummary.className = "summary-strip info";
			holidaySummary.innerHTML = "No South African holiday matches were found in the current posting-date population.";
		} else {
			holidaySummary.className = applicable ? "summary-strip success" : "summary-strip warn";
			holidaySummary.innerHTML = applicable
				? `Matched <b>${Number(holidayRows.length).toLocaleString()}</b> holiday date/name pair(s). In-scope counts reflect your current Manual/Auto procedure matrix.`
				: `Matched <b>${Number(holidayRows.length).toLocaleString()}</b> holiday date/name pair(s). Holiday procedure is currently marked Not Applicable.`;
		}
	}

	renderTable("holidayVerificationWrap",
		["Check Date", "Holiday Name", "Matched Lines", "Matched Journals", "In-Scope Lines", "In-Scope Journals", "In-Scope Exception Lines", "In-Scope Exception Journals"],
		holidayRows.map((h) => [
			esc(h.checkDate ? String(h.checkDate).slice(0, 10) : ""),
			esc(h.holidayName),
			esc(Number(h.matchedLines || 0).toLocaleString()),
			esc(Number(h.matchedJournalGroups || 0).toLocaleString()),
			esc(Number(h.inScopeLines || 0).toLocaleString()),
			esc(Number(h.inScopeJournalGroups || 0).toLocaleString()),
			esc(Number(h.inScopeExceptionLines || 0).toLocaleString()),
			esc(Number(h.inScopeExceptionJournalGroups || 0).toLocaleString()),
		]));

	renderTable("reconTableWrap",
		["Account", "Name", "Opening", "Closing", "GL Debit", "GL Credit", "GL Balance", "Difference", "Status"],
		(res.reconRows || res.reconTop || []).map((r) => [
			esc(r.accountNumber),
			esc(r.accountName),
			toCurrency(r.openingBalance),
			toCurrency(r.closingBalance),
			toCurrency(r.glDebit),
			toCurrency(r.glCredit),
			toCurrency(r.glBalance),
			toCurrency(r.difference),
			statusChip(r.status),
		]));

	renderTable("benfordTableWrap",
		["Digit", "Count", "Actual %", "Expected %", "Diff %", "Status"],
		(res.benford || []).map((b) => [
			esc(b.leadingDigit),
			esc(Number(b.count).toLocaleString()),
			esc(`${Number(b.actualPercent).toFixed(1)}%`),
			esc(`${Number(b.expectedPercent).toFixed(1)}%`),
			esc(`${Number(b.differencePercent).toFixed(1)}%`),
			statusChip(b.status),
		]));

	renderTable("riskSummaryWrap",
		["Test", "Total Risk Score", "Lines Flagged"],
		(res.riskScoreSummary || []).map((r) => [
			esc(r.test),
			esc(Number(r.totalRiskScore || 0).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })),
			esc(Number(r.linesFlagged || 0).toLocaleString()),
		]));

	renderTable("userAnalysisFullWrap",
		["Full Name", "Manual/Automated", "Line Count Debit", "Total Amt Debit", "Line Count Credit", "Total Amt Credit", "Total Line Count", "Total Reporting Amount"],
		(res.userAnalysisFull || []).map((u) => [
			esc(u.fullName),
			esc(u.manualAutomatedDescriptor),
			esc(Number(u.lineCountDebit || 0).toLocaleString()),
			toCurrency(u.totalReportingAmountDebit),
			esc(Number(u.lineCountCredit || 0).toLocaleString()),
			toCurrency(u.totalReportingAmountCredit),
			esc(Number(u.totalLineCount || 0).toLocaleString()),
			toCurrency(u.totalReportingAmount),
		]));

	renderTable("procedureRankedWrap",
		["Rank", "Procedure", "Count", "Ranked Label"],
		(res.procedureResultsRanked || []).map((p) => [
			esc(p.rank),
			esc(p.procedure),
			esc(Number(p.count || 0).toLocaleString()),
			esc(p.rankedLabel),
		]));

	renderTable("iriRankedWrap",
		["Rank", "IRI", "Description", "Count", "Ranked Label"],
		(res.iriRankedSummary || []).map((i) => [
			esc(i.rank),
			esc(i.iri),
			esc(i.description),
			esc(Number(i.count || 0).toLocaleString()),
			esc(i.rankedLabel),
		]));

	renderTable("iriDetailWrap",
		["IRI", "Description", "Account", "Journal Key", "Doc ID", "Posting Date", "CPU Date", "User", "Debit", "Credit", "Abs", "Day", "Holiday", "Months Backdated", "Round Category"],
		(res.iriDetailRows || []).map((d) => [
			esc(d.iri),
			esc(d.description),
			esc(d.acct),
			esc(d.journalKey),
			esc(d.journalId),
			esc(d.postingDate ? String(d.postingDate).slice(0, 10) : ""),
			esc(d.cpuDate ? String(d.cpuDate).slice(0, 10) : ""),
			esc(d.user),
			toCurrency(d.debit),
			toCurrency(d.credit),
			toCurrency(d.absAmount),
			esc(d.dayName),
			esc(d.holidayName),
			esc(d.monthsBackdated),
			esc(d.roundCategory),
		]));

	fillSelectKeepFirst("x_user", res.users || ["All"]);
	fillSelectKeepFirst("x_account", res.accounts || ["All"]);
	fillSelectKeepFirst("x_journal", res.journalKeys || ["All"]);

	status(`${res.message} Showing first ${Number((res.reconRows || []).length).toLocaleString()} recon rows in table.`);
}, $("btnRun")));

$("btnExport")?.addEventListener("click", () => run(async () => {
	status("Generating reports...");
	const res = await api("/api/caats/export", "POST", {
		outputFolder: $("outFolder").value,
		word: $("expWord").checked,
		excel: $("expExcel").checked,
		csv: $("expCsv").checked,
	});

	const linksHost = $("exportLinks");
	if (linksHost) {
		const files = res.files || [];
		const links = files.map((f) => {
			const name = String(f).split(/[/\\]/).pop();
			const href = `/api/caats/download?path=${encodeURIComponent(f)}`;
			return `<a href="${href}" target="_blank" rel="noopener">${esc(name || f)}</a>`;
		});
		if (links.length) {
			linksHost.style.display = "block";
			linksHost.className = "summary-strip success";
			linksHost.innerHTML = `<b>Downloads:</b> ${links.join(" | ")}`;
		}
	}
	status(`${res.message} ${(res.files || []).join(" | ")}`);
}, $("btnExport")));

$("btnExplore")?.addEventListener("click", () => run(async () => {
	status("Applying explorer filter...");
	const res = await api("/api/caats/explore", "POST", {
		filter: $("x_filter").value,
		user: $("x_user").value,
		account: $("x_account").value,
		journalKey: $("x_journal").value,
		maxRows: Number($("x_rows").value || "100"),
	});

	$("exploreStats").innerHTML = `Matched <b>${Number(res.journals).toLocaleString()}</b> journals and <b>${Number(res.lines).toLocaleString()}</b> lines.`;

	renderTable("exploreTableWrap",
		["Account", "Journal Key", "Doc ID", "Posting Date", "CPU Date", "User", "Debit", "Credit", "Round Category", "Months Backdated"],
		(res.rows || []).map((r) => [
			esc(r.acct),
			esc(r.journalKey),
			esc(r.journalId),
			esc(r.postingDate ? String(r.postingDate).slice(0, 10) : ""),
			esc(r.cpuDate ? String(r.cpuDate).slice(0, 10) : ""),
			esc(r.user),
			toCurrency(r.debit),
			toCurrency(r.credit),
			esc(r.roundCategory),
			esc(r.monthsBackdated),
		]));

	status(`${res.message} ${Number(res.journals).toLocaleString()} journals | ${Number(res.lines).toLocaleString()} lines.`);
}, $("btnExplore")));
