import { useState, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

// ── Fallback sample data (only used if no files are uploaded) ────────────────
const SAMPLE_DATA = [
  {"orig_code":"143803","name":"ايمان ابراهيم / السنبلاوين","address":"محلات السكة الحديد امام ميدان ام كلثوم / السنبلاوين / دقهلية","province":"الدقهلية","district":"السنبلاوين","distributor":"El-Masreya"},
  {"orig_code":"53321","name":"ابوزيد اجا","address":"مينة سمنود الشارع العام","province":"الدقهلية","district":"اجا","distributor":"El-Masreya"},
  {"orig_code":"51153","name":"ريهام عنتر الجديدة -السنبلاوين","address":"المعاهدة / طريق كفر صقر / بجوار ورشة المكاوى للبلاط / السنبلاوين","province":"الدقهلية","district":"السنبلاوين","distributor":"El-Masreya"},
  {"orig_code":"1000016910","name":"ادارة صيدلية احمد يحيى/ 6 اكتوبر","address":"الاردنيه. الحى ال7 . اكتوبر, Egypt","province":"","district":"Hay 7 - الحى السابع","distributor":"Purex"},
  {"orig_code":"1000039628","name":"ادارة صيدليات بلبع / سيدى جابر","address":"محطه سيدي جابر بجوار صيدليه بخيت ، الاسكندريه, Egypt","province":"","district":"Sidi Gaber - سيدى جابر","distributor":"Purex"}
];

const BATCH_SIZE = 5;

// ── Styles ────────────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #0d1117;
    --surface: #161b22;
    --surface2: #1c2330;
    --border: #30363d;
    --accent: #00c896;
    --accent2: #0099ff;
    --masreya: #f59e0b;
    --purex: #3b82f6;
    --text: #e6edf3;
    --muted: #8b949e;
    --danger: #f85149;
    --success: #3fb950;
    --radius: 8px;
  }

  body { background: var(--bg); color: var(--text); font-family: 'IBM Plex Sans Arabic', sans-serif; direction: rtl; margin: 0; padding: 0; }

  .app { min-height: 100vh; display: flex; flex-direction: column; }

  .header {
    background: linear-gradient(135deg, #0d1117 0%, #161b22 100%);
    border-bottom: 1px solid var(--border);
    padding: 20px 28px;
    display: flex; align-items: center; gap: 16px;
    position: sticky; top: 0; z-index: 100;
  }
  .header-icon { width: 40px; height: 40px; background: var(--accent); border-radius: 10px;
    display: flex; align-items: center; justify-content: center; font-size: 20px; flex-shrink: 0; }
  .header h1 { font-size: 18px; font-weight: 700; color: var(--text); margin: 0; }
  .header p { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .header-stats { margin-right: auto; display: flex; gap: 16px; }
  .stat-pill { background: var(--surface2); border: 1px solid var(--border); border-radius: 20px;
    padding: 4px 12px; font-size: 12px; color: var(--muted); font-family: 'JetBrains Mono', monospace; }
  .stat-pill span { color: var(--accent); font-weight: 600; }

  .main { display: grid; grid-template-columns: 320px 1fr; gap: 0; flex: 1; }

  .sidebar { border-left: 1px solid var(--border); padding: 20px 16px; display: flex; flex-direction: column; gap: 12px;
    height: calc(100vh - 73px); position: sticky; top: 73px; overflow-y: auto; }
  .sidebar-section h3 { font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; color: var(--muted);
    margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }

  .btn { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 10px 16px;
    border-radius: var(--radius); border: none; cursor: pointer; font-family: inherit;
    font-size: 13px; font-weight: 600; transition: all 0.15s; width: 100%; }
  .btn-primary { background: var(--accent); color: #000; }
  .btn-primary:hover { background: #00e6aa; }
  .btn-primary:disabled { background: #1a3d30; color: #3d7a5f; cursor: not-allowed; }
  .btn-secondary { background: var(--surface2); color: var(--text); border: 1px solid var(--border); }
  .btn-secondary:hover { border-color: var(--accent); color: var(--accent); }
  .btn-danger { background: transparent; color: var(--danger); border: 1px solid var(--danger); }
  .btn-danger:hover { background: rgba(248,81,73,0.1); }

  .input-group { display: flex; flex-direction: column; gap: 4px; }
  .input-group label { font-size: 11px; color: var(--muted); }
  .input-group input, .input-group select {
    background: var(--surface2); border: 1px solid var(--border); border-radius: var(--radius);
    color: var(--text); padding: 8px 12px; font-size: 13px; font-family: inherit; width: 100%; }
  .input-group input:focus, .input-group select:focus { outline: none; border-color: var(--accent); }

  .progress-bar-wrap { background: var(--surface2); border-radius: 4px; height: 6px; overflow: hidden; }
  .progress-bar { height: 100%; background: linear-gradient(90deg, var(--accent), var(--accent2));
    transition: width 0.4s ease; border-radius: 4px; }

  .content { padding: 20px 24px; overflow-y: auto; height: calc(100vh - 73px); }

  .master-group {
    background: var(--surface); border: 1px solid var(--border); border-radius: 12px;
    margin-bottom: 16px; overflow: hidden; transition: border-color 0.2s;
  }
  .master-group:hover { border-color: #404a56; }
  .master-group.confirmed { border-color: var(--success); }
  .master-group.pending { border-color: var(--border); }

  .master-header {
    padding: 14px 18px; display: flex; align-items: center; gap: 12px;
    background: var(--surface2); cursor: pointer;
  }
  .master-code { font-family: 'JetBrains Mono', monospace; font-size: 12px;
    background: rgba(0,200,150,0.1); color: var(--accent); border: 1px solid rgba(0,200,150,0.3);
    border-radius: 4px; padding: 2px 8px; flex-shrink: 0; }
  .master-name { font-size: 15px; font-weight: 600; flex: 1; }
  .master-meta { font-size: 11px; color: var(--muted); }
  .master-badges { display: flex; gap: 6px; margin-right: auto; }
  .badge { font-size: 10px; padding: 2px 8px; border-radius: 10px; font-weight: 600; font-family: 'JetBrains Mono', monospace; }
  .badge-masreya { background: rgba(245,158,11,0.15); color: var(--masreya); border: 1px solid rgba(245,158,11,0.3); }
  .badge-purex { background: rgba(59,130,246,0.15); color: var(--purex); border: 1px solid rgba(59,130,246,0.3); }
  .badge-status-ok { background: rgba(63,185,80,0.15); color: var(--success); }
  .badge-status-pending { background: rgba(139,148,158,0.15); color: var(--muted); }

  .master-body { padding: 0 18px 14px; }

  .pharmacy-row {
    display: flex; align-items: flex-start; gap: 10px; padding: 10px 0;
    border-bottom: 1px solid var(--border);
  }
  .pharmacy-row:last-child { border-bottom: none; }
  .dist-tag { font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 4px;
    flex-shrink: 0; margin-top: 2px; font-family: 'JetBrains Mono', monospace; }
  .dist-tag.masreya { background: rgba(245,158,11,0.15); color: var(--masreya); }
  .dist-tag.purex { background: rgba(59,130,246,0.15); color: var(--purex); }
  .pharmacy-info { flex: 1; }
  .row-actions { display: flex; gap: 6px; flex-shrink: 0; align-items: flex-start; }
  .pharmacy-orig-name { font-size: 13px; font-weight: 500; color: var(--text); }
  .pharmacy-orig-code { font-size: 10px; color: var(--muted); font-family: 'JetBrains Mono', monospace; }
  .pharmacy-addr { font-size: 11px; color: var(--muted); margin-top: 2px; }
  .maps-link { font-size: 10px; color: var(--accent2); text-decoration: none; margin-top: 2px; display: inline-flex; align-items: center; gap: 4px; }
  .maps-link:hover { text-decoration: underline; }

  .ai-result { background: rgba(0,200,150,0.04); border: 1px solid rgba(0,200,150,0.15);
    border-radius: 8px; padding: 12px 14px; margin-top: 10px; }
  .ai-result-label { font-size: 10px; color: var(--accent); font-weight: 700; letter-spacing: 1px;
    text-transform: uppercase; margin-bottom: 6px; }

  .processing-row { display: flex; align-items: center; gap: 8px; padding: 12px 0;
    color: var(--muted); font-size: 13px; }
  .spinner { width: 14px; height: 14px; border: 2px solid var(--border);
    border-top-color: var(--accent); border-radius: 50%; animation: spin 0.8s linear infinite; flex-shrink: 0; }
  @keyframes spin { to { transform: rotate(360deg); } }

  .queue-list { display: flex; flex-direction: column; gap: 6px; max-height: 280px; overflow-y: auto; }
  .queue-item { display: flex; align-items: center; gap: 8px; padding: 6px 8px;
    background: var(--surface2); border-radius: 6px; font-size: 11px; }
  .queue-item.done { opacity: 0.4; }
  .queue-item.active { background: rgba(0,200,150,0.08); border: 1px solid rgba(0,200,150,0.2); }
  .queue-item.error { background: rgba(248,81,73,0.08); border: 1px solid rgba(248,81,73,0.3); }
  .queue-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
  .queue-dot.done { background: var(--success); }
  .queue-dot.active { background: var(--accent); animation: pulse 1s infinite; }
  .queue-dot.pending { background: var(--border); }
  .queue-dot.error { background: var(--danger); }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }

  .export-section { margin-top: 12px; }

  .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center;
    height: 400px; color: var(--muted); gap: 12px; }
  .empty-state .icon { font-size: 48px; opacity: 0.3; }
  .empty-state p { font-size: 14px; }

  .tabs { display: flex; gap: 2px; background: var(--surface2); border-radius: var(--radius);
    padding: 3px; margin-bottom: 16px; }
  .tab { flex: 1; padding: 7px; border-radius: 6px; border: none; background: transparent;
    color: var(--muted); font-family: inherit; font-size: 12px; cursor: pointer; transition: all 0.15s; }
  .tab.active { background: var(--surface); color: var(--text); font-weight: 600; }

  .chips { display: flex; flex-wrap: wrap; gap: 6px; }
  .chip { font-size: 11px; padding: 3px 10px; border-radius: 12px; border: 1px solid var(--border);
    color: var(--muted); cursor: pointer; transition: all 0.15s; background: transparent; font-family: inherit; }
  .chip.active { background: rgba(0,200,150,0.1); border-color: var(--accent); color: var(--accent); }

  .alert { padding: 10px 14px; border-radius: var(--radius); font-size: 12px; line-height: 1.5; }
  .alert-info { background: rgba(0,153,255,0.08); border: 1px solid rgba(0,153,255,0.2); color: #7dc8ff; }
  .alert-warn { background: rgba(245,158,11,0.08); border: 1px solid rgba(245,158,11,0.2); color: #fbbf24; }

  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
`;

export default function PharmacyMatcher() {
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [masterGroups, setMasterGroups] = useState([]);
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [processedCount, setProcessedCount] = useState(0);
  const [activeTab, setActiveTab] = useState("results");
  const [filterDist, setFilterDist] = useState("all");
  const [batchSize, setBatchSize] = useState(BATCH_SIZE);
  const [startFrom, setStartFrom] = useState(0);
  const [queueItems, setQueueItems] = useState([]);
  const [currentItem, setCurrentItem] = useState(null);
  const [apiKey, setApiKey] = useState(() => localStorage.getItem("anthropicApiKey") || "");
  const [manualSearch, setManualSearch] = useState("");
  const [manualSelected, setManualSelected] = useState([]); // array of {batchIndex, ph}
  const abortRef = useRef(false);
  const nextMasterCode = useRef(1000);
  const fileInputRef = useRef(null);

  const getCombinedData = () => {
    if (uploadedFiles.length === 0) return SAMPLE_DATA;
    return uploadedFiles.flatMap(f => f.data);
  };

  // Adds a stable unique id to each record (used by manual matching mode)
  const getCombinedDataWithIds = () => {
    return getCombinedData().map((p, i) => ({ ...p, _uid: `${p.distributor || "U"}|${p.orig_code || "ROW"}|${i}` }));
  };

  const handleApiKeyChange = (e) => {
    const value = e.target.value.trim();
    setApiKey(value);
    if (value) localStorage.setItem("anthropicApiKey", value);
    else localStorage.removeItem("anthropicApiKey");
  };

  const parseAndAddFile = (file) => {
    const extension = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();
    reader.onload = (e) => {
      let rows = [];
      if (extension === 'csv') {
        const result = Papa.parse(e.target.result, { header: true, skipEmptyLines: true });
        rows = result.data;
      } else if (['xlsx', 'xls'].includes(extension)) {
        const workbook = XLSX.read(e.target.result, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet);
      } else {
        alert(`Unsupported file type: ${file.name}`);
        return;
      }

      const pharmacies = rows.map((row, idx) => ({
        orig_code: row.orig_code || row.code || row.id || `ROW${idx+1}`,
        name: row.name || row.pharmacy_name || row["اسم الصيدلية"] || "",
        address: row.address || row["العنوان"] || "",
        province: row.province || row["المحافظة"] || "",
        district: row.district || row["المركز/الحي"] || "",
        distributor: row.distributor || row["الموزع"] || `Uploaded (${file.name})`,
      })).filter(p => p.name);

      if (pharmacies.length === 0) {
        alert(`No valid pharmacy records found in ${file.name}`);
        return;
      }

      setUploadedFiles(prev => [...prev, { name: file.name, data: pharmacies }]);
      setMasterGroups([]);
      setProcessedCount(0);
      setProgress(0);
      setQueueItems([]);
    };
    if (extension === 'csv') reader.readAsText(file);
    else reader.readAsBinaryString(file);
  };

  const handleFileSelect = (e) => {
    const files = Array.from(e.target.files);
    files.forEach(file => parseAndAddFile(file));
    e.target.value = null;
  };

  const removeFile = (index) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index));
    setMasterGroups([]);
    setProcessedCount(0);
    setProgress(0);
    setQueueItems([]);
  };

  const clearAllFiles = () => {
    setUploadedFiles([]);
    setMasterGroups([]);
    setProcessedCount(0);
    setProgress(0);
    setQueueItems([]);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const identifyBatch = async (batch) => {
    const prompt = `You are a pharmacy data analyst. I have ${batch.length} pharmacy records from uploaded files.

For each pharmacy, use its **name** and **address** to:
1. Identify the TRUE pharmacy name (canonical Arabic name)
2. Identify the city/area from the address
3. **Decide if any two records are the exact same physical pharmacy** – pay extreme attention to the address (street, district, city). Use Google Maps web search to verify addresses when uncertain.

Return ONLY valid JSON (no markdown, no explanation):
{
  "pharmacies": [
    {
      "input_index": 0,
      "canonical_name": "...",
      "area": "city/district",
      "google_search_query": "name + area + pharmacy Egypt",
      "same_as_index": null
    }
  ],
  "groups": [
    {
      "master_name": "Canonical pharmacy name",
      "area": "Area",
      "indices": [0, 2]
    }
  ]
}

If no duplicates, groups array empty.

Records:
${batch.map((p, i) => `[${i}] Distributor: ${p.distributor} | Name: ${p.name} | Address: ${p.address || p.district || p.province}`).join("\n")}`;

    // Calls Anthropic API directly from the browser using the key the user
    // entered in the sidebar (stored only in localStorage on this device).
    if (!apiKey) {
      throw new Error("Missing Anthropic API key. Enter your API key in the sidebar.");
    }

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1000,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Anthropic API error ${response.status}: ${errText}`);
    }

    const data = await response.json();
    const text = data.content?.[0]?.text || "";
    const clean = text.replace(/```json|```/g, "").trim();
    try {
      return JSON.parse(clean);
    } catch {
      return {
        pharmacies: batch.map((p, i) => ({
          input_index: i,
          canonical_name: p.name,
          area: p.district || p.province || "",
          google_search_query: `${p.name} ${p.district || p.province} Egypt`,
          same_as_index: null,
        })),
        groups: [],
      };
    }
  };

  const startProcessing = useCallback(async () => {
    if (!apiKey) {
      alert("Please enter your Anthropic API key in the sidebar before starting.");
      return;
    }
    const dataToProcess = getCombinedData();
    if (!dataToProcess.length) {
      alert("No data to process. Upload files or use the built-in sample.");
      return;
    }
    
    abortRef.current = false;
    setIsRunning(true);
    setMasterGroups([]);
    nextMasterCode.current = 1000;

    const slice = dataToProcess.slice(startFrom, startFrom + parseInt(batchSize) * 50);
    const batches = [];
    for (let i = 0; i < slice.length; i += parseInt(batchSize)) {
      batches.push(slice.slice(i, i + parseInt(batchSize)));
    }

    setQueueItems(batches.map((b, i) => ({ id: i, label: `Batch ${i+1} (${b.length} pharmacies)`, status: "pending" })));

    const allGroups = [];
    for (let bi = 0; bi < batches.length; bi++) {
      if (abortRef.current) break;
      setQueueItems(prev => prev.map(q => q.id === bi ? { ...q, status: "active" } : q));
      setCurrentItem(`Processing batch ${bi+1}/${batches.length}…`);
      const batch = batches[bi];
      try {
        const result = await identifyBatch(batch);
        const batchGroups = [];
        const usedIndices = new Set();
        
        for (const grp of (result.groups || [])) {
          const entries = grp.indices.map(idx => {
            const ph = batch[idx];
            const meta = result.pharmacies?.find(p => p.input_index === idx);
            return { ...ph, canonical_name: meta?.canonical_name || ph.name, area: meta?.area || ph.district };
          });
          grp.indices.forEach(idx => usedIndices.add(idx));
          batchGroups.push({
            master_code: `PHM-${String(nextMasterCode.current++).padStart(5,"0")}`,
            master_name: grp.master_name || entries[0]?.canonical_name,
            area: grp.area || entries[0]?.area,
            entries,
            confirmed: true,
          });
        }
        
        for (let i = 0; i < batch.length; i++) {
          if (usedIndices.has(i)) continue;
          const ph = batch[i];
          const meta = result.pharmacies?.find(p => p.input_index === i);
          batchGroups.push({
            master_code: `PHM-${String(nextMasterCode.current++).padStart(5,"0")}`,
            master_name: meta?.canonical_name || ph.name,
            area: meta?.area || ph.district || ph.province,
            entries: [{ ...ph, canonical_name: meta?.canonical_name || ph.name, area: meta?.area || "" }],
            confirmed: false,
          });
        }
        
        allGroups.push(...batchGroups);
        setMasterGroups([...allGroups]);
        setProcessedCount(c => c + batch.length);
        setProgress(Math.round(((bi+1)/batches.length)*100));
      } catch (err) {
        console.error("Batch error:", err);
        setCurrentItem(`Error on batch ${bi+1}: ${err.message}`);
        setQueueItems(prev => prev.map(q => q.id === bi ? { ...q, status: "error" } : q));
        if (bi === 0) {
          alert(`Failed to process batch ${bi+1}:\n${err.message}\n\nCheck your API key and try again.`);
          abortRef.current = true;
        }
      }
      setQueueItems(prev => prev.map(q => q.id === bi ? { ...q, status: "done" } : q));
    }
    setCurrentItem(null);
    setIsRunning(false);
  }, [batchSize, startFrom, uploadedFiles]);

  const stopProcessing = () => {
    abortRef.current = true;
    setIsRunning(false);
    setCurrentItem(null);
  };

  const exportCSV = () => {
    const rows = [["Master Code","Master Name","Area","Distributor","Original Code","Original Name","Address","Province","District"]];
    for (const grp of masterGroups) {
      for (const e of grp.entries) {
        rows.push([grp.master_code, grp.master_name, grp.area, e.distributor, e.orig_code, e.name, e.address, e.province, e.district]);
      }
    }
    const csv = rows.map(r => r.map(c => `"${String(c||"").replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "pharmacy_master_list.csv";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
  };

  const totalPharmacies = getCombinedData().length;
  const matchedCount = masterGroups.filter(g => g.entries.length > 1).length;
  const filteredGroups = masterGroups.filter(g => {
    if (filterDist === "all") return true;
    if (filterDist === "matched") return g.entries.length > 1;
    if (filterDist === "masreya") return g.entries.some(e => e.distributor === "El-Masreya");
    if (filterDist === "purex") return g.entries.some(e => e.distributor === "Purex");
    return true;
  });

  const mapsUrl = (ph) => {
    const q = encodeURIComponent(`${ph.name} ${ph.address || ph.district || ph.province} Egypt`);
    return `https://www.google.com/maps/search/${q}`;
  };

  // ── Manual matching helpers ─────────────────────────────────────────────
  const groupedUids = new Set(masterGroups.flatMap(g => g.entries.map(e => e._uid)));

  const manualRecords = getCombinedDataWithIds().filter(p => {
    if (groupedUids.has(p._uid)) return false;
    if (!manualSearch.trim()) return true;
    const term = manualSearch.trim().toLowerCase();
    return [p.name, p.address, p.district, p.province, p.orig_code]
      .some(v => String(v || "").toLowerCase().includes(term));
  });

  const toggleManualSelect = (ph) => {
    setManualSelected(prev => {
      const exists = prev.find(x => x._uid === ph._uid);
      if (exists) return prev.filter(x => x._uid !== ph._uid);
      return [...prev, ph];
    });
  };

  const clearManualSelection = () => setManualSelected([]);

  const createManualGroup = () => {
    if (manualSelected.length === 0) return;
    const masterName = manualSelected[0].name;
    const area = manualSelected[0].district || manualSelected[0].province || "";
    const entries = manualSelected.map(p => ({
      ...p,
      canonical_name: p.name,
      area: p.district || p.province || "",
    }));
    setMasterGroups(prev => [
      ...prev,
      {
        master_code: `PHM-${String(nextMasterCode.current++).padStart(5, "0")}`,
        master_name: masterName,
        area,
        entries,
        confirmed: true,
      },
    ]);
    setManualSelected([]);
  };

  const addSingleAsGroup = (ph) => {
    const entries = [{ ...ph, canonical_name: ph.name, area: ph.district || ph.province || "" }];
    setMasterGroups(prev => [
      ...prev,
      {
        master_code: `PHM-${String(nextMasterCode.current++).padStart(5, "0")}`,
        master_name: ph.name,
        area: ph.district || ph.province || "",
        entries,
        confirmed: false,
      },
    ]);
  };

  const removeMasterGroup = (masterCode) => {
    setMasterGroups(prev => prev.filter(g => g.master_code !== masterCode));
  };

  return (
    <>
      <style>{css}</style>
      <div className="app">
        <header className="header">
          <div className="header-icon">🏥</div>
          <div>
            <h1>Pharmacy Master Matcher</h1>
            <p>AI-powered entity resolution | Multiple file upload | Google Maps verification</p>
          </div>
          <div className="header-stats">
            <div className="stat-pill">Total: <span>{totalPharmacies}</span></div>
            <div className="stat-pill">Processed: <span>{processedCount}</span></div>
            <div className="stat-pill">Groups: <span>{masterGroups.length}</span></div>
            <div className="stat-pill">Matched: <span>{matchedCount}</span></div>
          </div>
        </header>

        <div className="main">
          <aside className="sidebar">
            <div className="sidebar-section">
              <h3>📁 Upload Files (Multiple)</h3>
              <input
                type="file"
                multiple
                accept=".csv,.xlsx,.xls"
                ref={fileInputRef}
                onChange={handleFileSelect}
                disabled={isRunning}
                style={{ marginBottom: "8px", width: "100%" }}
              />
              {uploadedFiles.length > 0 && (
                <div style={{ marginTop: "8px", maxHeight: "150px", overflowY: "auto" }}>
                  {uploadedFiles.map((file, idx) => (
                    <div key={idx} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: "11px", marginBottom: "4px" }}>
                      <span>📄 {file.name} ({file.data.length})</span>
                      <button onClick={() => removeFile(idx)} style={{ background: "none", border: "none", color: "var(--danger)", cursor: "pointer", fontSize: "14px" }}>✖</button>
                    </div>
                  ))}
                </div>
              )}
              {uploadedFiles.length > 0 && (
                <button className="btn btn-secondary" onClick={clearAllFiles} disabled={isRunning} style={{ marginTop: "6px" }}>
                  🗑 Clear all files
                </button>
              )}
              {uploadedFiles.length === 0 && (
                <div className="alert alert-info" style={{ fontSize: "11px", marginTop: "6px" }}>
                  ℹ️ Using sample data. Upload your own CSV/Excel files to match.
                </div>
              )}
            </div>

            <div className="sidebar-section">
              <h3>🔑 Anthropic API Key</h3>
              <div className="input-group">
                <label>Your API key (stored only in this browser)</label>
                <input
                  type="password"
                  placeholder="sk-ant-..."
                  value={apiKey}
                  onChange={handleApiKeyChange}
                  disabled={isRunning}
                />
              </div>
            </div>

            <div className="sidebar-section">
              <h3>⚙️ Configuration</h3>
              <div className="input-group">
                <label>Start from record #</label>
                <input type="number" value={startFrom} min={0} max={totalPharmacies-1} onChange={e => setStartFrom(parseInt(e.target.value)||0)} disabled={isRunning} />
              </div>
              <div className="input-group">
                <label>Batch size (per AI call)</label>
                <select value={batchSize} onChange={e => setBatchSize(parseInt(e.target.value))} disabled={isRunning}>
                  <option value={3}>3 (accurate)</option>
                  <option value={5}>5 (recommended)</option>
                  <option value={8}>8 (faster)</option>
                  <option value={10}>10 (fastest)</option>
                </select>
              </div>
            </div>

            {isRunning ? (
              <button className="btn btn-danger" onClick={stopProcessing}>⏹ Stop Processing</button>
            ) : (
              <button className="btn btn-primary" onClick={startProcessing} disabled={isRunning}>
                ▶ Start AI Matching
              </button>
            )}

            {isRunning && (
              <div className="sidebar-section">
                <h3>Progress</h3>
                <div className="progress-bar-wrap" style={{ marginBottom: 8 }}>
                  <div className="progress-bar" style={{ width: `${progress}%` }} />
                </div>
                <div style={{ fontSize: 11, color: "var(--muted)", textAlign: "center" }}>{progress}% — {currentItem}</div>
              </div>
            )}

            {queueItems.length > 0 && (
              <div className="sidebar-section">
                <h3>Queue</h3>
                <div className="queue-list">
                  {queueItems.map(q => (
                    <div key={q.id} className={`queue-item ${q.status}`}>
                      <div className={`queue-dot ${q.status}`} />
                      <span>{q.label}</span>
                      {q.status === "done" && <span style={{ marginRight: "auto", color: "var(--success)" }}>✓</span>}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {masterGroups.length > 0 && (
              <div className="sidebar-section export-section">
                <h3>Export</h3>
                <button className="btn btn-secondary" onClick={exportCSV}>⬇ Export Master List CSV</button>
              </div>
            )}

            <div className="alert alert-info" style={{ fontSize: 11 }}>
              💡 Pharmacies with the same master code are duplicates — matched by AI, or grouped manually in the "Manual Match" tab (free, no API key needed).
            </div>
          </aside>

          <main className="content">
            <div className="tabs">
              <button className={`tab ${activeTab === "results" ? "active" : ""}`} onClick={() => setActiveTab("results")}>
                All Groups ({masterGroups.length})
              </button>
              <button className={`tab ${activeTab === "matched" ? "active" : ""}`} onClick={() => setActiveTab("matched")}>
                Matched Pairs ({matchedCount})
              </button>
              <button className={`tab ${activeTab === "manual" ? "active" : ""}`} onClick={() => setActiveTab("manual")}>
                Manual Match
              </button>
            </div>

            {activeTab === "results" && (
              <>
                <div className="chips" style={{ marginBottom: 14 }}>
                  {[["all","All"],["matched","Matched Only"],["masreya","El-Masreya"],["purex","Purex"]].map(([v,l]) => (
                    <button key={v} className={`chip ${filterDist===v?"active":""}`} onClick={() => setFilterDist(v)}>{l}</button>
                  ))}
                </div>

                {filteredGroups.length === 0 && !isRunning && (
                  <div className="empty-state">
                    <div className="icon">🔍</div>
                    <p>Click "Start AI Matching" to begin processing</p>
                  </div>
                )}

                {filteredGroups.map((grp) => (
                  <div key={grp.master_code} className={`master-group ${grp.entries.length > 1 ? "confirmed" : "pending"}`}>
                    <div className="master-header">
                      <div className="master-code">{grp.master_code}</div>
                      <div>
                        <div className="master-name">{grp.master_name}</div>
                        <div className="master-meta">{grp.area}</div>
                      </div>
                      <div className="master-badges">
                        {grp.entries.some(e => e.distributor === "El-Masreya") && <span className="badge badge-masreya">El-Masreya</span>}
                        {grp.entries.some(e => e.distributor === "Purex") && <span className="badge badge-purex">Purex</span>}
                        <span className={`badge ${grp.entries.length > 1 ? "badge-status-ok" : "badge-status-pending"}`}>
                          {grp.entries.length > 1 ? `✓ ${grp.entries.length} matched` : `${grp.entries.length} entry`}
                        </span>
                        <button
                          className="btn btn-danger"
                          style={{ width: "auto", padding: "4px 10px", fontSize: 11 }}
                          onClick={() => removeMasterGroup(grp.master_code)}
                          title="Remove this group and return its records to Manual Match"
                        >
                          ✕ Ungroup
                        </button>
                      </div>
                    </div>
                    <div className="master-body">
                      {grp.entries.map((e, ei) => (
                        <div key={ei} className="pharmacy-row">
                          <div className={`dist-tag ${e.distributor === "El-Masreya" ? "masreya" : "purex"}`}>
                            {e.distributor === "El-Masreya" ? "MASREYA" : e.distributor === "Purex" ? "PUREX" : "UPLOADED"}
                          </div>
                          <div className="pharmacy-info">
                            <div className="pharmacy-orig-name">{e.name}</div>
                            <div className="pharmacy-orig-code">Code: {e.orig_code}</div>
                            {e.address && <div className="pharmacy-addr">📍 {e.address}</div>}
                            {e.district && <div className="pharmacy-addr">🗺 {e.district}{e.province ? ` · ${e.province}` : ""}</div>}
                            <a className="maps-link" href={mapsUrl(e)} target="_blank" rel="noopener noreferrer">
                              🗺 Verify on Google Maps ↗
                            </a>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </>
            )}

            {activeTab === "matched" && (
              <>
                {masterGroups.filter(g => g.entries.length > 1).length === 0 ? (
                  <div className="empty-state">
                    <div className="icon">🔗</div>
                    <p>No duplicate pharmacies found yet</p>
                    <p style={{ fontSize: 12 }}>Upload files with potential duplicates and run matching</p>
                  </div>
                ) : (
                  masterGroups.filter(g => g.entries.length > 1).map((grp) => (
                    <div key={grp.master_code} className="master-group confirmed">
                      <div className="master-header">
                        <div className="master-code">{grp.master_code}</div>
                        <div>
                          <div className="master-name">{grp.master_name}</div>
                          <div className="master-meta">{grp.area}</div>
                        </div>
                        <div className="master-badges">
                          <span className="badge badge-status-ok">✓ Duplicate Set</span>
                        </div>
                      </div>
                      <div className="master-body">
                        {grp.entries.map((e, ei) => (
                          <div key={ei} className="pharmacy-row">
                            <div className={`dist-tag ${e.distributor === "El-Masreya" ? "masreya" : "purex"}`}>
                              {e.distributor === "El-Masreya" ? "MASREYA" : e.distributor === "Purex" ? "PUREX" : "UPLOADED"}
                            </div>
                            <div className="pharmacy-info">
                              <div className="pharmacy-orig-name">{e.name}</div>
                              <div className="pharmacy-orig-code">Code: {e.orig_code}</div>
                              {e.address && <div className="pharmacy-addr">📍 {e.address}</div>}
                              <a className="maps-link" href={mapsUrl(e)} target="_blank" rel="noopener noreferrer">
                                🗺 Verify on Google Maps ↗
                              </a>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))
                )}
              </>
            )}

            {activeTab === "manual" && (
              <>
                <div className="alert alert-info" style={{ marginBottom: 14 }}>
                  💡 Free manual mode — no AI calls. Search records below, select two or more that are the same
                  pharmacy, then click "Group Selected" to create a master group. Single records can also be added
                  to the master list on their own using "Add as single".
                </div>

                <div className="input-group" style={{ marginBottom: 14 }}>
                  <label>Search by name, address, district, province, or code</label>
                  <input
                    type="text"
                    placeholder="Type to search…"
                    value={manualSearch}
                    onChange={e => setManualSearch(e.target.value)}
                  />
                </div>

                {manualSelected.length > 0 && (
                  <div className="master-group confirmed" style={{ marginBottom: 14 }}>
                    <div className="master-header">
                      <div className="master-code">SELECTED</div>
                      <div>
                        <div className="master-name">{manualSelected.length} record{manualSelected.length > 1 ? "s" : ""} selected</div>
                        <div className="master-meta">Review, then group them as duplicates</div>
                      </div>
                      <div className="master-badges">
                        <button className="btn btn-primary" style={{ width: "auto", padding: "6px 14px" }} onClick={createManualGroup}>
                          ✓ Group Selected
                        </button>
                        <button className="btn btn-secondary" style={{ width: "auto", padding: "6px 14px" }} onClick={clearManualSelection}>
                          Clear
                        </button>
                      </div>
                    </div>
                    <div className="master-body">
                      {manualSelected.map((e) => (
                        <div key={e._uid} className="pharmacy-row">
                          <div className={`dist-tag ${e.distributor === "El-Masreya" ? "masreya" : "purex"}`}>
                            {e.distributor === "El-Masreya" ? "MASREYA" : e.distributor === "Purex" ? "PUREX" : "UPLOADED"}
                          </div>
                          <div className="pharmacy-info">
                            <div className="pharmacy-orig-name">{e.name}</div>
                            <div className="pharmacy-orig-code">Code: {e.orig_code}</div>
                            {e.address && <div className="pharmacy-addr">📍 {e.address}</div>}
                            {e.district && <div className="pharmacy-addr">🗺 {e.district}{e.province ? ` · ${e.province}` : ""}</div>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {manualRecords.length === 0 ? (
                  <div className="empty-state">
                    <div className="icon">✅</div>
                    <p>{manualSearch ? "No records match your search" : "All records have been grouped"}</p>
                  </div>
                ) : (
                  <div className="master-group pending">
                    <div className="master-body">
                      {manualRecords.map((e) => {
                        const isSelected = manualSelected.some(x => x._uid === e._uid);
                        return (
                          <div key={e._uid} className="pharmacy-row" style={isSelected ? { background: "rgba(0,200,150,0.08)" } : undefined}>
                            <div className={`dist-tag ${e.distributor === "El-Masreya" ? "masreya" : "purex"}`}>
                              {e.distributor === "El-Masreya" ? "MASREYA" : e.distributor === "Purex" ? "PUREX" : "UPLOADED"}
                            </div>
                            <div className="pharmacy-info">
                              <div className="pharmacy-orig-name">{e.name}</div>
                              <div className="pharmacy-orig-code">Code: {e.orig_code}</div>
                              {e.address && <div className="pharmacy-addr">📍 {e.address}</div>}
                              {e.district && <div className="pharmacy-addr">🗺 {e.district}{e.province ? ` · ${e.province}` : ""}</div>}
                              <a className="maps-link" href={mapsUrl(e)} target="_blank" rel="noopener noreferrer">
                                🗺 Verify on Google Maps ↗
                              </a>
                            </div>
                            <div className="row-actions">
                              <button
                                className={isSelected ? "btn btn-primary" : "btn btn-secondary"}
                                style={{ width: "auto", padding: "6px 12px" }}
                                onClick={() => toggleManualSelect(e)}
                              >
                                {isSelected ? "✓ Selected" : "Select"}
                              </button>
                              <button
                                className="btn btn-secondary"
                                style={{ width: "auto", padding: "6px 12px" }}
                                onClick={() => addSingleAsGroup(e)}
                              >
                                Add as single
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </>
            )}
          </main>
        </div>
      </div>
    </>
  );
}