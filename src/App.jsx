import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import logoImg from "./assets/logo.jpg";
import {
  Ship, FileText, Settings, Search, Plus, Trash2, Pencil, Printer,
  Download, X, Check, Filter, AlertTriangle, MapPin, Truck, ChevronDown,
  Lock, ClipboardList, Receipt, BadgeCheck, LayoutDashboard, Boxes, Wallet, LogOut,
  UploadCloud, CheckCircle2, Eye, FileDown
} from "lucide-react";
import { ResponsiveContainer, LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend } from "recharts";
import { loadKey, saveKey, supabaseConfigured, supabase } from "./storage";

/* ============================= DESIGN TOKENS ============================= */
const C = {
  navy: "#0B1F3A",
  navy2: "#122B4D",
  steel: "#3B5578",
  steelSoft: "#EDF1F6",
  orange: "#D9622B",
  orangeSoft: "#FBE6DA",
  bg: "#F6F4EE",
  card: "#FFFFFF",
  ink: "#16212E",
  inkMuted: "#66707C",
  border: "#E3DFD3",
  green: "#2F7A52",
  greenSoft: "#E4F1E9",
  amber: "#B9790F",
  amberSoft: "#FBF0DA",
  red: "#B3412C",
  redSoft: "#F8E4DE",
  invoiceBlue: "#1e3a8a",
  invoiceBlueSoft: "#E3E8F7",
};

const NATURES = [
  { key: "import", label: "Import" },
  { key: "export", label: "Export" },
  { key: "transfert", label: "Transfert" },
  { key: "mise_a_terre", label: "Mise à terre" },
];
const natureLabel = (k) => NATURES.find((n) => n.key === k)?.label || k;
const sympNatures = ["import", "export"];

const FIELD_LABELS = {
  date: "Date",
  refType: "Référence",
  refValue: "N° référence",
  booking: "Booking",
  typeConteneur: "Type de conteneur",
  numeroConteneur: "N° conteneur",
  numeroCamion: "N° camion",
  lieuPriseEnCharge: "Lieu de prise en charge",
  destination: "Destination",
  client: "Client (donneur d'ordre)",
  localiteTarifaire: "Localité (tarif Sympos)",
  pleinVide: "Plein / Vide",
  tarifManuel: "Tarif (FCFA, hors Sympos)",
};

const NATURE_FIELDS = {
  import: ["date", "refType", "refValue", "typeConteneur", "numeroConteneur", "numeroCamion", "lieuPriseEnCharge", "destination", "client", "localiteTarifaire"],
  export: ["date", "booking", "typeConteneur", "numeroConteneur", "numeroCamion", "lieuPriseEnCharge", "destination", "client", "localiteTarifaire"],
  transfert: ["date", "typeConteneur", "numeroConteneur", "numeroCamion", "lieuPriseEnCharge", "pleinVide", "tarifManuel"],
  mise_a_terre: ["date", "lieuPriseEnCharge", "destination", "numeroConteneur", "typeConteneur", "numeroCamion", "client", "pleinVide", "tarifManuel"],
};

const CONTAINER_TYPES = ["20 DV", "40 DV", "40 HC", "20 RE", "40 RE", "20 OT", "40 OT", "20 FR", "40 FR"];
const normalizeContainerType = (v) => v.toUpperCase().replace(/[^0-9A-Z ]/g, "");

const blankOp = (nature) => ({
  id: null,
  nature,
  date: new Date().toISOString().slice(0, 10),
  refType: "OT",
  refValue: "",
  booking: "",
  typeConteneur: "",
  numeroConteneur: "",
  numeroCamion: "",
  lieuPriseEnCharge: "",
  destination: "",
  client: "",
  localiteTarifaire: "",
  pleinVide: "Plein",
  tarifManuel: "",
  dateFin: "",
});

const defaultSettings = {
  companyName: "Votre Entreprise",
  address: "Adresse de l'entreprise, Dakar, Sénégal",
  phone: "+221 33 000 00 00",
  email: "contact@entreprise.sn",
  ninea: "NINEA XXXXXXX",
  rccm: "RCCM SN-DKR-XXXX-XXX",
  clientName: "CEVA GROUND AND RAIL",
  clientAddress: "Dakar, Sénégal",
  footer: "Merci de votre confiance. Paiement à réception de facture. Toute réclamation doit être formulée dans les 8 jours suivant la réception de la facture.",
  invoicePrefix: "FAC",
  nextInvoiceNumber: 1,
};

const uid = () => `id_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
const fmtPlain = (n) => (isNaN(n) ? "0" : Number(n).toLocaleString("fr-FR", { maximumFractionDigits: 0 }));
const fmt = (n) => fmtPlain(n) + " FCFA";
const todayISO = () => new Date().toISOString().slice(0, 10);
function computeDurationDays(start, end) {
  if (!start || !end) return null;
  const d1 = new Date(start);
  const d2 = new Date(end);
  const diff = Math.round((d2 - d1) / 86400000);
  return diff;
}

const natureColors = { import: "#3B5578", export: "#0B1F3A", transfert: "#D9622B", mise_a_terre: "#2F7A52" };
function monthKey(dateStr) {
  return dateStr ? dateStr.slice(0, 7) : null; // "YYYY-MM"
}
function monthLabel(key) {
  const [y, m] = key.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("fr-FR", { month: "short", year: "2-digit" });
}

/* ============================= STORAGE HELPERS (Supabase-backed, see src/storage.js) ============================= */

/* ============================= PRICING ============================= */
const normLoc = (s) => (s || "").trim().replace(/\s+/g, " ").toLowerCase();
function lookupSympos(tariffs, localite, type) {
  if (!localite) return null;
  const locNorm = normLoc(localite);
  const typeNorm = normalizeContainerType(type || "").trim();
  const sameLocality = tariffs.filter((t) => normLoc(t.localite) === locNorm);
  if (sameLocality.length === 0) return null;
  // Prefer an exact type match within that locality...
  const exact = sameLocality.find((t) => normalizeContainerType(t.typeConteneur).trim() === typeNorm);
  if (exact) return Number(exact.tarif);
  // ...but as long as the locality is configured, always surface a tarif rather than leaving it blank.
  return Number(sameLocality[0].tarif);
}
function computeLine(op, tarifBase) {
  if (sympNatures.includes(op.nature)) {
    const base = Number(tarifBase) || 0;
    const remise = base * 0.2;
    const netHT = base - remise;
    const tva = netHT * 0.18;
    return { tarifSympos: base, remise, ht: netHT, tva, ttc: netHT + tva };
  }
  const montant = Number(tarifBase) || 0;
  return { tarifSympos: null, remise: 0, ht: montant, tva: 0, ttc: montant };
}

/* ============================= SMALL UI ATOMS ============================= */
function Btn({ children, onClick, kind = "primary", icon: Icon, type = "button", disabled, small }) {
  const styles = {
    primary: { background: C.orange, color: "#fff" },
    dark: { background: C.navy, color: "#fff" },
    ghost: { background: "transparent", color: C.navy, border: `1px solid ${C.border}` },
    danger: { background: C.redSoft, color: C.red },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      className={`inline-flex items-center gap-1.5 rounded-md font-medium transition-opacity ${small ? "px-2.5 py-1.5 text-xs" : "px-3.5 py-2 text-sm"} ${disabled ? "opacity-40 cursor-not-allowed" : "hover:opacity-85"}`}
      style={styles[kind]}
    >
      {Icon && <Icon size={small ? 13 : 15} />}
      {children}
    </button>
  );
}

function Field({ label, children, required }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: C.inkMuted }}>
        {label}{required && <span style={{ color: C.orange }}> *</span>}
      </span>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md px-3 py-2 text-sm outline-none focus:ring-2 transition";
const inputStyle = { border: `1px solid ${C.border}`, background: "#fff", color: C.ink };

function Badge({ children, tone = "steel" }) {
  const tones = {
    steel: { bg: C.steelSoft, color: C.steel },
    green: { bg: C.greenSoft, color: C.green },
    amber: { bg: C.amberSoft, color: C.amber },
    orange: { bg: C.orangeSoft, color: C.orange },
  };
  const t = tones[tone];
  return (
    <span className="inline-flex items-center gap-1 rounded px-2 py-0.5 text-xs font-semibold" style={{ background: t.bg, color: t.color }}>
      {children}
    </span>
  );
}

function ContainerTag({ value }) {
  return (
    <span
      className="inline-block rounded px-2 py-0.5 text-xs font-bold tracking-wider"
      style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", border: `1.5px solid ${C.navy}`, color: C.navy, background: "#fff" }}
    >
      {value || "—"}
    </span>
  );
}

/* ============================= OPERATION FORM ============================= */
function OperationForm({ initial, tariffs, onCancel, onSave }) {
  const [nature, setNature] = useState(initial?.nature || "import");
  const [form, setForm] = useState(initial || blankOp("import"));

  useEffect(() => {
    if (!initial) setForm(blankOp(nature));
  }, [nature]);

  const localites = useMemo(() => [...new Set(tariffs.map((t) => t.localite))].sort(), [tariffs]);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const fields = NATURE_FIELDS[nature];
  const requiredOk = fields.every((f) => {
    if (f === "tarifManuel") return String(form.tarifManuel || "").trim() !== "";
    return String(form[f] || "").trim() !== "";
  });

  const renderField = (key) => {
    const label = FIELD_LABELS[key];
    if (key === "date") {
      return (
        <Field label={label} required key={key}>
          <input type="date" className={inputCls} style={inputStyle} value={form.date} onChange={(e) => set("date", e.target.value)} />
        </Field>
      );
    }
    if (key === "refType") {
      return (
        <Field label={label} required key={key}>
          <div className="flex gap-2">
            <select className={inputCls} style={inputStyle} value={form.refType} onChange={(e) => set("refType", e.target.value)}>
              <option>OT</option>
              <option>N° BL</option>
            </select>
            <input className={inputCls} style={inputStyle} placeholder="Valeur" value={form.refValue} onChange={(e) => set("refValue", e.target.value)} />
          </div>
        </Field>
      );
    }
    if (key === "refValue") return null;
    if (key === "typeConteneur") {
      return (
        <Field label={label} required key={key}>
          <input list="ctypes" className={inputCls} style={inputStyle} value={form.typeConteneur} onChange={(e) => set("typeConteneur", normalizeContainerType(e.target.value))} placeholder="ex: 40 HC" />
          <datalist id="ctypes">{CONTAINER_TYPES.map((t) => <option key={t} value={t} />)}</datalist>
        </Field>
      );
    }
    if (key === "numeroConteneur") {
      return (
        <Field label={label} required key={key}>
          <input className={inputCls} style={{ ...inputStyle, fontFamily: "ui-monospace, monospace", letterSpacing: "0.05em" }} value={form.numeroConteneur} onChange={(e) => set("numeroConteneur", e.target.value.toUpperCase())} placeholder="ex: CEVU1234567" />
        </Field>
      );
    }
    if (key === "pleinVide") {
      return (
        <Field label={label} required key={key}>
          <select className={inputCls} style={inputStyle} value={form.pleinVide} onChange={(e) => set("pleinVide", e.target.value)}>
            <option>Plein</option>
            <option>Vide</option>
          </select>
        </Field>
      );
    }
    if (key === "localiteTarifaire") {
      return (
        <Field label={label} required key={key}>
          <input list="localites" className={inputCls} style={inputStyle} value={form.localiteTarifaire} onChange={(e) => set("localiteTarifaire", e.target.value)} placeholder="ex: Dakar Port" />
          <datalist id="localites">{localites.map((l) => <option key={l} value={l} />)}</datalist>
          {localites.length === 0 && <span className="text-xs" style={{ color: C.amber }}>Aucun tarif Sympos configuré — onglet "Tarifs Sympos"</span>}
        </Field>
      );
    }
    if (key === "tarifManuel") {
      return (
        <Field label={label} required key={key}>
          <input type="number" min="0" className={inputCls} style={inputStyle} value={form.tarifManuel} onChange={(e) => set("tarifManuel", e.target.value)} placeholder="0" />
        </Field>
      );
    }
    return (
      <Field label={label} required key={key}>
        <input className={inputCls} style={inputStyle} value={form[key] || ""} onChange={(e) => set(key, e.target.value)} />
      </Field>
    );
  };

  return (
    <div className="rounded-lg p-5" style={{ background: C.card, border: `1px solid ${C.border}` }}>
      <div className="flex items-center justify-between mb-4">
        <h3 className="font-bold text-base" style={{ color: C.navy }}>{initial ? "Modifier l'opération" : "Nouvelle opération"}</h3>
        <button onClick={onCancel} className="opacity-60 hover:opacity-100"><X size={18} /></button>
      </div>

      <div className="mb-4">
        <span className="text-xs font-semibold uppercase tracking-wide block mb-2" style={{ color: C.inkMuted }}>Type d'opération *</span>
        <div className="flex flex-wrap gap-2">
          {NATURES.map((n) => (
            <button
              key={n.key}
              type="button"
              disabled={!!initial}
              onClick={() => setNature(n.key)}
              className="px-3 py-1.5 rounded-md text-sm font-semibold border transition"
              style={nature === n.key
                ? { background: C.navy, color: "#fff", borderColor: C.navy }
                : { background: "#fff", color: C.navy, borderColor: C.border }}
            >
              {n.label}
            </button>
          ))}
        </div>
      </div>

      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}>
        {fields.map(renderField)}
      </div>

      <div className="flex justify-end gap-2 mt-5 pt-4" style={{ borderTop: `1px solid ${C.border}` }}>
        <Btn kind="ghost" onClick={onCancel}>Annuler</Btn>
        <Btn kind="primary" icon={Check} disabled={!requiredOk} onClick={() => onSave({ ...form, nature })}>
          Enregistrer l'opération
        </Btn>
      </div>
    </div>
  );
}

/* ============================= END DATE / DURATION CELL ============================= */
function EndDateCell({ op, onSet }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(op.dateFin || "");

  if (op.dateFin && !editing) {
    const dur = computeDurationDays(op.date, op.dateFin);
    return (
      <div className="flex items-center gap-1.5 whitespace-nowrap">
        <span className="text-xs">{op.dateFin}</span>
        <Badge tone={dur > 5 ? "amber" : "green"}>{dur} j</Badge>
        <button title="Modifier la date de fin" onClick={() => { setVal(op.dateFin); setEditing(true); }} className="p-1 opacity-60 hover:opacity-100"><Pencil size={12} /></button>
      </div>
    );
  }
  return (
    <div className="flex items-center gap-1">
      <input
        type="date" min={op.date} className="rounded px-2 py-1 text-xs"
        style={inputStyle} value={val} onChange={(e) => setVal(e.target.value)}
      />
      <button
        disabled={!val}
        title="Enregistrer la date de fin"
        onClick={() => { onSet(op.id, val); setEditing(false); }}
        className="p-1.5 rounded disabled:opacity-30"
        style={{ color: C.green }}
      >
        <Check size={14} />
      </button>
    </div>
  );
}

/* ============================= DASHBOARD TAB ============================= */
function KpiCard({ icon: Icon, label, value, tone }) {
  return (
    <div className="rounded-lg p-4 flex items-center gap-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
      <div className="p-2.5 rounded-lg" style={{ background: tone.bg, color: tone.color }}><Icon size={18} /></div>
      <div>
        <div className="text-xs font-semibold uppercase tracking-wide" style={{ color: C.inkMuted }}>{label}</div>
        <div className="text-lg font-bold" style={{ color: C.navy }}>{value}</div>
      </div>
    </div>
  );
}

function DashboardTab({ operations, invoices }) {
  const caTotal = invoices.reduce((s, iv) => s + iv.totals.ttc, 0);
  const conteneursTotal = operations.length;

  const opById = useMemo(() => Object.fromEntries(operations.map((o) => [o.id, o])), [operations]);

  const truckStats = useMemo(() => {
    const map = {};
    operations.forEach((o) => {
      const camion = (o.numeroCamion || "").trim();
      if (!camion) return;
      if (!map[camion]) map[camion] = { camion, import: 0, export: 0, transfert: 0, mise_a_terre: 0, ca: 0, missions: 0 };
      map[camion][o.nature] += 1;
      map[camion].missions += 1;
    });
    invoices.forEach((iv) => {
      iv.lines.forEach((l) => {
        const op = opById[l.opId];
        const camion = op && (op.numeroCamion || "").trim();
        if (!camion || !map[camion]) return;
        map[camion].ca += l.ttc;
      });
    });
    return Object.values(map).sort((a, b) => b.ca - a.ca);
  }, [operations, invoices, opById]);

  const caEvolution = useMemo(() => {
    const map = {};
    invoices.forEach((iv) => {
      const k = monthKey(iv.date);
      if (!k) return;
      map[k] = (map[k] || 0) + iv.totals.ttc;
    });
    return Object.keys(map).sort().map((k) => ({ month: monthLabel(k), ca: Math.round(map[k]) }));
  }, [invoices]);

  const missionsEvolution = useMemo(() => {
    const map = {};
    operations.forEach((o) => {
      const k = monthKey(o.date);
      if (!k) return;
      if (!map[k]) map[k] = { key: k, month: monthLabel(k), import: 0, export: 0, transfert: 0, mise_a_terre: 0 };
      map[k][o.nature] += 1;
    });
    return Object.values(map).sort((a, b) => (a.key > b.key ? 1 : -1));
  }, [operations]);

  return (
    <div className="space-y-4">
      <h2 className="text-lg font-bold" style={{ color: C.navy }}>Tableau de bord</h2>

      <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))" }}>
        <KpiCard icon={Wallet} label="CA total" value={fmt(caTotal)} tone={{ bg: C.orangeSoft, color: C.orange }} />
        <KpiCard icon={Boxes} label="Conteneurs total" value={conteneursTotal} tone={{ bg: C.steelSoft, color: C.steel }} />
        <KpiCard icon={Receipt} label="Factures émises" value={invoices.length} tone={{ bg: C.greenSoft, color: C.green }} />
        <KpiCard icon={Truck} label="Camions actifs" value={truckStats.length} tone={{ bg: C.amberSoft, color: C.amber }} />
      </div>

      <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))" }}>
        <div className="rounded-lg p-4" style={{ background: C.card, border: `1px solid ${C.border}` }}>
          <div className="font-semibold text-sm mb-3" style={{ color: C.navy }}>Évolution du chiffre d'affaires</div>
          {caEvolution.length === 0 ? (
            <div className="text-sm py-10 text-center" style={{ color: C.inkMuted }}>Aucune facture pour le moment.</div>
          ) : (
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <LineChart data={caEvolution}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} stroke={C.inkMuted} />
                  <YAxis tick={{ fontSize: 12 }} stroke={C.inkMuted} tickFormatter={(v) => fmtPlain(v)} />
                  <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 13, borderRadius: 8, border: `1px solid ${C.border}` }} />
                  <Line type="monotone" dataKey="ca" name="CA (TTC)" stroke={C.orange} strokeWidth={2.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>

        <div className="rounded-lg p-4" style={{ background: C.card, border: `1px solid ${C.border}` }}>
          <div className="font-semibold text-sm mb-3" style={{ color: C.navy }}>Évolution des missions par nature</div>
          {missionsEvolution.length === 0 ? (
            <div className="text-sm py-10 text-center" style={{ color: C.inkMuted }}>Aucune opération pour le moment.</div>
          ) : (
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={missionsEvolution}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                  <XAxis dataKey="month" tick={{ fontSize: 12 }} stroke={C.inkMuted} />
                  <YAxis tick={{ fontSize: 12 }} stroke={C.inkMuted} allowDecimals={false} />
                  <Tooltip contentStyle={{ fontSize: 13, borderRadius: 8, border: `1px solid ${C.border}` }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} formatter={(v) => natureLabel(v)} />
                  <Bar dataKey="import" name="import" stackId="a" fill={natureColors.import} />
                  <Bar dataKey="export" name="export" stackId="a" fill={natureColors.export} />
                  <Bar dataKey="transfert" name="transfert" stackId="a" fill={natureColors.transfert} />
                  <Bar dataKey="mise_a_terre" name="mise_a_terre" stackId="a" fill={natureColors.mise_a_terre} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>

      <div>
        <div className="font-semibold text-sm mb-2" style={{ color: C.navy }}>Performance par camion</div>
        <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.border}` }}>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr style={{ background: C.steelSoft }}>
                  {["Camion", "CA", "Import", "Export", "Transfert", "Mise à terre", "Total missions"].map((h) => (
                    <th key={h} className="text-left px-3 py-2 font-semibold text-xs uppercase tracking-wide" style={{ color: C.steel }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {truckStats.length === 0 && (
                  <tr><td colSpan={7} className="text-center py-8 text-sm" style={{ color: C.inkMuted }}>Aucune donnée camion pour le moment.</td></tr>
                )}
                {truckStats.map((t) => (
                  <tr key={t.camion} style={{ borderTop: `1px solid ${C.border}` }}>
                    <td className="px-3 py-2 font-bold" style={{ fontFamily: "ui-monospace, monospace", color: C.navy }}>{t.camion}</td>
                    <td className="px-3 py-2 font-semibold">{fmt(t.ca)}</td>
                    <td className="px-3 py-2">{t.import}</td>
                    <td className="px-3 py-2">{t.export}</td>
                    <td className="px-3 py-2">{t.transfert}</td>
                    <td className="px-3 py-2">{t.mise_a_terre}</td>
                    <td className="px-3 py-2"><Badge tone="steel">{t.missions}</Badge></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================= OPERATIONS TAB ============================= */
function OperationsTab({ operations, tariffs, onAdd, onUpdate, onDelete, onSetEndDate }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [q, setQ] = useState("");
  const [natureFilter, setNatureFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");

  const filtered = operations
    .filter((o) => (q ? o.numeroConteneur.toLowerCase().includes(q.toLowerCase()) : true))
    .filter((o) => (natureFilter === "all" ? true : o.nature === natureFilter))
    .filter((o) => (statusFilter === "all" ? true : statusFilter === "facturee" ? o.facturee : !o.facturee))
    .sort((a, b) => (b.date > a.date ? 1 : -1));

  return (
    <div className="space-y-4">
      {!showForm && !editing && (
        <div className="flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-lg font-bold" style={{ color: C.navy }}>Opérations</h2>
          <Btn icon={Plus} onClick={() => setShowForm(true)}>Nouvelle opération</Btn>
        </div>
      )}

      {showForm && (
        <OperationForm
          tariffs={tariffs}
          onCancel={() => setShowForm(false)}
          onSave={(op) => { onAdd(op); setShowForm(false); }}
        />
      )}
      {editing && (
        <OperationForm
          initial={editing}
          tariffs={tariffs}
          onCancel={() => setEditing(null)}
          onSave={(op) => { onUpdate(op); setEditing(null); }}
        />
      )}

      {!showForm && !editing && (
        <>
          <div className="flex flex-wrap gap-2 items-center rounded-lg p-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
            <div className="flex items-center gap-2 rounded-md px-2.5 py-1.5 flex-1 min-w-[180px]" style={{ border: `1px solid ${C.border}` }}>
              <Search size={14} style={{ color: C.inkMuted }} />
              <input className="outline-none text-sm flex-1" placeholder="Rechercher un n° conteneur..." value={q} onChange={(e) => setQ(e.target.value)} />
            </div>
            <select className={inputCls} style={{ ...inputStyle, width: "auto" }} value={natureFilter} onChange={(e) => setNatureFilter(e.target.value)}>
              <option value="all">Toutes natures</option>
              {NATURES.map((n) => <option key={n.key} value={n.key}>{n.label}</option>)}
            </select>
            <select className={inputCls} style={{ ...inputStyle, width: "auto" }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              <option value="all">Tous statuts</option>
              <option value="facturee">Facturées</option>
              <option value="non_facturee">Non facturées</option>
            </select>
          </div>

          <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.border}` }}>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr style={{ background: C.steelSoft }}>
                    {["Date", "Nature", "Conteneur", "Type", "N° camion", "Client", "Lieu / Destination", "Date fin / Durée", "Statut", ""].map((h) => (
                      <th key={h} className="text-left px-3 py-2 font-semibold text-xs uppercase tracking-wide" style={{ color: C.steel }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 && (
                    <tr><td colSpan={10} className="text-center py-8 text-sm" style={{ color: C.inkMuted }}>Aucune opération. Ajoutez-en une pour commencer.</td></tr>
                  )}
                  {filtered.map((o) => (
                    <tr key={o.id} style={{ borderTop: `1px solid ${C.border}` }}>
                      <td className="px-3 py-2 whitespace-nowrap">{o.date}</td>
                      <td className="px-3 py-2"><Badge tone="steel">{natureLabel(o.nature)}</Badge></td>
                      <td className="px-3 py-2"><ContainerTag value={o.numeroConteneur} /></td>
                      <td className="px-3 py-2 whitespace-nowrap">{o.typeConteneur}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{o.numeroCamion || "—"}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{o.client || "—"}</td>
                      <td className="px-3 py-2">{o.lieuPriseEnCharge}{o.destination ? ` → ${o.destination}` : ""}</td>
                      <td className="px-3 py-2"><EndDateCell op={o} onSet={onSetEndDate} /></td>
                      <td className="px-3 py-2">
                        {o.facturee
                          ? <Badge tone="green"><BadgeCheck size={12} /> {o.factureNumero}</Badge>
                          : <Badge tone="amber">Non facturée</Badge>}
                      </td>
                      <td className="px-3 py-2">
                        <div className="flex gap-1 justify-end">
                          {!o.facturee ? (
                            <>
                              <button title="Modifier" onClick={() => setEditing(o)} className="p-1.5 rounded hover:opacity-70" style={{ color: C.steel }}><Pencil size={14} /></button>
                              <button title="Supprimer" onClick={() => onDelete(o.id)} className="p-1.5 rounded hover:opacity-70" style={{ color: C.red }}><Trash2 size={14} /></button>
                            </>
                          ) : (
                            <span title="Verrouillée (déjà facturée)" style={{ color: C.inkMuted }}><Lock size={14} /></span>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================= TARIFS SYMPOS TAB ============================= */
function TariffsTab({ tariffs, onAdd, onDelete }) {
  const [form, setForm] = useState({ localite: "", typeConteneur: "", tarif: "" });
  const valid = form.localite && form.typeConteneur && form.tarif;
  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-bold" style={{ color: C.navy }}>Tarifs Sympos</h2>
        <p className="text-sm mt-1" style={{ color: C.inkMuted }}>Les tarifs varient selon la localité. Configurez-les ici pour qu'ils s'appliquent automatiquement aux opérations d'import / export.</p>
      </div>

      <div className="rounded-lg p-4 grid gap-3" style={{ background: C.card, border: `1px solid ${C.border}`, gridTemplateColumns: "repeat(auto-fill, minmax(160px,1fr))" }}>
        <Field label="Localité"><input className={inputCls} style={inputStyle} value={form.localite} onChange={(e) => setForm({ ...form, localite: e.target.value })} placeholder="ex: Dakar Port" /></Field>
        <Field label="Type de conteneur">
          <input list="ctypes2" className={inputCls} style={inputStyle} value={form.typeConteneur} onChange={(e) => setForm({ ...form, typeConteneur: normalizeContainerType(e.target.value) })} placeholder="ex: 40 HC" />
          <datalist id="ctypes2">{CONTAINER_TYPES.map((t) => <option key={t} value={t} />)}</datalist>
        </Field>
        <Field label="TARIF (FCFA)"><input type="number" min="0" className={inputCls} style={inputStyle} value={form.tarif} onChange={(e) => setForm({ ...form, tarif: e.target.value })} /></Field>
        <div className="flex items-end">
          <Btn icon={Plus} disabled={!valid} onClick={() => { onAdd({ id: uid(), ...form, tarif: Number(form.tarif) }); setForm({ localite: "", typeConteneur: "", tarif: "" }); }}>
            Ajouter
          </Btn>
        </div>
      </div>

      <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <table className="w-full text-sm">
          <thead>
            <tr style={{ background: C.steelSoft }}>
              {["Localité", "Type", "TARIF", "Après remise -20%", "TVA 18%", "Net à facturer", ""].map((h) => (
                <th key={h} className="text-left px-3 py-2 font-semibold text-xs uppercase tracking-wide" style={{ color: C.steel }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tariffs.length === 0 && <tr><td colSpan={7} className="text-center py-8 text-sm" style={{ color: C.inkMuted }}>Aucun tarif configuré.</td></tr>}
            {tariffs.map((t) => {
              const c = computeLine({ nature: "import" }, t.tarif);
              return (
                <tr key={t.id} style={{ borderTop: `1px solid ${C.border}` }}>
                  <td className="px-3 py-2">{t.localite}</td>
                  <td className="px-3 py-2">{t.typeConteneur}</td>
                  <td className="px-3 py-2">{fmt(t.tarif)}</td>
                  <td className="px-3 py-2">{fmt(c.ht)}</td>
                  <td className="px-3 py-2">{fmt(c.tva)}</td>
                  <td className="px-3 py-2 font-semibold" style={{ color: C.navy }}>{fmt(c.ttc)}</td>
                  <td className="px-3 py-2"><button onClick={() => onDelete(t.id)} className="p-1.5" style={{ color: C.red }}><Trash2 size={14} /></button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================= NEW INVOICE TAB ============================= */
function NewInvoiceTab({ operations, tariffs, settings, onCreate }) {
  const unbilled = operations.filter((o) => !o.facturee);
  const [nature, setNature] = useState(NATURES[0].key);
  const [q, setQ] = useState("");
  const [selected, setSelected] = useState({}); // id -> true
  const [manualTarif, setManualTarif] = useState({}); // id -> override value

  const changeNature = (n) => { setNature(n); setSelected({}); setManualTarif({}); };

  const list = unbilled
    .filter((o) => o.nature === nature)
    .filter((o) => (q ? o.numeroConteneur.toLowerCase().includes(q.toLowerCase()) : true))
    .sort((a, b) => (b.date > a.date ? 1 : -1));

  const toggle = (op) => {
    if (!op.dateFin) return; // must be marked "terminée" (date de fin renseignée) before it can be invoiced
    setSelected((s) => {
      const n = { ...s };
      if (n[op.id]) delete n[op.id]; else n[op.id] = true;
      return n;
    });
  };

  const selectedOps = unbilled.filter((o) => selected[o.id]);

  const lineFor = (op) => {
    let base;
    if (sympNatures.includes(op.nature)) {
      const found = lookupSympos(tariffs, op.localiteTarifaire, op.typeConteneur);
      base = manualTarif[op.id] !== undefined ? manualTarif[op.id] : found;
    } else {
      base = manualTarif[op.id] !== undefined ? manualTarif[op.id] : op.tarifManuel;
    }
    const c = computeLine(op, base || 0);
    return { ...c, base, missing: base === null || base === undefined || base === "" };
  };

  const totals = selectedOps.reduce((acc, o) => {
    const l = lineFor(o);
    acc.ht += l.ht; acc.tva += l.tva; acc.ttc += l.ttc;
    return acc;
  }, { ht: 0, tva: 0, ttc: 0 });

  const anyMissing = selectedOps.some((o) => lineFor(o).missing);

  return (
    <div className="space-y-4">
      <h2 className="text-lg font-bold" style={{ color: C.navy }}>Nouvelle facture — {settings.clientName}</h2>

      <div>
        <span className="text-xs font-semibold uppercase tracking-wide block mb-2" style={{ color: C.inkMuted }}>
          Nature de la facture — les opérations d'import, export, transfert et mise à terre ne peuvent pas être mélangées dans une même facture *
        </span>
        <div className="flex flex-wrap gap-2">
          {NATURES.map((n) => (
            <button
              key={n.key}
              type="button"
              onClick={() => changeNature(n.key)}
              className="px-3 py-1.5 rounded-md text-sm font-semibold border transition"
              style={nature === n.key
                ? { background: C.navy, color: "#fff", borderColor: C.navy }
                : { background: "#fff", color: C.navy, borderColor: C.border }}
            >
              {n.label}
            </button>
          ))}
        </div>
      </div>

      <div className="flex items-center gap-2 rounded-md px-2.5 py-1.5" style={{ border: `1px solid ${C.border}`, background: C.card, maxWidth: 360 }}>
        <Search size={14} style={{ color: C.inkMuted }} />
        <input className="outline-none text-sm flex-1" placeholder="Rechercher par n° conteneur..." value={q} onChange={(e) => setQ(e.target.value)} />
      </div>
      <p className="text-xs" style={{ color: C.inkMuted }}>
        Seules les opérations marquées <b>Terminée</b> (date de fin renseignée dans l'onglet Opérations) peuvent être sélectionnées pour la facturation.
      </p>

      <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: C.steelSoft }}>
                {["", "Date", "Nature", "Conteneur", "Type", "Statut", "Référence", "Tarif appliqué", "Montant TTC"].map((h) => (
                  <th key={h} className="text-left px-3 py-2 font-semibold text-xs uppercase tracking-wide" style={{ color: C.steel }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.length === 0 && <tr><td colSpan={9} className="text-center py-8 text-sm" style={{ color: C.inkMuted }}>Aucune opération non facturée pour la nature "{natureLabel(nature)}".</td></tr>}
              {list.map((o) => {
                const l = lineFor(o);
                const terminee = !!o.dateFin;
                const ref = o.nature === "import" ? `${o.refType}: ${o.refValue}` : o.nature === "export" ? `Booking: ${o.booking}` : "—";
                return (
                  <tr key={o.id} style={{ borderTop: `1px solid ${C.border}`, background: selected[o.id] ? C.orangeSoft : "transparent", opacity: terminee ? 1 : 0.55 }}>
                    <td className="px-3 py-2">
                      <input
                        type="checkbox" checked={!!selected[o.id]} disabled={!terminee} onChange={() => toggle(o)}
                        title={terminee ? "" : "Marquez d'abord l'opération comme terminée (date de fin) dans l'onglet Opérations"}
                        style={{ cursor: terminee ? "pointer" : "not-allowed" }}
                      />
                    </td>
                    <td className="px-3 py-2 whitespace-nowrap">{o.date}</td>
                    <td className="px-3 py-2"><Badge tone="steel">{natureLabel(o.nature)}</Badge></td>
                    <td className="px-3 py-2"><ContainerTag value={o.numeroConteneur} /></td>
                    <td className="px-3 py-2 whitespace-nowrap">{o.typeConteneur}</td>
                    <td className="px-3 py-2">
                      {terminee ? <Badge tone="green">Terminée</Badge> : <Badge tone="amber">En cours</Badge>}
                    </td>
                    <td className="px-3 py-2 whitespace-nowrap">{ref}</td>
                    <td className="px-3 py-2">
                      {selected[o.id] ? (
                        <input
                          type="number" min="0" className="w-28 rounded px-2 py-1 text-sm" style={inputStyle}
                          value={manualTarif[o.id] !== undefined ? manualTarif[o.id] : (l.base ?? "")}
                          onChange={(e) => setManualTarif((m) => ({ ...m, [o.id]: e.target.value }))}
                          placeholder={l.missing ? "manquant" : ""}
                        />
                      ) : "—"}
                    </td>
                    <td className="px-3 py-2 font-semibold" style={{ color: l.missing && selected[o.id] ? C.red : C.navy }}>
                      {selected[o.id] ? fmt(l.ttc) : "—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {selectedOps.length > 0 && (
        <div className="rounded-lg p-4 flex flex-col gap-3" style={{ background: C.navy, color: "#fff" }}>
          {anyMissing && (
            <div className="flex items-center gap-2 text-sm rounded px-3 py-2" style={{ background: "rgba(255,255,255,0.12)" }}>
              <AlertTriangle size={16} style={{ color: "#F5C16C" }} />
              Certaines lignes n'ont pas de tarif Sympos correspondant. Renseignez un tarif manuellement ou complétez l'onglet "Tarifs Sympos".
            </div>
          )}
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="text-sm opacity-80">{selectedOps.length} opération(s) sélectionnée(s)</div>
            <div className="flex gap-6 text-sm">
              <div>Total HT <b className="text-base">{fmt(totals.ht)}</b></div>
              <div>TVA <b className="text-base">{fmt(totals.tva)}</b></div>
              <div>Total TTC <b className="text-base" style={{ color: C.orangeSoft }}>{fmt(totals.ttc)}</b></div>
            </div>
            <Btn
              icon={Receipt}
              disabled={anyMissing}
              onClick={() => {
                const lines = selectedOps.map((o) => {
                  const l = lineFor(o);
                  const ref = o.nature === "import" ? `${o.refType}: ${o.refValue}` : o.nature === "export" ? `Booking: ${o.booking}` : "—";
                  return {
                    opId: o.id, numeroConteneur: o.numeroConteneur, typeConteneur: o.typeConteneur,
                    destination: o.destination || "—", nature: o.nature, reference: ref,
                    tarifSympos: l.tarifSympos, remise: l.remise, ht: l.ht, tva: l.tva, ttc: l.ttc,
                  };
                });
                onCreate({ lines, totals, opIds: selectedOps.map((o) => o.id) });
                setSelected({}); setManualTarif({});
              }}
            >
              Générer la facture
            </Btn>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================= INVOICE DOCUMENT ============================= */
function InvoiceDocument({ invoice, settings }) {
  return (
    <div id="invoice-print-area" className="rounded-xl overflow-hidden" style={{ background: "#fff", border: `1px solid ${C.border}`, color: C.ink, boxShadow: "0 1px 3px rgba(30,58,138,0.10)" }}>
      <div style={{ height: 6, background: `linear-gradient(90deg, ${C.orange}, ${C.invoiceBlue})` }} />
      <div className="p-6 sm:p-9">
        <div className="flex justify-between items-start gap-4 pb-6" style={{ borderBottom: `1px solid ${C.border}` }}>
          <div className="flex items-start gap-3 min-w-0">
            <img src={logoImg} alt={settings.companyName} className="shrink-0 rounded" style={{ height: 64, width: 64, objectFit: "contain" }} />
            <div className="min-w-0">
              <div className="font-bold text-lg tracking-tight" style={{ color: C.invoiceBlue }}>{settings.companyName}</div>
              <div className="text-xs leading-relaxed mt-0.5" style={{ color: C.inkMuted }}>
                {settings.address}<br />
                Tél: {settings.phone} · {settings.email}<br />
                {settings.ninea} · {settings.rccm}
              </div>
            </div>
          </div>
          <div className="text-right shrink-0">
            <div className="text-[10px] font-semibold uppercase tracking-widest mb-1" style={{ color: C.inkMuted }}>Facture</div>
            <div
              className="font-bold text-sm rounded-md"
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", textAlign: "center",
                minWidth: 140, padding: "6px 12px",
                background: C.invoiceBlueSoft, color: C.invoiceBlue,
                fontFamily: "ui-monospace, monospace", letterSpacing: "0.04em",
              }}
            >
              {invoice.numero}
            </div>
            <div className="text-xs mt-1.5" style={{ color: C.inkMuted }}>Émise le {invoice.date}</div>
          </div>
        </div>

        <div className="flex justify-between items-start py-5">
          <div>
            <div className="text-[10px] font-semibold uppercase tracking-widest mb-1" style={{ color: C.inkMuted }}>Facturé à</div>
            <div className="font-bold" style={{ color: C.invoiceBlue }}>{settings.clientName}</div>
            <div className="text-xs" style={{ color: C.inkMuted }}>{settings.clientAddress}</div>
          </div>
        </div>

        <table className="w-full text-sm" style={{ borderCollapse: "separate", borderSpacing: 0 }}>
          <thead>
            <tr>
              {["N° Conteneur", "Type", "Destination", "Nature", "Référence", "HT", "TVA 18%", "TTC"].map((h, i) => (
                <th
                  key={h}
                  className="text-left px-3 py-2.5 text-[10px] font-bold uppercase tracking-wider"
                  style={{ color: "#fff", background: C.invoiceBlue, borderTopLeftRadius: i === 0 ? 8 : 0, borderTopRightRadius: i === 7 ? 8 : 0 }}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {invoice.lines.map((l, i) => (
              <tr key={i} style={{ background: i % 2 === 1 ? C.steelSoft : "transparent" }}>
                <td className="px-3 py-2.5 font-bold" style={{ fontFamily: "ui-monospace, monospace", letterSpacing: "0.03em" }}>{l.numeroConteneur || "—"}</td>
                <td className="px-3 py-2.5 whitespace-nowrap">{l.typeConteneur}</td>
                <td className="px-3 py-2.5">{l.destination}</td>
                <td className="px-3 py-2.5"><Badge tone="steel">{natureLabel(l.nature)}</Badge></td>
                <td className="px-3 py-2.5 whitespace-nowrap">{l.reference}</td>
                <td className="px-3 py-2.5">{fmtPlain(l.ht)}</td>
                <td className="px-3 py-2.5">{sympNatures.includes(l.nature) ? fmtPlain(l.tva) : "Exonéré"}</td>
                <td className="px-3 py-2.5 font-semibold" style={{ color: C.invoiceBlue }}>{fmtPlain(l.ttc)}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="flex justify-end mt-5">
          <div className="w-72 rounded-lg p-4 text-sm space-y-1.5" style={{ background: C.steelSoft }}>
            <div className="flex justify-between"><span style={{ color: C.inkMuted }}>Total HT</span><span>{fmtPlain(invoice.totals.ht)}</span></div>
            <div className="flex justify-between"><span style={{ color: C.inkMuted }}>Total TVA</span><span>{fmtPlain(invoice.totals.tva)}</span></div>
            <div className="flex justify-between items-center font-bold text-base pt-2.5 mt-1" style={{ borderTop: `1px solid ${C.border}`, color: C.invoiceBlue }}>
              <span>Total TTC</span>
              <span
                className="rounded-md"
                style={{ display: "flex", alignItems: "center", justifyContent: "center", textAlign: "center", minWidth: 130, padding: "6px 10px", background: C.invoiceBlue, color: "#fff" }}
              >
                {fmt(invoice.totals.ttc)}
              </span>
            </div>
          </div>
        </div>

        <div className="mt-8 pt-4 text-xs" style={{ borderTop: `1px solid ${C.border}`, color: C.inkMuted }}>
          {settings.footer}
        </div>
      </div>
    </div>
  );
}

/* ============================= PDF EXPORT & GOOGLE DRIVE ============================= */
function triggerFileDownload(url, filename) {
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

// Renders the given DOM node (the invoice) to a real, multi-page-aware PDF blob.
async function waitForImages(node) {
  const imgs = Array.from(node.querySelectorAll("img"));
  await Promise.all(imgs.map((img) => (img.complete ? Promise.resolve() : new Promise((res) => { img.onload = res; img.onerror = res; }))));
}

async function invoiceNodeToPdfBlob(node) {
  await waitForImages(node);
  const canvas = await html2canvas(node, { scale: 2, backgroundColor: "#ffffff", useCORS: true });
  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF({ unit: "pt", format: "a4" });
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imgWidth = pageWidth;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  let heightLeft = imgHeight;
  let position = 0;
  pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
  heightLeft -= pageHeight;
  while (heightLeft > 0) {
    position -= pageHeight;
    pdf.addPage();
    pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
    heightLeft -= pageHeight;
  }
  return pdf.output("blob");
}

// Google Drive upload via Google Identity Services (loaded through the <script> tag
// in index.html). Requires VITE_GOOGLE_CLIENT_ID — see README for setup.
function getGoogleAccessToken() {
  return new Promise((resolve, reject) => {
    const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID;
    if (!clientId) return reject(new Error("Google Drive non configuré (VITE_GOOGLE_CLIENT_ID manquant)"));
    if (!window.google || !window.google.accounts || !window.google.accounts.oauth2) {
      return reject(new Error("Service Google non chargé, réessayez dans quelques secondes"));
    }
    const tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: "https://www.googleapis.com/auth/drive.file",
      callback: (resp) => (resp.error ? reject(resp) : resolve(resp.access_token)),
    });
    tokenClient.requestAccessToken();
  });
}

async function uploadBlobToDrive(blob, filename) {
  const token = await getGoogleAccessToken();
  const metadata = { name: filename, mimeType: "application/pdf" };
  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", blob);
  const res = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: form,
  });
  if (!res.ok) throw new Error("Échec de l'envoi vers Google Drive");
  return res.json();
}

/* ============================= INVOICE PREVIEW MODAL ============================= */
function InvoiceModal({ invoice, settings, onExportExcel, onClose }) {
  const docRef = useRef(null);
  const [busy, setBusy] = useState(null); // "pdf" | "drive" | null
  const [driveStatus, setDriveStatus] = useState(null); // "ok" | "error" | null
  const driveConfigured = !!import.meta.env.VITE_GOOGLE_CLIENT_ID;

  const handleDownloadPdf = async () => {
    setBusy("pdf");
    try {
      const blob = await invoiceNodeToPdfBlob(docRef.current);
      const url = URL.createObjectURL(blob);
      triggerFileDownload(url, `${invoice.numero}.pdf`);
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    } catch (e) {
      alert("Erreur lors de la génération du PDF : " + e.message);
    }
    setBusy(null);
  };

  const handleSaveToDrive = async () => {
    setBusy("drive");
    setDriveStatus(null);
    try {
      const blob = await invoiceNodeToPdfBlob(docRef.current);
      await uploadBlobToDrive(blob, `${invoice.numero}.pdf`);
      setDriveStatus("ok");
    } catch (e) {
      setDriveStatus("error");
      console.error(e);
    }
    setBusy(null);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-6" style={{ background: "rgba(11,31,58,0.6)" }} onClick={onClose}>
      <div
        className="w-full max-w-3xl rounded-xl overflow-hidden flex flex-col"
        style={{ background: C.bg, maxHeight: "94vh" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 sm:px-5 py-3 shrink-0" style={{ background: "#fff", borderBottom: `1px solid ${C.border}` }}>
          <div className="font-bold text-sm" style={{ color: C.invoiceBlue, fontFamily: "ui-monospace, monospace" }}>{invoice.numero}</div>
          <button onClick={onClose} className="p-1.5 rounded hover:opacity-70" style={{ color: C.navy }}><X size={18} /></button>
        </div>

        <div className="overflow-y-auto p-3 sm:p-6 flex-1">
          <div ref={docRef}>
            <InvoiceDocument invoice={invoice} settings={settings} />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 justify-end px-4 sm:px-5 py-3 shrink-0" style={{ background: "#fff", borderTop: `1px solid ${C.border}` }}>
          {driveStatus === "ok" && <span className="text-xs flex items-center gap-1" style={{ color: C.green }}><CheckCircle2 size={14} /> Envoyée sur Drive</span>}
          {driveStatus === "error" && <span className="text-xs" style={{ color: C.red }}>Échec de l'envoi vers Drive</span>}
          <Btn kind="ghost" icon={Download} onClick={() => onExportExcel(invoice)}>Excel</Btn>
          <button
            title={driveConfigured ? "Enregistrer sur Google Drive" : "Google Drive non configuré — voir README"}
            disabled={busy === "drive" || !driveConfigured}
            onClick={handleSaveToDrive}
            className={`inline-flex items-center justify-center rounded-md p-2.5 transition ${busy === "drive" || !driveConfigured ? "opacity-40 cursor-not-allowed" : "hover:opacity-85"}`}
            style={{ background: "transparent", color: C.invoiceBlue, border: `1px solid ${C.border}` }}
          >
            <UploadCloud size={18} />
          </button>
          <button
            title="Télécharger en PDF"
            disabled={busy === "pdf"}
            onClick={handleDownloadPdf}
            className={`inline-flex items-center justify-center rounded-md p-2.5 transition ${busy === "pdf" ? "opacity-40 cursor-not-allowed" : "hover:opacity-85"}`}
            style={{ background: C.invoiceBlue, color: "#fff" }}
          >
            <FileDown size={18} />
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================= INVOICES TAB ============================= */
function InvoicesTab({ invoices, settings }) {
  const [q, setQ] = useState("");
  const [active, setActive] = useState(null);

  const filtered = invoices.filter((inv) => {
    if (!q) return true;
    const qq = q.toLowerCase();
    return inv.numero.toLowerCase().includes(qq) || inv.lines.some((l) => l.numeroConteneur.toLowerCase().includes(qq));
  }).sort((a, b) => (b.numero > a.numero ? 1 : -1));

  const exportExcel = (inv) => {
    const rows = inv.lines.map((l) => ({
      "N° Conteneur": l.numeroConteneur, "Type": l.typeConteneur, "Destination": l.destination,
      "Nature": natureLabel(l.nature), "Référence": l.reference,
      "Montant HT": l.ht, "TVA": l.tva, "Montant TTC": l.ttc,
    }));
    rows.push({});
    rows.push({ "N° Conteneur": "TOTAL", "Montant HT": inv.totals.ht, "TVA": inv.totals.tva, "Montant TTC": inv.totals.ttc });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Facture");
    XLSX.writeFile(wb, `${inv.numero}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <h2 className="text-lg font-bold" style={{ color: C.navy }}>Factures</h2>
      <div className="flex items-center gap-2 rounded-md px-2.5 py-1.5" style={{ border: `1px solid ${C.border}`, background: C.card, maxWidth: 360 }}>
        <Search size={14} style={{ color: C.inkMuted }} />
        <input className="outline-none text-sm flex-1" placeholder="N° facture ou n° conteneur..." value={q} onChange={(e) => setQ(e.target.value)} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <table className="w-full text-sm">
          <thead>
            <tr style={{ background: C.steelSoft }}>
              {["N° Facture", "Date", "Lignes", "Total TTC", ""].map((h) => (
                <th key={h} className="text-left px-3 py-2 font-semibold text-xs uppercase tracking-wide" style={{ color: C.steel }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={5} className="text-center py-8 text-sm" style={{ color: C.inkMuted }}>Aucune facture.</td></tr>}
            {filtered.map((inv) => (
              <tr key={inv.id} style={{ borderTop: `1px solid ${C.border}` }}>
                <td className="px-3 py-2 font-semibold" style={{ fontFamily: "ui-monospace, monospace", color: C.navy }}>{inv.numero}</td>
                <td className="px-3 py-2">{inv.date}</td>
                <td className="px-3 py-2">{inv.lines.length}</td>
                <td className="px-3 py-2 font-semibold">{fmt(inv.totals.ttc)}</td>
                <td className="px-3 py-2">
                  <button
                    title="Aperçu"
                    onClick={() => setActive(inv)}
                    className="p-1.5 rounded hover:opacity-70"
                    style={{ color: C.invoiceBlue }}
                  >
                    <Eye size={16} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {active && <InvoiceModal invoice={active} settings={settings} onExportExcel={exportExcel} onClose={() => setActive(null)} />}
    </div>
  );
}

/* ============================= SETTINGS TAB ============================= */
function SettingsTab({ settings, onSave }) {
  const [form, setForm] = useState(settings);
  useEffect(() => setForm(settings), [settings]);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  return (
    <div className="space-y-5 max-w-2xl">
      <h2 className="text-lg font-bold" style={{ color: C.navy }}>Paramètres</h2>

      <div className="rounded-lg p-4 space-y-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="font-semibold text-sm" style={{ color: C.navy }}>En-tête de facture</div>
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(220px,1fr))" }}>
          <Field label="Nom de l'entreprise"><input className={inputCls} style={inputStyle} value={form.companyName} onChange={(e) => set("companyName", e.target.value)} /></Field>
          <Field label="Adresse"><input className={inputCls} style={inputStyle} value={form.address} onChange={(e) => set("address", e.target.value)} /></Field>
          <Field label="Téléphone"><input className={inputCls} style={inputStyle} value={form.phone} onChange={(e) => set("phone", e.target.value)} /></Field>
          <Field label="Email"><input className={inputCls} style={inputStyle} value={form.email} onChange={(e) => set("email", e.target.value)} /></Field>
          <Field label="NINEA"><input className={inputCls} style={inputStyle} value={form.ninea} onChange={(e) => set("ninea", e.target.value)} /></Field>
          <Field label="RCCM"><input className={inputCls} style={inputStyle} value={form.rccm} onChange={(e) => set("rccm", e.target.value)} /></Field>
        </div>
      </div>

      <div className="rounded-lg p-4 space-y-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="font-semibold text-sm" style={{ color: C.navy }}>Client facturé (unique)</div>
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(220px,1fr))" }}>
          <Field label="Nom du client"><input className={inputCls} style={inputStyle} value={form.clientName} onChange={(e) => set("clientName", e.target.value)} /></Field>
          <Field label="Adresse du client"><input className={inputCls} style={inputStyle} value={form.clientAddress} onChange={(e) => set("clientAddress", e.target.value)} /></Field>
        </div>
      </div>

      <div className="rounded-lg p-4 space-y-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="font-semibold text-sm" style={{ color: C.navy }}>Numérotation</div>
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(220px,1fr))" }}>
          <Field label="Préfixe"><input className={inputCls} style={inputStyle} value={form.invoicePrefix} onChange={(e) => set("invoicePrefix", e.target.value)} /></Field>
          <Field label="Prochain numéro"><input type="number" min="1" className={inputCls} style={inputStyle} value={form.nextInvoiceNumber} onChange={(e) => set("nextInvoiceNumber", Number(e.target.value))} /></Field>
        </div>
      </div>

      <div className="rounded-lg p-4 space-y-3" style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="font-semibold text-sm" style={{ color: C.navy }}>Bas de page (mentions modifiables)</div>
        <textarea rows={3} className={inputCls} style={inputStyle} value={form.footer} onChange={(e) => set("footer", e.target.value)} />
      </div>

      <Btn icon={Check} onClick={() => onSave(form)}>Enregistrer les paramètres</Btn>
    </div>
  );
}

/* ============================= APP ============================= */
/* ============================= LOGIN PAGE ============================= */
function LoginPage() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (error) setError("Email ou mot de passe incorrect.");
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-6" style={{ background: C.navy }}>
      <form onSubmit={submit} className="w-full max-w-xs rounded-lg p-6 space-y-3" style={{ background: "#fff" }}>
        <div className="flex items-center gap-2 mb-1">
          <Truck size={20} style={{ color: C.orange }} />
          <span className="font-bold" style={{ color: C.navy }}>Facturation Ground &amp; Rail</span>
        </div>
        <Field label="Email" required>
          <input type="email" required autoFocus className={inputCls} style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} />
        </Field>
        <Field label="Mot de passe" required>
          <input type="password" required className={inputCls} style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} />
        </Field>
        {error && <div className="text-xs" style={{ color: C.red }}>{error}</div>}
        <Btn type="submit" icon={Check} disabled={loading}>{loading ? "Connexion…" : "Se connecter"}</Btn>
        <p className="text-xs pt-1" style={{ color: C.inkMuted }}>Pas de compte ? Demandez à un administrateur de vous en créer un.</p>
      </form>
    </div>
  );
}

/* ============================= APP ============================= */
export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [operations, setOperations] = useState([]);
  const [tariffs, setTariffs] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [settings, setSettings] = useState(defaultSettings);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      const [rawOps, tar, rawInv, set] = await Promise.all([
        loadKey("ceva-operations", []),
        loadKey("ceva-sympos-tariffs", []),
        loadKey("ceva-invoices", []),
        loadKey("ceva-settings", defaultSettings),
      ]);

      // Migration: older data saved "Ordre de transport" before it was shortened to "OT".
      let opsChanged = false;
      const ops = rawOps.map((o) => {
        if (o.refType === "Ordre de transport") { opsChanged = true; return { ...o, refType: "OT" }; }
        return o;
      });
      let invChanged = false;
      const inv = rawInv.map((iv) => {
        const lines = iv.lines.map((l) => {
          if (typeof l.reference === "string" && l.reference.startsWith("Ordre de transport:")) {
            invChanged = true;
            return { ...l, reference: l.reference.replace("Ordre de transport:", "OT:") };
          }
          return l;
        });
        return { ...iv, lines };
      });
      if (opsChanged) saveKey("ceva-operations", ops);
      if (invChanged) saveKey("ceva-invoices", inv);

      setOperations(ops); setTariffs(tar); setInvoices(inv); setSettings(set);
      setLoading(false);
    })();
  }, []);

  const notify = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2500); };

  const addOperation = async (op) => {
    const withId = { ...op, id: uid(), facturee: false, factureId: null, factureNumero: null, createdAt: Date.now() };
    const next = [...operations, withId];
    setOperations(next); await saveKey("ceva-operations", next); notify("Opération enregistrée");
  };
  const updateOperation = async (op) => {
    const next = operations.map((o) => (o.id === op.id ? { ...o, ...op } : o));
    setOperations(next); await saveKey("ceva-operations", next); notify("Opération modifiée");
  };
  const deleteOperation = async (id) => {
    const next = operations.filter((o) => o.id !== id);
    setOperations(next); await saveKey("ceva-operations", next);
  };
  const setEndDate = async (id, dateFin) => {
    const next = operations.map((o) => (o.id === id ? { ...o, dateFin } : o));
    setOperations(next); await saveKey("ceva-operations", next); notify("Date de fin enregistrée");
  };

  const addTariff = async (t) => {
    const next = [...tariffs, t];
    setTariffs(next); await saveKey("ceva-sympos-tariffs", next); notify("Tarif ajouté");
  };
  const deleteTariff = async (id) => {
    const next = tariffs.filter((t) => t.id !== id);
    setTariffs(next); await saveKey("ceva-sympos-tariffs", next);
  };

  const saveSettings = async (s) => {
    setSettings(s); await saveKey("ceva-settings", s); notify("Paramètres enregistrés");
  };

  const createInvoice = async ({ lines, totals, opIds }) => {
    const year = new Date().getFullYear();
    const numero = `${settings.invoicePrefix}-${year}-${String(settings.nextInvoiceNumber).padStart(4, "0")}`;
    const invoice = { id: uid(), numero, date: todayISO(), clientName: settings.clientName, lines, totals, createdAt: Date.now() };
    const nextInvoices = [...invoices, invoice];
    const nextOps = operations.map((o) => opIds.includes(o.id) ? { ...o, facturee: true, factureId: invoice.id, factureNumero: numero } : o);
    const nextSettings = { ...settings, nextInvoiceNumber: settings.nextInvoiceNumber + 1 };
    setInvoices(nextInvoices); setOperations(nextOps); setSettings(nextSettings);
    await Promise.all([
      saveKey("ceva-invoices", nextInvoices),
      saveKey("ceva-operations", nextOps),
      saveKey("ceva-settings", nextSettings),
    ]);
    notify(`Facture ${numero} créée`);
    setTab("invoices");
  };

  const tabs = [
    { key: "dashboard", label: "Tableau de bord", icon: LayoutDashboard },
    { key: "operations", label: "Opérations", icon: ClipboardList },
    { key: "newinvoice", label: "Nouvelle facture", icon: Receipt },
    { key: "invoices", label: "Factures", icon: FileText },
    { key: "tariffs", label: "Tarifs Sympos", icon: MapPin },
    { key: "settings", label: "Paramètres", icon: Settings },
  ];

  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [session, setSession] = useState(undefined); // undefined = checking, null = logged out, object = logged in

  useEffect(() => {
    if (!supabase) { setSession(null); return; }
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
    return () => listener.subscription.unsubscribe();
  }, []);

  if (!supabaseConfigured) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6" style={{ background: C.bg }}>
        <div className="max-w-md rounded-lg p-6 text-sm" style={{ background: C.card, border: `1px solid ${C.border}`, color: C.ink }}>
          <div className="font-bold mb-2" style={{ color: C.navy }}>Configuration manquante</div>
          <p>Les variables d'environnement <code>VITE_SUPABASE_URL</code> et <code>VITE_SUPABASE_ANON_KEY</code> ne sont pas définies. Ajoutez-les dans les paramètres de votre hébergeur (Netlify/Vercel) puis relancez le déploiement — voir README.md.</p>
        </div>
      </div>
    );
  }

  if (session === undefined) {
    return <div className="min-h-screen flex items-center justify-center" style={{ background: C.bg, color: C.navy }}>Vérification de la connexion…</div>;
  }
  if (session === null) {
    return <LoginPage />;
  }

  if (loading) {
    return <div className="min-h-screen flex items-center justify-center" style={{ background: C.bg, color: C.navy }}>Chargement…</div>;
  }

  return (
    <div className="min-h-screen" style={{ background: C.bg, fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
      <div className="no-print" style={{ background: C.navy }}>
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <Truck size={20} style={{ color: C.orange }} className="shrink-0" />
            <div className="min-w-0">
              <div className="text-white font-bold text-sm leading-tight truncate">Facturation Ground &amp; Rail</div>
              <div className="text-xs truncate" style={{ color: "#9FB2CC" }}>Client unique : {settings.clientName}</div>
            </div>
          </div>
          <nav className="hidden md:flex gap-1 flex-wrap items-center">
            {tabs.map((t) => (
              <button
                key={t.key}
                onClick={() => setTab(t.key)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition"
                style={tab === t.key ? { background: C.orange, color: "#fff" } : { background: "transparent", color: "#C7D3E3" }}
              >
                <t.icon size={14} /> {t.label}
              </button>
            ))}
            <button
              title={session?.user?.email}
              onClick={() => supabase.auth.signOut()}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium ml-1"
              style={{ background: "transparent", color: "#C7D3E3" }}
            >
              <LogOut size={14} /> Déconnexion
            </button>
          </nav>
          <button className="md:hidden p-2 rounded-md shrink-0" style={{ background: "rgba(255,255,255,0.1)", color: "#fff" }} onClick={() => setMobileNavOpen((v) => !v)}>
            {mobileNavOpen ? <X size={18} /> : <ChevronDown size={18} />}
          </button>
        </div>
        {mobileNavOpen && (
          <div className="md:hidden px-4 pb-3 flex flex-col gap-1">
            {tabs.map((t) => (
              <button
                key={t.key}
                onClick={() => { setTab(t.key); setMobileNavOpen(false); }}
                className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-medium text-left"
                style={tab === t.key ? { background: C.orange, color: "#fff" } : { background: "rgba(255,255,255,0.06)", color: "#C7D3E3" }}
              >
                <t.icon size={15} /> {t.label}
              </button>
            ))}
            <button
              onClick={() => supabase.auth.signOut()}
              className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-medium text-left"
              style={{ background: "rgba(255,255,255,0.06)", color: "#C7D3E3" }}
            >
              <LogOut size={15} /> Déconnexion
            </button>
          </div>
        )}
      </div>

      <div className="max-w-6xl mx-auto px-3 sm:px-4 py-4 sm:py-6">
        {tab === "dashboard" && <DashboardTab operations={operations} invoices={invoices} />}
        {tab === "operations" && <OperationsTab operations={operations} tariffs={tariffs} onAdd={addOperation} onUpdate={updateOperation} onDelete={deleteOperation} onSetEndDate={setEndDate} />}
        {tab === "newinvoice" && <NewInvoiceTab operations={operations} tariffs={tariffs} settings={settings} onCreate={createInvoice} />}
        {tab === "invoices" && <InvoicesTab invoices={invoices} settings={settings} />}
        {tab === "tariffs" && <TariffsTab tariffs={tariffs} onAdd={addTariff} onDelete={deleteTariff} />}
        {tab === "settings" && <SettingsTab settings={settings} onSave={saveSettings} />}
      </div>

      {toast && (
        <div className="no-print fixed bottom-5 right-5 px-4 py-2.5 rounded-md text-sm font-medium shadow-lg" style={{ background: C.navy, color: "#fff" }}>
          {toast}
        </div>
      )}
    </div>
  );
}
