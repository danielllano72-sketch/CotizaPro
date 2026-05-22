import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Search, Upload, Plus, Trash2, FileDown, Copy, Save } from "lucide-react";
import "./style.css";
import { supabase } from "./supabase";

const IVA_RATE = 0.16;

const fixedNotes = [
  "Condiciones de pago anticipado a la entrega",
  "Forma de pago transferencia",
  "Entrega de 3 a 5 días",
  "Pedidos mínimos de entrega a domicilio de $5000 más IVA CDMX y área metropolitana; fuera de esta demarcación el envío corre por cuenta del cliente.",
  "Nos gustaría ser parte de su cadena de proveedores llevando Ahorro, Calidad y Servicio.",
];

const initialProducts = [
  { supplier: "GCP", code: "14832", name: "Higiénico en Bobina GCP Professional Blanco HD c/12 Rollos 200 m", unit: "Caja", cost: 248 },
  { supplier: "GCP", code: "30302", name: "Toalla Interdoblada GCP Professional Premium c/20 Paquetes 100 Hojas", unit: "Caja", cost: 194 },
  { supplier: "Sra. Julia", code: "QASEM002", name: "Jabón Gel Asempre p/manos Almendras", unit: "Litro", cost: 10 },
  { supplier: "Sra. Julia", code: "MQ0211", name: "Multiusos Asempre Lavanda", unit: "Litro", cost: 5 },
  { supplier: "Sra. Julia", code: "QASEM013", name: "Cloro Asempre al 6%", unit: "Litro", cost: 6 },
];

function safeParse(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function money(n) {
  return new Intl.NumberFormat("es-MX", {
    style: "currency",
    currency: "MXN",
  }).format(Number(n || 0));
}

function parseNumber(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  return Number(String(v).replace(/\$/g, "").replace(/,/g, "").trim()) || 0;
}

function normalizeText(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}
function normalizeRow(row, supplier = "Importado") {
  const keys = Object.fromEntries(
    Object.entries(row).map(([k, v]) => [
      String(k).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(),
      v,
    ])
  );

  return {
    supplier,
    code: String(keys.codigo || keys.cod || keys.clave || keys.sku || ""),
    name: String(keys.producto || keys.descripcion || keys.nombre || keys.articulo || ""),
    unit: String(keys.unidad || keys.unit || "Pieza"),
    cost: parseNumber(keys.precio || keys.costo || keys["precio unitario"] || keys["costo compra"]),
  };
}

function quotedPrice(cost, pct) {
  const divisor = 1 - Number(pct || 0) / 100;
  if (divisor <= 0) return 0;
  return Math.round((Number(cost || 0) / divisor) * 100) / 100;
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function plusDaysISO(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function getNumberFromFolio(id) {
  return Number(String(id || "").replace(/\D/g, "")) || 0;
}

function makeFolio(num) {
  return `COT-${String(num).padStart(6, "0")}`;
}

function getNextFolioFromQuotes(quotes) {
  const nums = quotes.map((q) => getNumberFromFolio(q.id));
  return makeFolio(Math.max(0, ...nums) + 1);
}

function fixDuplicateFolios(quotes) {
  const used = new Set();

  return quotes.map((q, index) => {
    let id = q.id || makeFolio(index + 1);

    if (used.has(id)) id = makeFolio(index + 1);

    while (used.has(id)) {
      id = makeFolio(getNumberFromFolio(id) + 1);
    }

    used.add(id);
    return { ...q, id };
  });
}

function App() {
  const fileRef = useRef(null);

  const [products, setProducts] = useState(() => {
    const saved = safeParse(localStorage.getItem("cotizapro_products"), null);
    return Array.isArray(saved) && saved.length > 0 ? saved : initialProducts;
  });

  const [supplier, setSupplier] = useState("GCP");
  const [query, setQuery] = useState("");
  const [historyQuery, setHistoryQuery] = useState("");
  const [clientQuery, setClientQuery] = useState("");
  const [selectedClientId, setSelectedClientId] = useState(null);
  const [globalPct, setGlobalPct] = useState(15);
  const [client, setClient] = useState({ name: "", email: "", company: "" });
  const [view, setView] = useState("quote");
  const [quotes, setQuotes] = useState([]);
  const [clients, setClients] = useState([]);
  const [folio, setFolio] = useState("COT-000001");
  const [items, setItems] = useState([]);
  const [issueDate] = useState(todayISO());
  const [dueDate] = useState(plusDaysISO(7));
  const [session, setSession] = useState(null);
const [authEmail, setAuthEmail] = useState("");
const [authPassword, setAuthPassword] = useState("");
const [authMode, setAuthMode] = useState("login");
const [authError, setAuthError] = useState("");

  useEffect(() => {
  async function loadData() {
    const { data: quotesData, error: quotesError } = await supabase
      .from("quotes")
      .select("*")
      .order("created_at", { ascending: true });

    if (!quotesError) {
      const fixedQuotes = fixDuplicateFolios(
        (quotesData || []).map((q) => ({
          id: q.id,
          client: q.client || {},
          items: q.items || [],
          subtotal: q.subtotal,
          iva: q.iva,
          total: q.total,
          globalPct: q.globalpct ?? 15,
          issueDate: q.issuedate,
          dueDate: q.duedate,
          createdAt: q.created_at,
        }))
      );

      setQuotes(fixedQuotes);
      setFolio(getNextFolioFromQuotes(fixedQuotes));
    }

    const { data: clientsData, error: clientsError } = await supabase
      .from("clients")
      .select("*")
      .order("created_at", { ascending: true });

    if (!clientsError) {
      setClients(
        (clientsData || []).map((c) => ({
          id: c.id,
          name: c.name || "",
          company: c.company || "",
          email: c.email || "",
          createdAt: c.created_at,
        }))
      );
    }
  }

  loadData();
}, []);
  useEffect(() => {
  supabase.auth.getSession().then(({ data }) => {
    setSession(data.session);
  });

  const {
    data: { subscription }
  } = supabase.auth.onAuthStateChange((_event, session) => {
    setSession(session);
  });

  return () => subscription.unsubscribe();
}, []);

  useEffect(() => {
  async function loadProducts() {
    const { data, error } = await supabase
      .from("products")
      .select("*")
      .order("created_at", { ascending: true });

    if (error) {
      console.error("Error cargando productos:", error.message);
      return;
    }

    if (data && data.length > 0) {
      setProducts(
        data.map((p) => ({
          supplier: p.supplier || "",
          code: p.code || "",
          name: p.name || "",
          unit: p.unit || "Pieza",
          cost: Number(p.cost || 0),
        }))
      );
    }
  }

  if (session) loadProducts();
}, [session]);

  const filtered = useMemo(() => {
  const q = normalizeText(query);

  if (!q) return products.slice(0, 50);

  const terms = q.split(" ").filter(Boolean);

  return products
    .map((p) => {
      const searchable = normalizeText(`
        ${p.supplier}
        ${p.code}
        ${p.name}
        ${p.unit}
      `);

      let score = 0;

      terms.forEach((term) => {
        if (searchable.includes(term)) score += 1;

        if (normalizeText(p.name).startsWith(term)) score += 3;

        if (normalizeText(p.code) === term) score += 10;
      });

      return {
        ...p,
        score,
      };
    })
    .filter((p) => p.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 50);
}, [products, query]);

  const filteredQuotes = useMemo(() => {
    const q = historyQuery.toLowerCase().trim();
    if (!q) return quotes;

    return quotes.filter((quote) => {
      const text = `${quote.id} ${quote.client?.name || ""} ${quote.client?.company || ""} ${quote.total || ""}`.toLowerCase();
      return text.includes(q);
    });
  }, [quotes, historyQuery]);

  const filteredClients = useMemo(() => {
    const q = clientQuery.toLowerCase().trim();
    if (!q) return clients;

    return clients.filter((c) => {
      const text = `${c.name || ""} ${c.company || ""} ${c.email || ""}`.toLowerCase();
      return text.includes(q);
    });
  }, [clients, clientQuery]);

  const subtotal = useMemo(() => items.reduce((s, i) => s + Number(i.lineTotal || 0), 0), [items]);
  const iva = Math.round(subtotal * IVA_RATE * 100) / 100;
  const total = Math.round((subtotal + iva) * 100) / 100;

  async function handleUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const parsed = rows
      .map((r) => normalizeRow(r, supplier))
      .filter((p) => p.name && p.cost > 0);

    const { error: deleteError } = await supabase
  .from("products")
  .delete()
  .eq("supplier", supplier);

if (deleteError) {
  alert("Error al limpiar catálogo anterior: " + deleteError.message);
  return;
}

const { error: insertError } = await supabase
  .from("products")
  .insert(
    parsed.map((p) => ({
      supplier: p.supplier,
      code: p.code,
      name: p.name,
      unit: p.unit,
      cost: p.cost,
    }))
  );

if (insertError) {
  alert("Error al guardar productos: " + insertError.message);
  return;
}

setProducts((prev) => {
  const sinEsteProveedor = prev.filter((p) => p.supplier !== supplier);
  return [...sinEsteProveedor, ...parsed];
});

alert(`Importados ${parsed.length} productos de ${supplier} y guardados en Supabase`);
}
  function addProduct(product) {
    const pct = Number(globalPct || 0);
    const unitPrice = quotedPrice(product.cost, pct);

    setItems((prev) => [
      ...prev,
      {
        ...product,
        quantity: 1,
        pct,
        unitPrice,
        lineTotal: unitPrice,
      },
    ]);
  }

  function updateItem(idx, patch) {
    setItems((prev) =>
      prev.map((item, i) => {
        if (i !== idx) return item;

        const next = { ...item, ...patch };
        next.unitPrice = quotedPrice(next.cost, next.pct);
        next.lineTotal = Math.round(Number(next.quantity || 0) * Number(next.unitPrice || 0) * 100) / 100;

        return next;
      })
    );
  }

  function removeItem(idx) {
    setItems((prev) => prev.filter((_, i) => i !== idx));
  }

  function newQuote() {
  setFolio(getNextFolioFromQuotes(quotes));
  setClient({ name: "", email: "", company: "" });
  setItems([]);
  setGlobalPct(15);
  setView("quote");
}

function duplicateQuote() {
  setFolio(getNextFolioFromQuotes(quotes));
  alert("Cotización duplicada como nueva.");
}

async function saveClient() {
  if (!client.name && !client.company) {
    alert("Agrega nombre o empresa del cliente.");
    return;
  }

  const exists = clients.some(
    (c) =>
      c.email &&
      client.email &&
      c.email.toLowerCase() === client.email.toLowerCase()
  );

  if (exists) {
    alert("Ese cliente ya existe por correo.");
    return;
  }

  const { data, error } = await supabase
    .from("clients")
    .insert({
      name: client.name,
      company: client.company,
      email: client.email,
    })
    .select()
    .single();

  if (error) {
    alert("Error al guardar cliente: " + error.message);
    return;
  }

  setClients([
    ...clients,
    {
      id: data.id,
      name: data.name || "",
      company: data.company || "",
      email: data.email || "",
      createdAt: data.created_at,
    },
  ]);

  alert("Cliente guardado correctamente.");
}
  function useClient(c) {
    setClient({
      name: c.name || "",
      company: c.company || "",
      email: c.email || "",
    });
    setView("quote");
  }

  async function deleteClient(id) {
  if (!confirm("¿Eliminar este cliente?")) return;

  const { error } = await supabase.from("clients").delete().eq("id", id);

  if (error) {
    alert("Error al eliminar cliente: " + error.message);
    return;
  }

  setClients(clients.filter((c) => c.id !== id));
}
async function saveQuote() {
  if (!client.name && !client.company) {
    alert("Agrega al menos nombre o empresa del cliente.");
    return;
  }

  if (items.length === 0) {
    alert("Agrega productos antes de guardar.");
    return;
  }

  let finalFolio = folio;

  if (quotes.some((q) => q.id === finalFolio)) {
    finalFolio = getNextFolioFromQuotes(quotes);
  }

  const quote = {
    id: finalFolio,
    client,
    issueDate,
    dueDate,
    items,
    subtotal,
    iva,
    total,
    globalPct,
    createdAt: new Date().toISOString(),
  };

  const { error } = await supabase.from("quotes").upsert({
    id: finalFolio,
    client,
    items,
    subtotal,
    iva,
    total,
    globalpct: globalPct,
    issuedate: issueDate,
    duedate: dueDate,
  });

  if (error) {
    alert("Error al guardar cotización: " + error.message);
    return;
  }

  const updatedQuotes = fixDuplicateFolios([...quotes, quote]);
  setQuotes(updatedQuotes);

  const clientExists = clients.some(
    (c) =>
      (client.email && c.email?.toLowerCase() === client.email.toLowerCase()) ||
      (client.company && c.company?.toLowerCase() === client.company.toLowerCase())
  );

  if (!clientExists && (client.name || client.company)) {
    const { data: newClientData } = await supabase
      .from("clients")
      .insert({
        name: client.name,
        company: client.company,
        email: client.email,
      })
      .select()
      .single();

    if (newClientData) {
      setClients([
        ...clients,
        {
          id: newClientData.id,
          name: newClientData.name || "",
          company: newClientData.company || "",
          email: newClientData.email || "",
          createdAt: newClientData.created_at,
        },
      ]);
    }
  }

  alert(`Cotización ${finalFolio} guardada correctamente.`);
  setFolio(getNextFolioFromQuotes(updatedQuotes));
}
  function loadQuote(q) {
    setFolio(q.id);
    setClient(q.client || { name: "", email: "", company: "" });
    setItems(q.items || []);
    setGlobalPct(q.globalPct || 15);
    setView("quote");
  }

  async function deleteQuote(id) {
  if (!confirm(`¿Eliminar la cotización ${id}?`)) return;

  const { error } = await supabase.from("quotes").delete().eq("id", id);

  if (error) {
    alert("Error al eliminar cotización: " + error.message);
    return;
  }

  const updatedQuotes = quotes.filter((q) => q.id !== id);
  setQuotes(updatedQuotes);
  setFolio(getNextFolioFromQuotes(updatedQuotes));
}
function getClientQuotes(c) {
  return quotes.filter((q) => {
    const qc = q.client || {};
    return (
      (c.email && qc.email && c.email.toLowerCase() === qc.email.toLowerCase()) ||
      (c.company && qc.company && c.company.toLowerCase() === qc.company.toLowerCase()) ||
      (c.name && qc.name && c.name.toLowerCase() === qc.name.toLowerCase())
    );
  });
}
  async function signIn() {
  setAuthError("");

  const { error } = await supabase.auth.signInWithPassword({
    email: authEmail,
    password: authPassword
  });

  if (error) setAuthError(error.message);
}

async function signUp() {
  setAuthError("");

  const { error } = await supabase.auth.signUp({
    email: authEmail,
    password: authPassword
  });

  if (error) setAuthError(error.message);
}

async function signOut() {
  await supabase.auth.signOut();
}
  function generatePdf() {
    if (items.length === 0) {
      alert("Agrega productos antes de generar PDF.");
      return;
    }

    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();

    doc.setFillColor(15, 15, 15);
    doc.rect(0, 0, pageWidth, 86, "F");
    doc.setDrawColor(255, 102, 0);
doc.setLineWidth(2);
doc.line(40, 92, 555, 92);

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.text("ASEO EMPRESARIAL", 40, 42);

    doc.setFontSize(10);
    doc.text("Tel. 5580013349", 40, 60);
    doc.text("Gardenias No. 50, Col. San Juan Bosco, C.P. 52946, Atizapán de Zaragoza", 40, 74);

    doc.setTextColor(255, 102, 0);
    doc.setFontSize(18);
    doc.text("COTIZACIÓN", 40, 120);

    doc.setFontSize(10);
    doc.text(`Número: ${folio}`, 400, 112);
    doc.text(`Fecha: ${issueDate}`, 400, 128);
    doc.text(`Vencimiento: ${dueDate}`, 400, 144);

    doc.setFontSize(11);
    doc.text("Receptor", 40, 158);

    doc.setFontSize(10);
    doc.text(`Empresa: ${client.company || "Sin empresa"}`, 40, 176);
    doc.text(`Contacto: ${client.name || "Sin nombre"}`, 40, 192);
    doc.text(`Correo: ${client.email || "Sin correo"}`, 40, 208);

    autoTable(doc, {
      startY: 232,
      head: [["Cód.", "Producto", "Cant.", "Unidad", "P. Unit.", "Importe"]],
      body: items.map((i) => [
        i.code || "S/C",
        i.name,
        i.quantity,
        i.unit,
        money(i.unitPrice),
        money(i.lineTotal),
      ]),
      styles: { fontSize: 8, cellPadding: 5 },
      headStyles: { fillColor: [255, 102, 0] },
      columnStyles: {
        1: { cellWidth: 190 },
        4: { halign: "right" },
        5: { halign: "right" },
      },
    });

    const y = doc.lastAutoTable.finalY + 26;

    doc.setFontSize(11);
    doc.text(`Subtotal: ${money(subtotal)}`, 380, y);
    doc.text(`I.V.A. 16%: ${money(iva)}`, 380, y + 18);
    
doc.setTextColor(255, 102, 0);
doc.setFont(undefined, "bold");
    
    doc.setFontSize(15);
    doc.text(`Total: ${money(total)}`, 380, y + 42);
    doc.setTextColor(25, 25, 25);
doc.setFont(undefined, "normal");

    doc.setFontSize(10);
    doc.text("Notas", 40, y + 72);

    let noteY = y + 90;
    fixedNotes.forEach((n) => {
      const lines = doc.splitTextToSize(`• ${n}`, 500);
      doc.text(lines, 40, noteY);
      noteY += lines.length * 12 + 3;
    });

    doc.save(`${folio}.pdf`);
  }
if (!session) {
  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        background: "#f4f6f8",
        padding: 20
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 400,
          background: "#fff",
          padding: 32,
          borderRadius: 18,
          boxShadow: "0 10px 30px rgba(0,0,0,0.08)"
        }}
      >
        <h1 style={{ marginBottom: 8 }}>COTIZAPRO</h1>

        <p style={{ color: "#666", marginBottom: 24 }}>
          Acceso privado
        </p>

        <input
          type="email"
          placeholder="Correo"
          value={authEmail}
          onChange={(e) => setAuthEmail(e.target.value)}
          style={{
            width: "100%",
            padding: 12,
            marginBottom: 12,
            borderRadius: 10,
            border: "1px solid #ddd"
          }}
        />

        <input
          type="password"
          placeholder="Contraseña"
          value={authPassword}
          onChange={(e) => setAuthPassword(e.target.value)}
          style={{
            width: "100%",
            padding: 12,
            marginBottom: 12,
            borderRadius: 10,
            border: "1px solid #ddd"
          }}
        />

        {authError && (
          <p style={{ color: "red", marginBottom: 12 }}>
            {authError}
          </p>
        )}

        <button
          onClick={authMode === "login" ? signIn : signUp}
          style={{
            width: "100%",
            padding: 12,
            borderRadius: 10,
            border: "none",
            background: "#0c1d2e",
            color: "#fff",
            fontWeight: "bold",
            cursor: "pointer"
          }}
        >
          {authMode === "login"
            ? "Iniciar sesión"
            : "Crear cuenta"}
        </button>

        <button
          onClick={() =>
            setAuthMode(
              authMode === "login" ? "signup" : "login"
            )
          }
          style={{
            width: "100%",
            padding: 12,
            borderRadius: 10,
            border: "none",
            background: "transparent",
            marginTop: 10,
            cursor: "pointer"
          }}
        >
          {authMode === "login"
            ? "Crear usuario"
            : "Ya tengo cuenta"}
        </button>
      </div>
    </div>
  );
}
  
  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand">
          <div className="mark">AE</div>
          <div>
            <h1>COTIZAPRO</h1>
            <p>Aseo Empresarial</p>
          </div>
        </div>

        <nav>
          <a className={view === "quote" ? "active" : ""} onClick={() => setView("quote")}>
            Nueva Cotización
          </a>
          <a>Catálogos</a>
          <a className={view === "clients" ? "active" : ""} onClick={() => setView("clients")}>
            Clientes
          </a>
          <a className={view === "history" ? "active" : ""} onClick={() => setView("history")}>
            Historial
          </a>
          <button
  onClick={signOut}
  style={{
    marginTop: 20,
    width: "100%"
  }}
>
  Cerrar sesión
</button>
        </nav>
      </aside>

      <main className="main">
        {view === "quote" && (
          <>
            <header className="topbar">
              <div>
                <h2>Nueva cotización</h2>
                <p>Flujo rápido para celular y web</p>
              </div>

              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={newQuote}>Nueva</button>
                <button className="primary" onClick={generatePdf} disabled={!items.length}>
                  <FileDown size={18} /> Generar PDF
                </button>
              </div>
            </header>

            <section className="grid">
              <div className="card client" style={{ position: "relative", zIndex: 10 }}>
                <h3>Datos del cliente</h3>

                <div className="fields">
                  <label>
                    Nombre
                    <input value={client.name} onChange={(e) => setClient({ ...client, name: e.target.value })} placeholder="Nombre del contacto" />
                  </label>

                  <label>
                    Empresa
                    <input value={client.company} onChange={(e) => setClient({ ...client, company: e.target.value })} placeholder="Empresa" />
                  </label>

                  <label>
                    Correo
                    <input value={client.email} onChange={(e) => setClient({ ...client, email: e.target.value })} placeholder="correo@empresa.com" />
                  </label>
                </div>

                <button onClick={saveClient} style={{ marginTop: 12 }}>
                  <Save size={16} /> Guardar cliente
                </button>
              </div>

              <div className="card quoteData">
                <h3>Datos de cotización</h3>
                <p><b>Folio:</b> {folio}</p>
                <p><b>Fecha:</b> {issueDate}</p>
                <p><b>Vence:</b> {dueDate}</p>
              </div>
            </section>

            <section className="card">
              <div className="catalogControls">
                <select value={supplier} onChange={(e) => setSupplier(e.target.value)}>
                  <option>GCP</option>
                  <option>Elite</option>
                  <option>Uniplas</option>
                  <option>Palmer Fixture</option>
                  <option>Chrisalim</option>
                  <option>Escorpion</option>
                  <option>Sra. Julia</option>
                </select>

                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} hidden />

                <button onClick={() => fileRef.current.click()}>
                  <Upload size={16} /> Importar lista Excel/CSV
                </button>
              </div>
            </section>

            <section className="card">
              <div className="searchRow">
                <div className="searchBox">
                  <Search size={18} />
                  <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar por código o nombre..." />
                </div>

                <label className="pct">
                  Porcentaje general
                  <input type="number" value={globalPct} onChange={(e) => setGlobalPct(e.target.value)} />
                  %
                </label>
              </div>

              <div className="productList">
                {filtered.map((p, idx) => (
                  <button key={`${p.supplier}-${p.code}-${idx}`} onClick={() => addProduct(p)} className="product">
                    <span><b>{p.code || "S/C"}</b> · {p.name}</span>
                    <small>{p.supplier} · {p.unit} · Costo {money(p.cost)}</small>
                    <Plus size={16} />
                  </button>
                ))}
              </div>
            </section>

            <section className="quoteLayout">
              <div className="card itemsCard">
                <h3>Productos agregados</h3>

                {items.length === 0 ? (
                  <p className="empty">Agrega productos para empezar la cotización.</p>
                ) : (
                  <div className="items">
                    {items.map((item, idx) => (
                      <div className="item" key={idx}>
                        <div className="itemTitle">
                          <b>{item.code}</b> {item.name}
                        </div>

                        <div className="itemControls">
                          <label>
                            Cant.
                            <input type="number" value={item.quantity} onChange={(e) => updateItem(idx, { quantity: Number(e.target.value) || 1 })} />
                          </label>

                          <label>
                            %
                            <input type="number" value={item.pct} onChange={(e) => updateItem(idx, { pct: Number(e.target.value) || 0 })} />
                          </label>

                          <span>{money(item.unitPrice)}</span>
                          <strong>{money(item.lineTotal)}</strong>

                          <button className="danger" onClick={() => removeItem(idx)}>
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                <div className="actions">
                  <button onClick={duplicateQuote}>
                    <Copy size={16} /> Duplicar
                  </button>

                  <button onClick={saveQuote} disabled={!items.length}>
                    <Save size={16} /> Guardar
                  </button>
                </div>
              </div>

              <div className="card summary">
                <h3>Resumen</h3>
                <p><span>Subtotal</span><b>{money(subtotal)}</b></p>
                <p><span>IVA 16%</span><b>{money(iva)}</b></p>

                <div className="total">
                  <span>Total</span>
                  <strong>{money(total)}</strong>
                </div>

                <button className="primary full" onClick={generatePdf} disabled={!items.length}>
                  <FileDown size={18} /> Generar PDF
                </button>
              </div>
            </section>
          </>
        )}

        {view === "clients" && (
  <div className="card">
    <h3>Clientes registrados</h3>

    <div className="searchBox" style={{ margin: "12px 0" }}>
      <Search size={18} />
      <input
        value={clientQuery}
        onChange={(e) => setClientQuery(e.target.value)}
        placeholder="Buscar por nombre, empresa o correo..."
      />
    </div>

    {filteredClients.length === 0 && <p>No hay clientes registrados</p>}

    {filteredClients.map((c) => {
      const clientQuotes = getClientQuotes(c);
      const isOpen = selectedClientId === c.id;

      return (
        <div
          key={c.id}
          style={{
            borderBottom: "1px solid #eee",
            padding: "12px 0",
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: 12,
            }}
          >
            <button
              type="button"
              onClick={() => setSelectedClientId(isOpen ? null : c.id)}
              style={{
                background: "transparent",
                border: "0",
                textAlign: "left",
                width: "100%",
                cursor: "pointer",
              }}
            >
              <strong>{c.company || "Sin empresa"}</strong>
              <div>{c.name || "Sin nombre"}</div>
              <div>{c.email || "Sin correo"}</div>
              <small>{clientQuotes.length} cotización(es)</small>
            </button>

            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => useClient(c)}>Usar</button>
              <button className="danger" onClick={() => deleteClient(c.id)}>
                <Trash2 size={16} />
              </button>
            </div>
          </div>

          {isOpen && (
            <div style={{ marginTop: 12, paddingLeft: 16 }}>
              <h4>Cotizaciones realizadas</h4>

              {clientQuotes.length === 0 && (
                <p>No hay cotizaciones para este cliente.</p>
              )}

              {clientQuotes.map((q) => (
                <button
                  key={q.id}
                  type="button"
                  onClick={() => loadQuote(q)}
                  style={{
                    display: "block",
                    width: "100%",
                    background: "#f8fafc",
                    border: "1px solid #eee",
                    borderRadius: 10,
                    padding: 10,
                    marginBottom: 8,
                    textAlign: "left",
                    cursor: "pointer",
                  }}
                >
                  <strong>{q.id}</strong>
                  <div>Total: {money(q.total)}</div>
                  <div>Fecha: {q.issueDate || "Sin fecha"}</div>
                  <div>Porcentaje: {q.globalPct}%</div>
                </button>
              ))}
            </div>
          )}
        </div>
      );
    })}
  </div>
)}
        {view === "history" && (
          <div className="card">
            <h3>Historial de cotizaciones</h3>

            <div className="searchBox" style={{ margin: "12px 0" }}>
              <Search size={18} />
              <input value={historyQuery} onChange={(e) => setHistoryQuery(e.target.value)} placeholder="Buscar por folio, cliente o empresa..." />
            </div>

            {filteredQuotes.length === 0 && <p>No hay cotizaciones guardadas</p>}

            {filteredQuotes.map((q) => (
              <div
                key={q.id}
                style={{
                  borderBottom: "1px solid #eee",
                  padding: "12px 0",
                  display: "flex",
                  justifyContent: "space-between",
                  gap: 12,
                  alignItems: "center",
                }}
              >
                <button
                  type="button"
                  onClick={() => loadQuote(q)}
                  style={{
                    background: "transparent",
                    border: "0",
                    textAlign: "left",
                    width: "100%",
                    cursor: "pointer",
                  }}
                >
                  <strong>{q.id}</strong>
                  <div>{q.client?.name || "Sin nombre"} - {q.client?.company || ""}</div>
                  <div>Total: {money(q.total)}</div>
                  <div>Porcentaje: {q.globalPct}%</div>
                </button>

                <button className="danger" onClick={() => deleteQuote(q.id)}>
                  <Trash2 size={16} />
                </button>
              </div>
            ))}
          </div>
        )}
      </main>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
