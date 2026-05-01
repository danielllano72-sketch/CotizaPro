
import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { Search, Upload, Plus, Trash2, FileDown, Copy, Save } from "lucide-react";
import "./style.css";

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
  { supplier: "Chrisalim", code: "WIE001", name: "Pastilla Media Luna Wiese 70 grs", unit: "Pieza", cost: 10.70 },
  { supplier: "Chrisalim", code: "P-96", name: "Fibra Verde Scotch Brite", unit: "Pieza", cost: 12.50 },
  { supplier: "Chrisalim", code: "P-76", name: "Fibra Negra Scotch Brite", unit: "Pieza", cost: 16.73 },
  { supplier: "Uniplas", code: "BNR6090", name: "Bolsa negra en rollo 60x90 cm", unit: "Kilo", cost: 42 },
  { supplier: "Uniplas", code: "BNR90120", name: "Bolsa negra en rollo 90x1.20", unit: "Kilo", cost: 42 },
];

function money(n) {
  return new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN" }).format(Number(n || 0));
}
function parseNumber(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  return Number(String(v).replace(/\$/g, "").replace(/,/g, "").trim()) || 0;
}
function normalizeRow(row, supplier = "Importado") {
  const keys = Object.fromEntries(Object.entries(row).map(([k, v]) => [
    String(k).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(), v
  ]));
  const code = keys.codigo || keys.cod || keys.clave || keys.sku || keys["codigo sap"] || "";
  const name = keys.producto || keys.descripcion || keys.description || keys.nombre || keys.articulo || "";
  const unit = keys.unidad || keys.unit || "Pieza";
  const cost = parseNumber(keys.precio || keys.costo || keys["precio unitario"] || keys["costo compra"] || keys.cost);
  return { supplier, code: String(code || ""), name: String(name || ""), unit: String(unit || "Pieza"), cost };
}
function quotedPrice(cost, pct) {
  const p = Number(pct || 0);
  const divisor = 1 - p / 100;
  if (divisor <= 0) return 0;
  return Math.round((Number(cost || 0) / divisor) * 100) / 100;
}
function todayISO() { return new Date().toISOString().slice(0, 10); }
function plusDaysISO(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function App() {
  const fileRef = useRef(null);
  const [products, setProducts] = useState(() => {
  const saved = localStorage.getItem("cotizapro_products");
  return saved ? JSON.parse(saved) : initialProducts;
});
  const [supplier, setSupplier] = useState("GCP");
  const [query, setQuery] = useState("");
  const [globalPct, setGlobalPct] = useState(15);
  const [client, setClient] = useState({ name: "", email: "", company: "" });
  const [folio, setFolio] = useState("COT-000001");
  const [items, setItems] = useState([]);
  useEffect(() => {
  localStorage.setItem("cotizapro_products", JSON.stringify(products));
}, [products]);
  const [issueDate] = useState(todayISO());
  const [dueDate] = useState(plusDaysISO(7));

  const filtered = useMemo(() => {
    const q = query.toLowerCase().trim();
    if (!q) return products.slice(0, 25);
    return products.filter(p => `${p.supplier} ${p.code} ${p.name}`.toLowerCase().includes(q)).slice(0, 25);
  }, [products, query]);

  const subtotal = useMemo(() => items.reduce((s, i) => s + i.lineTotal, 0), [items]);
  const iva = Math.round(subtotal * IVA_RATE * 100) / 100;
  const total = Math.round((subtotal + iva) * 100) / 100;

  async function handleUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const parsed = rows.map(r => normalizeRow(r, supplier)).filter(p => p.name && p.cost > 0);
    setProducts(prev => [...parsed, ...prev]);
    alert(`Importados ${parsed.length} productos de ${supplier}`);
  }

  function addProduct(product) {
    const pct = Number(globalPct || 0);
    const unitPrice = quotedPrice(product.cost, pct);
    setItems(prev => [...prev, { ...product, quantity: 1, pct, unitPrice, lineTotal: unitPrice }]);
  }
  function updateItem(idx, patch) {
    setItems(prev => prev.map((item, i) => {
      if (i !== idx) return item;
      const next = { ...item, ...patch };
      next.unitPrice = quotedPrice(next.cost, next.pct);
      next.lineTotal = Math.round(next.quantity * next.unitPrice * 100) / 100;
      return next;
    }));
  }
  function removeItem(idx) { setItems(prev => prev.filter((_, i) => i !== idx)); }
  function duplicateQuote() {
    setFolio(`COT-${String(Number(folio.replace(/\D/g, "")) + 1).padStart(6, "0")}`);
    alert("Cotización duplicada.");
  }

  function generatePdf() {
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFillColor(12, 29, 46);
    doc.rect(0, 0, pageWidth, 86, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.text("ASEO EMPRESARIAL", 40, 42);
    doc.setFontSize(10);
    doc.text("Tel. 5580013349", 40, 60);
    doc.text("Gardenias No. 50, Col. San Juan Bosco, C.P. 52946, Atizapán de Zaragoza", 40, 74);

    doc.setTextColor(20, 28, 39);
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
      head: [["Cód.", "Producto", "Cant.", "Unidad", "%", "P. Unit.", "Importe"]],
      body: items.map(i => [i.code, i.name, i.quantity, i.unit, `${i.pct}%`, money(i.unitPrice), money(i.lineTotal)]),
      styles: { fontSize: 8, cellPadding: 5 },
      headStyles: { fillColor: [12, 29, 46] },
      columnStyles: { 1: { cellWidth: 180 }, 5: { halign: "right" }, 6: { halign: "right" } },
    });

    const y = doc.lastAutoTable.finalY + 26;
    doc.setFontSize(11);
    doc.text(`Subtotal: ${money(subtotal)}`, 380, y);
    doc.text(`I.V.A. 16%: ${money(iva)}`, 380, y + 18);
    doc.setFontSize(15);
    doc.text(`Total: ${money(total)}`, 380, y + 42);

    doc.setFontSize(10);
    doc.text("Notas", 40, y + 72);
    let noteY = y + 90;
    fixedNotes.forEach(n => {
      const lines = doc.splitTextToSize(`• ${n}`, 500);
      doc.text(lines, 40, noteY);
      noteY += lines.length * 12 + 3;
    });
    doc.save(`${folio}.pdf`);
  }

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand"><div className="mark">AE</div><div><h1>COTIZAPRO</h1><p>Aseo Empresarial</p></div></div>
        <nav><a className="active">Nueva Cotización</a><a>Catálogos</a><a>Clientes</a><a>Historial</a></nav>
      </aside>
      <main className="main">
        <header className="topbar">
          <div><h2>Nueva cotización</h2><p>Flujo rápido para celular y web</p></div>
          <button className="primary" onClick={generatePdf} disabled={!items.length}><FileDown size={18}/> Generar PDF</button>
        </header>

        <section className="grid">
          <div className="card client">
            <h3>Datos del cliente</h3>
            <div className="fields">
              <label>Nombre<input value={client.name} onChange={e=>setClient({...client, name:e.target.value})} placeholder="Gabriela Flores Méndez"/></label>
              <label>Empresa<input value={client.company} onChange={e=>setClient({...client, company:e.target.value})} placeholder="Limpia Tap"/></label>
              <label>Correo<input value={client.email} onChange={e=>setClient({...client, email:e.target.value})} placeholder="correo@empresa.com"/></label>
            </div>
          </div>
          <div className="card quoteData"><h3>Datos de cotización</h3><p><b>Folio:</b> {folio}</p><p><b>Fecha:</b> {issueDate}</p><p><b>Vence:</b> {dueDate}</p></div>
        </section>

        <section className="card">
          <div className="catalogControls">
            <select value={supplier} onChange={e=>setSupplier(e.target.value)}>
              <option>GCP</option><option>Elite</option><option>Uniplas</option><option>Palmer Fixture</option><option>Chrisalim</option><option>Escorpion</option><option>Sra. Julia</option>
            </select>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleUpload} hidden />
            <button onClick={()=>fileRef.current.click()}><Upload size={16}/> Importar lista Excel/CSV</button>
          </div>
        </section>

        <section className="card">
          <div className="searchRow">
            <div className="searchBox"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar por código o nombre..." /></div>
            <label className="pct">Porcentaje general <input type="number" value={globalPct} onChange={e=>setGlobalPct(e.target.value)} />%</label>
          </div>
          <div className="productList">
            {filtered.map((p, idx)=> (
              <button key={`${p.supplier}-${p.code}-${idx}`} onClick={()=>addProduct(p)} className="product">
                <span><b>{p.code || "S/C"}</b> · {p.name}</span>
                <small>{p.supplier} · {p.unit} · Costo {money(p.cost)}</small>
                <Plus size={16}/>
              </button>
            ))}
          </div>
        </section>

        <section className="quoteLayout">
          <div className="card itemsCard">
            <h3>Productos agregados</h3>
            {items.length === 0 ? <p className="empty">Agrega productos para empezar la cotización.</p> : (
              <div className="items">
                {items.map((item, idx)=>(
                  <div className="item" key={idx}>
                    <div className="itemTitle"><b>{item.code}</b> {item.name}</div>
                    <div className="itemControls">
                      <label>Cant.<input type="number" value={item.quantity} onChange={e=>updateItem(idx,{quantity:Number(e.target.value)||1})}/></label>
                      <label>%<input type="number" value={item.pct} onChange={e=>updateItem(idx,{pct:Number(e.target.value)||0})}/></label>
                      <span>{money(item.unitPrice)}</span>
                      <strong>{money(item.lineTotal)}</strong>
                      <button className="danger" onClick={()=>removeItem(idx)}><Trash2 size={16}/></button>
                    </div>
                  </div>
                ))}
              </div>
            )}
            <div className="actions"><button onClick={duplicateQuote}><Copy size={16}/> Duplicar</button><button><Save size={16}/> Guardar</button></div>
          </div>
          <div className="card summary">
            <h3>Resumen</h3>
            <p><span>Subtotal</span><b>{money(subtotal)}</b></p>
            <p><span>IVA 16%</span><b>{money(iva)}</b></p>
            <div className="total"><span>Total</span><strong>{money(total)}</strong></div>
            <button className="primary full" onClick={generatePdf} disabled={!items.length}><FileDown size={18}/> Generar PDF</button>
          </div>
        </section>
      </main>
    </div>
  );
}
createRoot(document.getElementById("root")).render(<App />);
