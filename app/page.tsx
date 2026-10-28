"use client";

import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Package, ShoppingCart, Megaphone, Warehouse,
  Wallet, FileSpreadsheet, Upload, Search, Bell, ChevronDown,
  ArrowUpRight, CircleHelp, CheckCircle2, AlertCircle, Menu
} from "lucide-react";

type Row = Record<string, unknown>;
type WorkbookData = Record<string, Row[]>;

const nav = [
  { label: "经营总览", icon: LayoutDashboard },
  { label: "产品管理", icon: Package },
  { label: "订单明细", icon: ShoppingCart },
  { label: "广告管理", icon: Megaphone },
  { label: "库存管理", icon: Warehouse },
  { label: "费用与回款", icon: Wallet },
  { label: "Excel 数据导入", icon: FileSpreadsheet }
];

const numberValue = (value: unknown) => {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const n = Number(value.replace(/[, SAR]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
};

const money = (n: number) => new Intl.NumberFormat("en-SA", {
  style: "currency", currency: "SAR", maximumFractionDigits: 2
}).format(n);

export default function Home() {
  const [active, setActive] = useState("经营总览");
  const [workbook, setWorkbook] = useState<WorkbookData>({});
  const [fileName, setFileName] = useState("");
  const [message, setMessage] = useState("");
  const [search, setSearch] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const orders = workbook["订单明细"] ?? [];
  const ads = workbook["广告明细"] ?? [];
  const products = workbook["产品主表"] ?? [];
  const inventory = workbook["库存"] ?? [];

  const metrics = useMemo(() => {
    const validOrders = orders.filter(r => r.order_id || r.orderId || r.sku);
    const units = validOrders.reduce((s, r) => s + numberValue(r.units), 0);
    const sales = validOrders.reduce((s, r) =>
      s + numberValue(r.gross_sales_incl_vat_sar ?? r.gross_sales_sar), 0);
    const adSpend = ads.reduce((s, r) =>
      s + numberValue(r.spend_ex_vat_sar ?? r.spend_sar), 0);
    const stocked = inventory.reduce((s, r) => s + numberValue(r.on_hand_available), 0);
    return { orders: validOrders.length, units, sales, adSpend, stocked };
  }, [orders, ads, inventory]);

  async function handleUpload(file?: File) {
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array", cellDates: true });
      const parsed: WorkbookData = {};
      for (const sheetName of wb.SheetNames) {
        if (sheetName.startsWith("旧档_")) continue;
        parsed[sheetName] = XLSX.utils.sheet_to_json<Row>(wb.Sheets[sheetName], { defval: null });
      }
      setWorkbook(parsed);
      setFileName(file.name);
      setMessage("文件读取成功。请核对导入数据后再用于经营决策。");
      setActive("经营总览");
    } catch {
      setMessage("读取失败，请确认文件为有效的 .xlsx、.xls 或 .csv 文件。");
    }
  }

  const activeRows = active === "产品管理" ? products
    : active === "订单明细" ? orders
    : active === "广告管理" ? ads
    : active === "库存管理" ? inventory
    : [];

  const filteredRows = activeRows.filter(row =>
    !search || Object.values(row).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()))
  );
  const columns = activeRows.length ? Object.keys(activeRows[0]).slice(0, 8) : [];

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">SA</div>
          <div><strong>SAUDI OPS</strong><span>SELLER WORKSPACE</span></div>
        </div>
        <div className="market-label">工作空间</div>
        <button className="store-select"><span className="store-dot" /><span><b>Amazon.sa</b><small>沙特阿拉伯 · SAR</small></span><ChevronDown size={16} /></button>
        <div className="nav-label">运营管理</div>
        <nav>
          {nav.map(item => {
            const Icon = item.icon;
            return <button key={item.label} className={active === item.label ? "nav-item active" : "nav-item"} onClick={() => { setActive(item.label); setSearch(""); }}>
              <Icon size={18} strokeWidth={1.8} /><span>{item.label}</span>{item.label === "Excel 数据导入" && <span className="nav-new">开始</span>}
            </button>;
          })}
        </nav>
        <div className="sidebar-bottom">
          <div className="help-card"><CircleHelp size={18} /><div><b>需要帮助？</b><p>系统初版正在搭建中。</p></div></div>
          <div className="profile"><div className="avatar">SA</div><div><b>运营管理员</b><small>个人工作空间</small></div><ChevronDown size={15} /></div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="breadcrumb"><Menu size={18} className="mobile-menu" /><span>工作空间</span><span className="slash">/</span><b>{active}</b></div>
          <div className="top-actions"><span className="status-pill"><i /> 沙特站模式</span><button className="icon-button" aria-label="通知"><Bell size={18} /></button><div className="top-avatar">SA</div></div>
        </header>

        <div className="content">
          <section className="welcome-row">
            <div><div className="eyebrow">AMAZON SAUDI ARABIA · OPERATIONS</div><h1>{active}</h1><p className="subtitle">把销售、广告、库存与资金数据放在一个清晰的工作台中。</p></div>
            <button className="primary-button" onClick={() => inputRef.current?.click()}><Upload size={17} /> 导入 Excel 数据</button>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={e => handleUpload(e.target.files?.[0])} />
          </section>

          {message && <div className="notice"><CheckCircle2 size={17} /><span>{message}</span><button onClick={() => setMessage("")}>关闭</button></div>}

          {active === "经营总览" && <>
            <div className="section-heading"><div><h2>经营概况</h2><p>数据来自当前导入的工作簿；未导入时显示为 0。</p></div><span className="period-chip">当前导入数据</span></div>
            <div className="metric-grid">
              <Metric title="订单记录" value={metrics.orders.toLocaleString()} suffix="条" note="来自订单明细" icon={<ShoppingCart size={19} />} />
              <Metric title="销售额（含税字段）" value={money(metrics.sales)} note="按工作簿字段汇总" icon={<ArrowUpRight size={19} />} />
              <Metric title="广告花费（未税字段）" value={money(metrics.adSpend)} note="按广告明细汇总" icon={<Megaphone size={19} />} />
              <Metric title="可售库存" value={metrics.stocked.toLocaleString()} suffix="件" note="按库存快照汇总" icon={<Warehouse size={19} />} />
            </div>
            <div className="grid-two">
              <section className="panel">
                <div className="panel-heading"><div><h3>数据准备情况</h3><p>上传你的修订版经营数据库开始查看</p></div><FileSpreadsheet size={20} /></div>
                <div className="import-zone" onClick={() => inputRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleUpload(e.dataTransfer.files?.[0]); }}>
                  <div className="upload-icon"><Upload size={22} /></div><b>{fileName || "拖放 Excel 文件到这里"}</b><span>支持 .xlsx、.xls 和 .csv</span><button className="secondary-button">选择文件</button>
                </div>
                <div className="file-footnote"><AlertCircle size={15} /><span>第一版在浏览器本地读取文件，不会自动把工作簿上传到服务器。</span></div>
              </section>
              <section className="panel">
                <div className="panel-heading"><div><h3>经营数据模块</h3><p>围绕沙特站工作簿设计</p></div><span className="count-badge">7 个模块</span></div>
                <div className="module-list">
                  {[
                    ["产品与成本", "产品主表", products.length],
                    ["订单与销售", "订单明细", orders.length],
                    ["广告表现", "广告明细", ads.length],
                    ["库存快照", "库存", inventory.length]
                  ].map(([title, sheet, count]) => <div className="module-row" key={String(sheet)}><div className="module-symbol"><FileSpreadsheet size={17} /></div><div className="module-copy"><b>{title}</b><small>{sheet}</small></div><span className="module-count">{count} 行</span></div>)}
                </div>
              </section>
            </div>
            <section className="panel next-panel">
              <div className="panel-heading"><div><h3>下一步建议</h3><p>先验证数据，再逐步增加自动化能力</p></div></div>
              <div className="steps">
                <div><span>01</span><b>导入修订版工作簿</b><p>读取沙特站运营页，跳过旧档归档页。</p></div>
                <div><span>02</span><b>核对列名与金额口径</b><p>销售额、广告花费和库存按对应字段统计。</p></div>
                <div><span>03</span><b>接入持久化数据库</b><p>确认字段后再加入云端保存、账号和权限。</p></div>
              </div>
            </section>
          </>}

          {active === "Excel 数据导入" && <section className="panel import-page">
            <div className="panel-heading"><div><h3>导入经营工作簿</h3><p>选择本地文件，先在浏览器中解析工作表。</p></div><FileSpreadsheet size={22} /></div>
            <div className="import-zone large" onClick={() => inputRef.current?.click()} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleUpload(e.dataTransfer.files?.[0]); }}>
              <div className="upload-icon"><Upload size={26} /></div><b>{fileName || "点击选择或拖入 Excel 文件"}</b><span>支持 Excel 与 CSV 格式</span><button className="primary-button">选择文件</button>
            </div>
            <div className="warning-box"><AlertCircle size={18} /><p><b>导入说明</b><br />当前为原型版本：文件只在当前浏览器会话中解析，不会保存到云端，也不会自动同步到其他设备。请勿将它视为正式账务或税务系统。</p></div>
            {Object.keys(workbook).length > 0 && <div className="sheet-chips">{Object.entries(workbook).map(([name, rows]) => <span key={name}>{name}<b>{rows.length} 行</b></span>)}</div>}
          </section>}

          {["产品管理", "订单明细", "广告管理", "库存管理"].includes(active) && <section className="panel table-panel">
            <div className="panel-heading"><div><h3>{active}数据</h3><p>从导入工作簿中读取对应工作表的预览。</p></div><span className="count-badge">{activeRows.length} 行</span></div>
            {activeRows.length > 0 ? <>
              <div className="table-toolbar"><div className="search-box"><Search size={16} /><input value={search} onChange={e => setSearch(e.target.value)} placeholder="搜索当前数据..." /></div><span>显示 {filteredRows.length} 条匹配记录</span></div>
              <div className="table-wrap"><table><thead><tr>{columns.map(c => <th key={c}>{c}</th>)}</tr></thead><tbody>{filteredRows.slice(0, 100).map((row, i) => <tr key={i}>{columns.map(c => <td key={c}>{row[c] instanceof Date ? (row[c] as Date).toLocaleDateString() : String(row[c] ?? "—")}</td>)}</tr>)}</tbody></table></div>
              {filteredRows.length > 100 && <p className="table-note">当前预览最多显示 100 行；完整数据未被删除。</p>}
            </> : <div className="empty-state"><FileSpreadsheet size={30} /><b>还没有导入数据</b><p>导入经营数据库后，这里会显示 {active} 的工作表内容。</p><button className="secondary-button" onClick={() => inputRef.current?.click()}>导入 Excel</button></div>}
          </section>}

          {["费用与回款"].includes(active) && <section className="panel empty-state large-empty"><Wallet size={34} /><h3>费用与回款模块</h3><p>下一阶段将接入结算回款、费用核对、现金流与资金占用工作表。</p><button className="primary-button" onClick={() => { setActive("Excel 数据导入"); }}>先导入工作簿</button></section>}
        </div>
        <footer className="footer"><span>SAUDI OPS · MVP</span><span>仅用于数据整理与分析辅助，不替代 Seller Central、会计或税务记录。</span></footer>
      </main>
    </div>
  );
}

function Metric({ title, value, suffix, note, icon }: { title: string; value: string; suffix?: string; note: string; icon: React.ReactNode }) {
  return <div className="metric-card"><div className="metric-top"><span>{title}</span><div className="metric-icon">{icon}</div></div><div className="metric-value">{value}{suffix && <small>{suffix}</small>}</div><div className="metric-note"><span className="muted-dot" />{note}</div></div>;
}