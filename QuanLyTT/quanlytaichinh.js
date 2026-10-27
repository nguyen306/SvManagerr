/* ======================================================================
   Quản lý tài chính: sinh viên ĐÃ nộp đủ / CHƯA nộp đủ học phí.
   - Bọc trong IIFE nên không đụng biến của các phần khác.
   - Cấu hình qua thuộc tính của <section id="tc">:
       data-excel="script.xlsx"   đường dẫn file Excel
       data-price="450000"        đơn giá mặc định / tín chỉ
   - Cần thư viện SheetJS (XLSX) nạp trước.
   - Tự chạy khi trang tải xong; cũng gọi được thủ công: TaiChinh.init()
   ====================================================================== */
(function () {
  "use strict";

  const LS_PRICE = "tc_price";
  const LS_PAID = "tc_paid";

  let root, EXCEL_URL, DEFAULT_PRICE;
  let students = [];      // [{id, name, cls}]
  let totalCredits = 0;
  let price = 0;
  let payments = {};      // { msv: {t:"full"} | {t:"part", amt:number} }
  let filter = "all";
  let started = false;

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
  const money = (n) => Math.round(n).toLocaleString("vi-VN") + " đ";
  const norm = (v) => String(v == null ? "" : v).trim().toLowerCase();
  const stripTone = (s) => norm(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d");

  const load = (k, fb) => { try { const v = localStorage.getItem(k); return v == null ? fb : JSON.parse(v); } catch (e) { return fb; } };
  const save = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };

  /* ---------- Tính toán ---------- */
  const dueOf = () => totalCredits * price;

  function paidOf(s) {
    const p = payments[s.id];
    if (!p) return 0;
    return p.t === "full" ? dueOf() : Math.min(p.amt, dueOf());
  }
  function statusOf(s) {
    const due = dueOf(), paid = paidOf(s);
    if (due > 0 && paid >= due) return "paid";
    return paid > 0 ? "part" : "none";
  }

  /* Dữ liệu mẫu ổn định theo MSV: ~62% nộp đủ, ~13% nộp thiếu, còn lại chưa nộp */
  function seedPayments() {
    const out = {};
    students.forEach((s) => {
      const r = ((Number(s.id.replace(/\D/g, "")) * 9301 + 49297) % 233280) / 233280;
      if (r < 0.62) out[s.id] = { t: "full" };
      else if (r < 0.75) out[s.id] = { t: "part", amt: Math.round((0.3 + (r * 1000 % 50) / 100) * dueOf() / 10000) * 10000 };
    });
    return out;
  }

  /* ---------- Đọc Excel ---------- */
  function parseWorkbook(wb) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
    if (rows.length < 3) throw new Error("File quá ít dòng.");
    const h0 = rows[0] || [], h1 = rows[1] || [];
    const find = (kw, fb) => { const i = h0.findIndex((c) => norm(c).includes(kw)); return i >= 0 ? i : fb; };
    const iName = find("họ và tên", 1), iId = find("msv", 2), iCls = find("lớp", 3);
    const gpa = h0.findIndex((c) => norm(c) === "gpa");
    const start = Math.max(iName, iId, iCls, find("email", 4), find("điện thoại", 5)) + 1;
    const end = gpa >= 0 ? gpa : h1.length;

    let credits = 0;
    for (let c = start; c < end; c++) {
      const raw = String(h1[c] == null ? "" : h1[c]).trim();
      const m = raw.match(/\((\d+(?:[.,]\d+)?)\)\s*$/);
      if (m) credits += Number(m[1].replace(",", "."));
    }
    if (!credits) throw new Error("Không tìm thấy số tín chỉ ở dòng 2.");

    const list = [];
    rows.slice(2).forEach((r) => {
      if (!r || !r[iName]) return;
      const id = typeof r[iId] === "number" ? String(Math.round(r[iId])) : String(r[iId] == null ? "" : r[iId]).trim();
      list.push({ id, name: String(r[iName]).trim(), cls: String(r[iCls] == null ? "" : r[iCls]).trim() });
    });
    if (!list.length) throw new Error("Không đọc được sinh viên nào.");
    return { list, credits };
  }

  function setMsg(text, type) { $("tcMsg").textContent = text; $("tcMsg").className = "tc-msg" + (type ? " " + type : ""); }

  function ingest(buffer, label) {
    try {
      const { list, credits } = parseWorkbook(XLSX.read(new Uint8Array(buffer), { type: "array" }));
      students = list;
      totalCredits = credits;
      payments = load(LS_PAID, null);
      if (!payments) { payments = seedPayments(); save(LS_PAID, payments); }
      const classes = [...new Set(students.map((s) => s.cls))].sort((a, b) => a.localeCompare(b, "vi"));
      $("tcClass").innerHTML = `<option value="">Tất cả các lớp</option>` + classes.map((c) => `<option>${esc(c)}</option>`).join("");
      setMsg(`Đã nạp "${label}": ${students.length} sinh viên, ${totalCredits} tín chỉ toàn khóa.`);
      render();
    } catch (err) {
      setMsg("Lỗi: " + err.message, "error");
    }
  }

  function autoLoad() {
    if (typeof XLSX === "undefined") { setMsg("Không tải được thư viện đọc Excel (SheetJS). Kiểm tra kết nối mạng.", "error"); return; }
    fetch(EXCEL_URL)
      .then((r) => { if (!r.ok) throw new Error("HTTP " + r.status); return r.arrayBuffer(); })
      .then((b) => ingest(b, EXCEL_URL))
      .catch(() => setMsg(`Không đọc được "${EXCEL_URL}" (hãy chạy bằng Live Server) – hoặc bấm "Chọn file Excel" để nạp thủ công.`, "error"));
  }

  /* ---------- Hiển thị ---------- */
  function filtered() {
    const kw = stripTone($("tcSearch").value);
    const cls = $("tcClass").value;
    return students.filter((s) => {
      if (cls && s.cls !== cls) return false;
      if (kw && !(stripTone(s.name).includes(kw) || s.id.includes(kw))) return false;
      const st = statusOf(s);
      if (filter === "paid") return st === "paid";
      if (filter === "unpaid") return st !== "paid";
      return true;
    });
  }

  function render() {
    const due = dueOf();
    const all = students.length;
    const paidCnt = students.filter((s) => statusOf(s) === "paid").length;
    const collected = students.reduce((a, s) => a + paidOf(s), 0);
    const expected = all * due;

    $("tcTotal").textContent = all;
    $("tcPaid").textContent = paidCnt;
    $("tcUnpaid").textContent = all - paidCnt;
    $("tcMoney").textContent = money(collected);
    $("tcPct").textContent = `trên ${money(expected)} (${expected ? Math.round(collected / expected * 100) : 0}%)`;
    $("tcCntAll").textContent = all;
    $("tcCntPaid").textContent = paidCnt;
    $("tcCntUnpaid").textContent = all - paidCnt;

    const list = filtered();
    $("tcNone").hidden = list.length > 0;
    $("tcBody").innerHTML = list.map((s, i) => {
      const st = statusOf(s), paid = paidOf(s), owe = Math.max(due - paid, 0);
      const badge = st === "paid" ? `<span class="tc-badge paid">Đã nộp đủ</span>`
        : st === "part" ? `<span class="tc-badge part">Nộp thiếu</span>` : `<span class="tc-badge none">Chưa nộp</span>`;
      const act = st === "paid"
        ? `<button class="tc-btn tc-row-btn undo" data-act="undo" data-id="${esc(s.id)}">Hoàn tác</button>`
        : `<button class="tc-btn tc-row-btn" data-act="pay" data-id="${esc(s.id)}">Ghi nhận nộp đủ</button>`;
      return `<tr>
        <td class="num">${i + 1}</td><td class="num">${esc(s.id)}</td><td>${esc(s.name)}</td><td>${esc(s.cls)}</td>
        <td class="r">${money(due)}</td><td class="r">${money(paid)}</td>
        <td class="r ${owe ? "owe" : ""}">${money(owe)}</td><td>${badge}</td><td>${act}</td></tr>`;
    }).join("");
  }

  /* ---------- Xuất CSV ---------- */
  function exportCsv() {
    const due = dueOf();
    const label = { paid: "Đã nộp đủ", part: "Nộp thiếu", none: "Chưa nộp" };
    const head = ["STT", "MSV", "Họ và tên", "Lớp", "Học phí phải nộp", "Đã nộp", "Còn thiếu", "Trạng thái"];
    const lines = filtered().map((s, i) => {
      const paid = paidOf(s);
      return [i + 1, s.id, s.name, s.cls, due, paid, Math.max(due - paid, 0), label[statusOf(s)]]
        .map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",");
    });
    const blob = new Blob(["\ufeff" + [head.join(","), ...lines].join("\r\n")], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `hoc-phi-${filter}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  /* ---------- Khởi tạo ---------- */
  function init() {
    root = document.getElementById("tc");
    if (!root || started) return;
    started = true;

    EXCEL_URL = root.dataset.excel || "script.xlsx";
    DEFAULT_PRICE = Number(root.dataset.price) || 450000;
    price = Number(load(LS_PRICE, DEFAULT_PRICE)) || DEFAULT_PRICE;
    $("tcPrice").value = price;

    $("tcPrice").addEventListener("change", (e) => {
      price = Math.max(0, Number(e.target.value) || 0);
      save(LS_PRICE, price);
      render();
    });

    root.querySelectorAll(".tc-tab").forEach((t) => t.addEventListener("click", () => {
      root.querySelectorAll(".tc-tab").forEach((x) => x.classList.remove("active"));
      t.classList.add("active");
      filter = t.dataset.filter;
      render();
    }));

    $("tcSearch").addEventListener("input", render);
    $("tcClass").addEventListener("change", render);
    $("tcExport").addEventListener("click", exportCsv);

    $("tcBody").addEventListener("click", (e) => {
      const b = e.target.closest("button[data-act]");
      if (!b) return;
      if (b.dataset.act === "pay") payments[b.dataset.id] = { t: "full" };
      else delete payments[b.dataset.id];
      save(LS_PAID, payments);
      render();
    });

    $("tcReset").addEventListener("click", () => {
      if (!confirm("Đặt lại tình trạng nộp tiền về dữ liệu mẫu ban đầu?")) return;
      payments = seedPayments();
      save(LS_PAID, payments);
      render();
    });

    $("tcPick").addEventListener("click", () => $("tcFile").click());
    $("tcFile").addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (f) f.arrayBuffer().then((b) => ingest(b, f.name));
    });

    autoLoad();
  }

  window.TaiChinh = { init };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();