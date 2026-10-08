/* ======================================================================
   Trang chủ sinh viên - dữ liệu được đọc từ file Excel người dùng chọn
   (cấu trúc: dòng 1 = tiêu đề + "Năm X", dòng 2 = tên môn(số tín chỉ),
   từ dòng 3 = từng sinh viên; cột cuối là GPA)
   ====================================================================== */

/* Xếp loại theo thang điểm 4 */
const LEVELS = [
  { label: "Giỏi",       min: 3.2, cls: "badge-green",  color: "#12c46a" },
  { label: "Khá",        min: 2.5, cls: "badge-blue",   color: "#2b7de0" },
  { label: "Trung bình", min: 1.0, cls: "badge-yellow", color: "#f5a524" },
  { label: "Chưa đạt",   min: 0,   cls: "badge-orange", color: "#f26b3a" }
];
const RANK_PREVIEW = 5;

/* File Excel có sẵn, đặt cùng thư mục với index.html */
const EXCEL_URL = "script.xlsx";

/* ======================== Trạng thái ======================== */
let COURSES = [];            // [{name, credits, year, col}]
let TOTAL_CREDITS = 0;
let DATA = [];               // danh sách sinh viên
let rankingByClass = {};
let current = null;
let showAllRank = false;

/* ======================== Tiện ích ======================== */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
const fmt = (n) => n.toFixed(2);
const pad = (n) => String(n).padStart(2, "0");
const norm = (v) => String(v == null ? "" : v).trim().toLowerCase();
const levelOf = (score) => LEVELS.find((l) => score >= l.min) || LEVELS[LEVELS.length - 1];

/* GPA = tổng(điểm × tín chỉ) / tổng tín chỉ (khớp cột GPA trong Excel) */
function weightedAvg(grades, courses) {
  let sum = 0, cr = 0;
  grades.forEach((g, i) => { sum += g * courses[i].credits; cr += courses[i].credits; });
  return cr ? sum / cr : 0;
}

/* ======================== Đọc file Excel ======================== */
function parseWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (rows.length < 3) throw new Error("File quá ít dòng, không có dữ liệu sinh viên.");

  const h0 = rows[0] || [];
  const h1 = rows[1] || [];

  // tìm cột theo tiêu đề (có dự phòng theo vị trí mặc định)
  const find = (kw, fallback) => {
    const i = h0.findIndex((c) => norm(c).includes(kw));
    return i >= 0 ? i : fallback;
  };
  const iName = find("họ và tên", 1);
  const iId = find("msv", 2);
  const iCls = find("lớp", 3);
  const iMail = find("email", 4);
  const iPhone = find("điện thoại", 5);

  const gpaCol = h0.findIndex((c) => norm(c) === "gpa");
  const start = Math.max(iName, iId, iCls, iMail, iPhone) + 1;
  const end = gpaCol >= 0 ? gpaCol : h1.length;

  // các cột môn học: tên(số tín chỉ), năm lấy từ ô "Năm X" ở dòng 1 (kế thừa sang các cột trống)
  const courses = [];
  let year = 1;
  for (let c = start; c < end; c++) {
    const yr = String(h0[c] == null ? "" : h0[c]).match(/\d+/);
    if (yr) year = Number(yr[0]);
    const raw = String(h1[c] == null ? "" : h1[c]).trim();
    if (!raw) continue;
    const cr = raw.match(/\((\d+(?:[.,]\d+)?)\)\s*$/);
    courses.push({
      name: raw.replace(/\s*\(\d+(?:[.,]\d+)?\)\s*$/, ""),
      credits: cr ? Number(cr[1].replace(",", ".")) : 1,
      year,
      col: c
    });
  }
  if (!courses.length) throw new Error("Không tìm thấy cột điểm môn học (dòng 2).");

  const students = [];
  rows.slice(2).forEach((r) => {
    if (!r || !r[iName]) return;
    const grades = courses.map((c) => parseFloat(String(r[c.col]).replace(",", ".")));
    if (grades.some((g) => isNaN(g))) return;           // bỏ dòng thiếu điểm
    const phone = r[iPhone] == null ? "" : String(r[iPhone]).trim();
    students.push({
      name: String(r[iName]).trim(),
      id: typeof r[iId] === "number" ? String(Math.round(r[iId])) : String(r[iId] == null ? "" : r[iId]).trim(),
      cls: String(r[iCls] == null ? "" : r[iCls]).trim(),
      email: String(r[iMail] == null ? "" : r[iMail]).trim(),
      phone: /^\d{9}$/.test(phone) ? "0" + phone : phone,   // Excel hay làm mất số 0 đầu
      grades
    });
  });
  if (!students.length) throw new Error("Không đọc được sinh viên nào (kiểm tra lại cấu trúc file).");

  return { courses, students };
}

function loadData({ courses, students }) {
  COURSES = courses;
  TOTAL_CREDITS = courses.reduce((s, c) => s + c.credits, 0);
  DATA = students;
  DATA.forEach((s) => (s.gpa = weightedAvg(s.grades, COURSES)));

  // xếp hạng theo GPA giảm dần trong từng lớp
  rankingByClass = {};
  DATA.forEach((s) => (rankingByClass[s.cls] = rankingByClass[s.cls] || []).push(s));
  Object.values(rankingByClass).forEach((list) => {
    list.sort((a, b) => b.gpa - a.gpa || a.name.localeCompare(b.name, "vi"));
    list.forEach((s, i) => (s.rank = i + 1));
  });

  // gợi ý cho ô tìm kiếm
  $("studentList").innerHTML = DATA.map((s) => `<option value="${esc(s.id)} - ${esc(s.name)}">${esc(s.cls)}</option>`).join("");
  $("studentSearch").disabled = false;

  current = DATA[0];
  showAllRank = false;
  $("page").classList.remove("is-empty");
  renderAll();
}

function setMsg(text, type) {
  const msg = $("uploadMsg");
  msg.textContent = text;
  msg.className = "upload-msg" + (type ? " " + type : "");
}

/* Xử lý dữ liệu nhị phân của file Excel */
function processBuffer(buffer, label) {
  if (typeof XLSX === "undefined") {
    setMsg("Không tải được thư viện đọc Excel (SheetJS). Hãy kiểm tra kết nối mạng rồi tải lại trang.", "error");
    return;
  }
  try {
    const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
    const parsed = parseWorkbook(wb);
    loadData(parsed);
    setMsg(`Đã nạp "${label}": ${parsed.students.length} sinh viên, ${parsed.courses.length} môn.`, "ok");
  } catch (err) {
    setMsg("Lỗi: " + err.message, "error");
  }
}

/* Tự động nạp file Excel có sẵn khi mở trang */
function autoLoad() {
  setMsg("Đang đọc " + EXCEL_URL + "...");
  fetch(encodeURI(EXCEL_URL))
    .then((r) => {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.arrayBuffer();
    })
    .then((buf) => processBuffer(buf, EXCEL_URL))
    .catch(() => {
      setMsg(`Không đọc được "${EXCEL_URL}". Hãy chắc chắn file nằm cùng thư mục với index.html và trang đang chạy bằng Live Server.`, "error");
      $("emptyState").querySelector("h2").textContent = "Không tải được dữ liệu";
    });
}

/* ======================== Render ======================== */
function renderHeader(s) {
  $("studentName").textContent = s.name;
  $("studentId").textContent = s.id;
  document.title = "Trang chủ - " + s.name;
}

function renderStats(s) {
  $("statGpa").textContent = fmt(s.gpa) + " / 4.0";
  $("statCredits").textContent = TOTAL_CREDITS;

  let best = 0;
  s.grades.forEach((g, i) => { if (g > s.grades[best]) best = i; });
  $("statBest").textContent = fmt(s.grades[best]);
  $("statBestName").textContent = COURSES[best].name;

  $("statRank").textContent = "#" + s.rank + " / " + rankingByClass[s.cls].length;
  $("statClass").textContent = s.cls;
}

function renderGrades(s) {
  $("gradeBody").innerHTML = COURSES.map((c, i) => {
    const g = s.grades[i];
    const lv = levelOf(g);
    return `<tr>
      <td class="num">${i + 1}</td>
      <td>${esc(c.name)}</td>
      <td>Năm ${c.year}</td>
      <td class="num">${c.credits}</td>
      <td class="num score">${fmt(g)}</td>
      <td><span class="badge ${lv.cls}">${lv.label}</span></td>
    </tr>`;
  }).join("");
}

function renderInfo(s) {
  const rows = [
    ["Họ và tên", s.name],
    ["Mã sinh viên", s.id],
    ["Lớp", s.cls],
    ["Email", s.email],
    ["Số điện thoại", s.phone]
  ];
  $("infoList").innerHTML = rows.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join("");

  const passed = s.grades.filter((g) => g >= 1.0).length;
  const sorted = s.grades.slice().sort((a, b) => a - b);
  $("infoSummary").innerHTML =
    `<span class="chip">Môn đạt: <b>${passed}/${COURSES.length}</b></span>` +
    `<span class="chip">Điểm thấp nhất: <b>${fmt(sorted[0])}</b></span>` +
    `<span class="chip">Điểm cao nhất: <b>${fmt(sorted[sorted.length - 1])}</b></span>`;
}

function renderRanking(s) {
  const list = rankingByClass[s.cls];
  const shown = showAllRank ? list : list.slice(0, RANK_PREVIEW);

  $("rankBody").innerHTML = shown.map((st) => `
    <tr class="${st === s ? "is-me" : ""}">
      <td class="num">${st.rank}</td>
      <td class="num">${esc(st.id)}</td>
      <td>${esc(st.name)}</td>
      <td class="num score">${fmt(st.gpa)}</td>
      <td><button class="btn" type="button" data-id="${esc(st.id)}">Xem chi tiết</button></td>
    </tr>`).join("");

  $("toggleRank").textContent = showAllRank ? "Thu gọn" : "Xem tất cả (" + list.length + ")";
  $("toggleRank").style.display = list.length > RANK_PREVIEW ? "" : "none";
  $("rankWrap").classList.toggle("collapsed", !showAllRank);
}

/* ---- Biểu đồ tròn: số môn theo xếp loại ---- */
function renderDonut(s) {
  const counts = LEVELS.map((lv, i) => {
    const upper = i === 0 ? Infinity : LEVELS[i - 1].min;
    return s.grades.filter((g) => g >= lv.min && g < upper).length;
  });
  const total = counts.reduce((a, b) => a + b, 0);

  const R = 48, C = 2 * Math.PI * R, W = 20;
  let offset = 0;
  const arcs = counts.map((n, i) => {
    if (!n) return "";
    const len = (n / total) * C;
    const el = `<circle cx="70" cy="70" r="${R}" fill="none" stroke="${LEVELS[i].color}" stroke-width="${W}"
      stroke-dasharray="${len} ${C - len}" stroke-dashoffset="${-offset}" transform="rotate(-90 70 70)">
      <title>${LEVELS[i].label}: ${n} môn</title></circle>`;
    offset += len;
    return el;
  }).join("");

  $("donut").innerHTML = `<svg viewBox="0 0 140 140" role="img" aria-label="Phân bố xếp loại">
    <circle cx="70" cy="70" r="${R}" fill="none" stroke="#eef2f7" stroke-width="${W}"/>
    ${arcs}
    <text x="70" y="70" text-anchor="middle" class="donut-center">${total}</text>
    <text x="70" y="85" text-anchor="middle" class="donut-center-sub">môn học</text>
  </svg>`;

  $("donutLegend").innerHTML = LEVELS.map((lv, i) =>
    `<li><i style="background:${lv.color}"></i><span>${lv.label}<br><b>${counts[i]} môn</b></span></li>`).join("");
}

/* ---- Biểu đồ cột: GPA theo năm học ---- */
function renderBars(s) {
  const yearList = [...new Set(COURSES.map((c) => c.year))].sort((a, b) => a - b);
  const years = yearList.map((y) => {
    const idx = COURSES.map((c, i) => (c.year === y ? i : -1)).filter((i) => i >= 0);
    const avg = weightedAvg(idx.map((i) => s.grades[i]), idx.map((i) => COURSES[i]));
    return { label: "Năm " + y, avg };
  });

  const W = 300, H = 190, L = 30, B = 24, T = 14, R = 8;
  const plotH = H - B - T, plotW = W - L - R;
  const MAX = 4;
  const slot = plotW / years.length, bw = Math.min(slot * 0.5, 60);

  let grid = "";
  for (let v = 0; v <= MAX; v++) {
    const y = T + plotH - (v / MAX) * plotH;
    grid += `<line x1="${L}" x2="${W - R}" y1="${y}" y2="${y}" stroke="#e5e7eb"/>
             <text x="${L - 5}" y="${y + 3}" text-anchor="end" class="axis-text">${v}</text>`;
  }

  const bars = years.map((d, i) => {
    const h = (d.avg / MAX) * plotH;
    const x = L + slot * i + (slot - bw) / 2;
    const y = T + plotH - h;
    return `<g>
      <rect class="bar" x="${x}" y="${y}" width="${bw}" height="${h}" rx="3"><title>${d.label}: ${fmt(d.avg)}</title></rect>
      <text class="bar-label" x="${x + bw / 2}" y="${y - 4}">${fmt(d.avg)}</text>
      <text class="axis-text" x="${x + bw / 2}" y="${H - 8}" text-anchor="middle">${d.label}</text>
    </g>`;
  }).join("");

  $("bars").innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="GPA theo năm học">${grid}${bars}</svg>`;
}

function renderAll() {
  renderHeader(current);
  renderStats(current);
  renderGrades(current);
  renderInfo(current);
  renderRanking(current);
  renderDonut(current);
  renderBars(current);
}

function selectStudent(id, scrollTop) {
  const s = DATA.find((x) => x.id === String(id));
  if (!s) return;
  current = s;
  showAllRank = false;
  $("studentSearch").value = "";
  renderAll();
  if (scrollTop) window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ======================== Đồng hồ ======================== */
function tickClock() {
  const d = new Date();
  $("clock").textContent =
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

/* ======================== Khởi tạo & sự kiện ======================== */
function init() {
  $("page").classList.add("is-empty");
  $("studentSearch").disabled = true;

  $("studentSearch").addEventListener("change", (e) => {
    const v = e.target.value.trim();
    const m = v.match(/^\d{6,}/);
    if (m) { selectStudent(m[0], false); return; }
    const kw = v.toLowerCase();
    const found = kw && DATA.find((s) => s.name.toLowerCase().includes(kw));
    if (found) selectStudent(found.id, false); else e.target.value = "";
  });

  $("toggleRank").addEventListener("click", () => { showAllRank = !showAllRank; renderRanking(current); });

  $("rankBody").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-id]");
    if (btn) selectStudent(btn.dataset.id, true);
  });

  tickClock();
  setInterval(tickClock, 1000);
  autoLoad();
}

document.addEventListener("DOMContentLoaded", init);