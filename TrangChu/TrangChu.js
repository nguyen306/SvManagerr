/* ======================================================================
   Trang chủ sinh viên: banner chào mừng, lối tắt, tin nhà trường,
   tin đào tạo, hoạt động sinh viên.
   Tên sinh viên được lấy từ script.xlsx (theo MSV bên dưới).
   ====================================================================== */

const EXCEL_URL = "TrangChu/script.xlsx"; // đường dẫn file Excel (như bản cũ)
const MY_ID = "251210934"; // MSV của sinh viên đang đăng nhập
const FALLBACK_NAME = "ADMIN"; // dùng khi không đọc được Excel

/* ---- Dữ liệu tin tức: sửa/thêm tại đây ---- */
const SCHOOL_NEWS = [
  {
    title:
      "Thông báo thực hiện chế độ chính sách sinh viên học kỳ I năm học 2026 - 2027 (đợt bổ sung)",
    date: "16/10/2026",
    url: "#",
  },
  {
    title: "Cập nhật số CCCD trên Quản lý đào tạo",
    date: "14/10/2026",
    url: "#",
  },
  {
    title:
      "TKB Giáo dục quốc phòng an ninh (Khóa 67) cho sinh viên khoa Cơ khí- HK1- năm 2026-2027",
    date: "29/09/2026",
    url: "#",
  },
];
const TRAINING_NEWS = [
  {
    title:
      "Thông báo về việc triển khai rút bớt khối lượng học tập đã đăng ký trong đợt học chính khóa (đợt học 1) và học kì phụ (đợt học 2), học kì 1 năm học 2026-2027 đối với sinh viên các khóa",
    date: "09/10/2026",
    url: "#",
  },
  {
    title:
      "Lịch đăng ký học các lớp học phần ngoại ngữ tăng cường cho K67 học kỳ 1 năm học 2026 - 2027 đợt 11",
    date: "09/10/2026",
    url: "#",
  },
  {
    title:
      "Lịch đăng ký học các lớp học phần ngoại ngữ tăng cường cho K66 trở về trước học kỳ 1 năm học 2026 - 2027 đợt 8",
    date: "08/10/2026",
    url: "#",
  },
];
const ACTIVITY_NEWS = []; // để trống => hiện "Chưa có tin nào"

/* ======================== Tiện ích ======================== */
const $ = (id) => document.getElementById(id);
const esc = (s) =>
  String(s).replace(
    /[&<>"']/g,
    (m) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[
        m
      ])
  );
const pad = (n) => String(n).padStart(2, "0");

const ICON_PIN = `<svg class="pin" viewBox="0 0 24 24" width="18" height="18" fill="currentColor" aria-hidden="true"><path d="M16 3l5 5-2 1-3 3 .5 4.5-1.5 1.5-3.5-3.5L6 20l-1-1 4.5-5.5L6 10l1.5-1.5L12 9l3-3z"/></svg>`;
const ICON_CAL = `<svg viewBox="0 0 24 24" width="15" height="15" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M3 10h18M8 3v4M16 3v4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>`;
const ICON_INBOX = `<svg viewBox="0 0 24 24" width="32" height="32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 13l3-8h12l3 8v6H3z"/><path d="M3 13h5l1 3h6l1-3h5"/></svg>`;

/* ======================== Banner ======================== */
function renderGreeting() {
  const d = new Date();
  const h = d.getHours();
  const greet =
    h < 11
      ? "Chào buổi sáng"
      : h < 14
      ? "Chào buổi trưa"
      : h < 18
      ? "Chào buổi chiều"
      : "Chào buổi tối";
  $("greeting").textContent = greet + ",";

  const days = [
    "Chủ Nhật",
    "Thứ Hai",
    "Thứ Ba",
    "Thứ Tư",
    "Thứ Năm",
    "Thứ Sáu",
    "Thứ Bảy",
  ];
  $("today").textContent = `${days[d.getDay()]}, ngày ${pad(d.getDate())}/${pad(
    d.getMonth() + 1
  )}/${d.getFullYear()}`;
}

/* ======================== Tin tức ======================== */
function renderNews(id, items) {
  const ul = $(id);
  if (!items.length) {
    ul.innerHTML = `<li class="empty">${ICON_INBOX}<span>Chưa có tin nào</span></li>`;
    return;
  }
  ul.innerHTML = items
    .map(
      (n) => `
    <li>
      ${ICON_PIN}
      <div>
        <a href="${esc(n.url || "#")}">${esc(n.title)}</a>
        <div class="date">${ICON_CAL}<span>${esc(n.date)}</span></div>
      </div>
    </li>`
    )
    .join("");
}

/* ======================== Lấy tên từ Excel ======================== */
function loadName() {
  $("studentName").textContent = FALLBACK_NAME;
  if (typeof XLSX === "undefined") return;

  fetch(encodeURI(EXCEL_URL))
    .then((r) => {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.arrayBuffer();
    })
    .then((buf) => {
      const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        header: 1,
        defval: null,
      });
      const head = rows[0] || [];
      const iName = Math.max(
        0,
        head.findIndex((c) =>
          String(c || "")
            .toLowerCase()
            .includes("họ và tên")
        )
      );
      const iId = Math.max(
        0,
        head.findIndex((c) =>
          String(c || "")
            .toLowerCase()
            .includes("msv")
        )
      );
      const me = rows
        .slice(2)
        .find((r) => r && String(Math.round(Number(r[iId]))) === MY_ID);
      if (me && me[iName]) {
        const name = String(me[iName]).trim();
        $("studentName").textContent = name;
        document.title = "Trang chủ - " + name;
      }
    })
    .catch(() => {
      /* giữ tên mặc định */
    });
}

/* ======================== Khởi tạo ======================== */
document.addEventListener("DOMContentLoaded", () => {
  renderGreeting();
  renderNews("schoolNews", SCHOOL_NEWS);
  renderNews("trainingNews", TRAINING_NEWS);
  renderNews("activityNews", ACTIVITY_NEWS);
  loadName();
});
