const body = document.body;
const toggleBtn = document.getElementById("toggleNav");
const avatar = document.getElementById("avatarBtn");
const dropdown = document.getElementById("userDropdown");

/* Thu gọn / mở rộng thanh bên (nhớ lựa chọn lần trước) */
try {
  if (localStorage.getItem("navCollapsed") === "1") body.classList.add("nav-collapsed");
} catch (e) { /* bỏ qua nếu trình duyệt chặn localStorage */ }

toggleBtn.addEventListener("click", () => {
  const collapsed = body.classList.toggle("nav-collapsed");
  dropdown.classList.remove("show");
  try { localStorage.setItem("navCollapsed", collapsed ? "1" : "0"); } catch (e) {}
});

/* Menu người dùng */
avatar.addEventListener("click", () => dropdown.classList.toggle("show"));
avatar.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); dropdown.classList.toggle("show"); }
});
document.addEventListener("click", (e) => {
  if (!e.target.closest(".user-profile")) dropdown.classList.remove("show");
});

/* Mục đang chọn (menu chính + menu con) */
const mainLinks = document.querySelectorAll(".menu > li:not(.has-sub) > .nav-link");
const subLinks = document.querySelectorAll(".sub-link");
const subItem = document.getElementById("subStudy");
const subToggle = subItem.querySelector(".sub-toggle");

function clearActive() {
  mainLinks.forEach((l) => l.classList.remove("active"));
  subLinks.forEach((l) => l.classList.remove("active"));
  subItem.classList.remove("has-active");
}

mainLinks.forEach((link) => {
  link.addEventListener("click", (e) => {
    if (link.getAttribute("href") !== "#") return;   // link thật thì cho chuyển trang
    e.preventDefault();
    clearActive();
    link.classList.add("active");
  });
});

/* Mở / đóng nhóm "Quản lý học tập" */
subToggle.addEventListener("click", (e) => {
  e.preventDefault();
  if (body.classList.contains("nav-collapsed")) return;   // thu gọn: menu con hiện khi rê chuột
  const open = subItem.classList.toggle("open");
  subToggle.setAttribute("aria-expanded", open);
});

/* Chọn mục con: Tra cứu điểm / Tra cứu lịch */
subLinks.forEach((link) => {
  link.addEventListener("click", (e) => {
    if (link.getAttribute("href") === "#") e.preventDefault();
    clearActive();
    link.classList.add("active");
    subItem.classList.add("has-active", "open");
    subToggle.setAttribute("aria-expanded", "true");
    /* Gắn với trang tổng: nghe sự kiện này để chuyển nội dung theo data-target */
    document.dispatchEvent(new CustomEvent("nav:select", { detail: { target: link.dataset.target } }));
  });
});