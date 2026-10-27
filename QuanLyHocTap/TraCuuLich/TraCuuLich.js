const $=id=>document.getElementById(id);
const NO_DETAIL = [
  ["IT1.109.3-1-1-26(N04.TH1)","Cấu trúc dữ liệu và giải thuật-1-1-26(N04.TH1)"],
  ["IT1.108.3-1-1-26(N02.TH1)","Lập trình hướng đối tượng-1-1-26(N02.TH1)"],
  ["IT1.106.3-1-1-26(N04.TH1)","Thiết kế Web-1-1-26(N04.TH1)"],
  ["IT1.107.3-1-1-26(N04.TH1)","Kiến trúc và tổ chức máy tính-1-1-26(N04.TH1)"] ];
const EVENTS = [
  { d:0, s:"15:35", e:"18:00", tiet:"10-12", name:"Kiến trúc và tổ chức máy tính", cls:"Kiến trúc và tổ chức máy tính-1-1-26(N04.BT1)", room:"204-A8", c:"#d49d3a" },
  { d:1, s:"07:00", e:"09:25", tiet:"1-3",   name:"Thiết kế Web", cls:"Thiết kế Web-1-1-26(N04.BT1)", room:"104-A5", c:"#d49d3a" },
  { d:1, s:"13:00", e:"15:25", tiet:"7-9",   name:"Giáo dục thể chất F3", cls:"Giáo dục thể chất F3-1-1-26(N17).TH", room:"Ngoài trời 32", c:"#1f2f6b" },
  { d:2, s:"09:35", e:"12:00", tiet:"4-6",   name:"Toán rời rạc", cls:"Toán rời rạc-1-1-26(N04.BT1)", room:"204-A8", c:"#e84c00" },
  { d:4, s:"07:00", e:"09:25", tiet:"1-3",   name:"Cấu trúc dữ liệu và giải thuật", cls:"Cấu trúc dữ liệu và giải thuật-1-1-26(N04.BT1)", room:"204-A8", c:"#5878d8" },
  { d:5, s:"07:00", e:"09:25", tiet:"1-3",   name:"Lập trình hướng đối tượng", cls:"Lập trình hướng đối tượng-1-1-26(N02.BT1)", room:"201-A8", c:"#3a4f94" },
  { d:5, s:"13:00", e:"15:25", tiet:"7-9",   name:"Tư tưởng Hồ Chí Minh", cls:"Tư tưởng Hồ Chí Minh-1-1-26(N08.TL1)", room:"202-A8", c:"#1f2f6b" } ];
/* ================= TRA CỨU LỊCH ================= */
const H = 56, START = 7, DAYS = ["Thứ 2","Thứ 3","Thứ 4","Thứ 5","Thứ 6","Thứ 7","Chủ nhật"];
const mins = (t) => { const [h, m] = t.split(":").map(Number); return h * 60 + m; };
const fmt = (d) => d.toLocaleDateString("vi-VN");
const mondayOf = (d) => { const x = new Date(d); x.setHours(0,0,0,0); x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); return x; };
let weekStart = mondayOf(new Date());

function renderCal() {
  const today = new Date().toDateString();
  let html = `<div class="cal-h">🕒 Giờ VN</div>`;
  const dates = DAYS.map((_, i) => { const d = new Date(weekStart); d.setDate(d.getDate() + i); return d; });
  dates.forEach((d, i) => html += `<div class="cal-h ${d.toDateString() === today ? "today" : ""}"><b>${d.getDate()}</b>${DAYS[i]}</div>`);
  html += `<div class="hours">${Array.from({length: 13}, (_, i) => `<div>${START + i}:00</div>`).join("")}</div>`;
  for (let i = 0; i < 7; i++) {
    html += `<div class="col">` + EVENTS.filter(e => e.d === i).map(e => {
      const top = (mins(e.s) - START * 60) / 60 * H, h = (mins(e.e) - mins(e.s)) / 60 * H;
      return `<div class="ev" style="top:${top}px;height:${h}px"><header style="background:${e.c}"><b>${e.name}</b><small>${e.s} - ${e.e} (Tiết ${e.tiet})</small></header><p>${e.cls}<br>${e.room}</p></div>`;
    }).join("") + `</div>`;
  }
  $("cal").innerHTML = html;
  $("wkRange").textContent = `${fmt(dates[0])} - ${fmt(dates[6])}`;
}
$("prevWk").onclick = () => { weekStart.setDate(weekStart.getDate() - 7); renderCal(); };
$("nextWk").onclick = () => { weekStart.setDate(weekStart.getDate() + 7); renderCal(); };
$("todayWk").onclick = () => { weekStart = mondayOf(new Date()); renderCal(); };
$("noDetail").innerHTML = NO_DETAIL.map(r => `<tr><td class="code">${r[0]}</td><td>${r[1]}</td><td>Thực hành</td><td>10/08/2026</td><td>22/11/2026</td><td></td></tr>`).join("");