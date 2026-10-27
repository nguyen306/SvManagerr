const EXCEL_URL = "./script.xlsx";

let allStudents = [];
let currentFilter = "all";
let currentSearch = "";
let selectedMsv = null;

const COURSES = [
  { name: "Triết Học Mác-lênin", credits: 3, year: 1 },
  { name: "Tin đại cương", credits: 3, year: 1 },
  { name: "Giải tích", credits: 2, year: 1 },
  { name: "Đại số tuyến tính", credits: 3, year: 1 },
  { name: "Thiết kế web", credits: 3, year: 2 },
  { name: "Lập trình hướng đối tượng", credits: 3, year: 2 },
  { name: "Cấu trúc dữ liệu và giải thuật", credits: 3, year: 2 },
  { name: "Tư tưởng Hồ Chí Minh", credits: 2, year: 2 },
  { name: "Lập trình Web", credits: 3, year: 3 },
  { name: "Lập trình trực quan", credits: 3, year: 3 },
  { name: "Thuật toán và ứng dụng", credits: 3, year: 3 },
  { name: "Mạng máy tính", credits: 3, year: 3 },
  { name: "Bảo vệ đồ án", credits: 10, year: 4 }
];

function init() {
  fetch(EXCEL_URL)
    .then(r => r.arrayBuffer())
    .then(buf => {
      const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });

      allStudents = [];
      for (let i = 2; i < rows.length; i++) {
        const r = rows[i];
        if (!r || !r[1]) continue;

        const name = String(r[1]).trim();
        const msv = String(r[2] || "").trim();
        const cls = String(r[3] || "").trim();
        const email = String(r[4] || "").trim();
        const grades = r.slice(6, 19).map(v => parseFloat(v));

        let totalS = 0, totalC = 0;
        grades.forEach((g, idx) => {
          if (!isNaN(g)) {
            totalS += g * COURSES[idx].credits;
            totalC += COURSES[idx].credits;
          }
        });
        const gpa = totalC > 0 ? (totalS / totalC) : 0;

        allStudents.push({ name, msv, cls, email, grades, gpa });
      }

      renderStudentList();
    })
    .catch(err => {
      document.getElementById("studentList").innerHTML = `<div style="color:red;padding:10px">Lỗi đọc file script.xlsx</div>`;
    });

  // Sự kiện lọc GPA
  document.querySelectorAll(".btn-filter").forEach(btn => {
    btn.addEventListener("click", (e) => {
      document.querySelectorAll(".btn-filter").forEach(b => b.classList.remove("active"));
      e.target.classList.add("active");
      currentFilter = e.target.dataset.gpa;
      renderStudentList();
    });
  });

  // Sự kiện tìm kiếm
  document.getElementById("searchInput").addEventListener("input", (e) => {
    currentSearch = e.target.value.toLowerCase().trim();
    renderStudentList();
  });
}

function renderStudentList() {
  const container = document.getElementById("studentList");

  const filtered = allStudents.filter(s => {
    // Lọc theo GPA
    let passGpa = true;
    if (currentFilter === "3.6") passGpa = s.gpa > 3.6;
    else if (currentFilter === "3.2") passGpa = s.gpa > 3.2;
    else if (currentFilter === "2.5") passGpa = s.gpa > 2.5;

    // Lọc theo từ khóa tìm kiếm
    let passSearch = true;
    if (currentSearch) {
      passSearch = s.name.toLowerCase().includes(currentSearch) || s.msv.toLowerCase().includes(currentSearch);
    }

    return passGpa && passSearch;
  });

  if (filtered.length === 0) {
    container.innerHTML = `<div class="placeholder-text">Không tìm thấy sinh viên phù hợp</div>`;
    return;
  }

  container.innerHTML = filtered.map(s => `
    <div class="student-item ${selectedMsv === s.msv ? 'active' : ''}" onclick="selectStudent('${s.msv}')">
      <div class="s-info">
        <strong>${s.name}</strong>
        <small>MSV: ${s.msv} | ${s.cls}</small>
      </div>
      <span class="s-gpa">${s.gpa.toFixed(2)}</span>
    </div>
  `).join("");
}

function selectStudent(msv) {
  selectedMsv = msv;
  renderStudentList();

  const student = allStudents.find(s => s.msv === msv);
  if (!student) return;

  const detailBox = document.getElementById("detailContent");

  let tablesHtml = "";
  [1, 2, 3, 4].forEach(yr => {
    const yrCourses = COURSES.map((c, idx) => ({ ...c, idx })).filter(c => c.year === yr);
    if (yrCourses.length === 0) return;

    const rowsHtml = yrCourses.map(c => {
      const score = student.grades[c.idx];
      return `
        <tr>
          <td>${c.name}</td>
          <td style="text-align:center">${c.credits}</td>
          <td style="text-align:center"><strong>${isNaN(score) ? "-" : score.toFixed(1)}</strong></td>
        </tr>
      `;
    }).join("");

    tablesHtml += `
      <div class="year-block">
        <div class="year-title">NĂM HỌC ${yr}</div>
        <table class="grade-tbl">
          <thead>
            <tr><th>Tên môn học</th><th style="width:80px;text-align:center">Tín chỉ</th><th style="width:100px;text-align:center">Điểm hệ 4</th></tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
  });

  detailBox.innerHTML = `
    <div class="student-detail-header">
      <h2>${student.name}</h2>
      <p>MSV: ${student.msv} | Lớp: ${student.cls} | Email: ${student.email} | <strong>GPA Tích lũy: ${student.gpa.toFixed(2)}</strong></p>
    </div>
    ${tablesHtml}
  `;
}

document.addEventListener("DOMContentLoaded", init);