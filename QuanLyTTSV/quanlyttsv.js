const EXCEL_URL = "./script.xlsx";

let globalStudents = [];
let currentEditingStudent = null;
let isInitialLoad = true;

// Tính GPA theo trọng số tín chỉ tương ứng với cấu trúc file script.xlsx
function calculateWeightedGPA(rowGrades) {
  const credits = [3, 3, 2, 3, 3, 3, 3, 2, 3, 3, 3, 3, 10];
  let totalScore = 0;
  let totalCredits = 0;
  
  rowGrades.forEach((g, idx) => {
    if (!isNaN(g) && idx < credits.length) {
      totalScore += g * credits[idx];
      totalCredits += credits[idx];
    }
  });
  return totalCredits > 0 ? (totalScore / totalCredits).toFixed(2) : "0.00";
}

// Nạp dữ liệu từ file Excel
function loadData() {
  fetch(EXCEL_URL)
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const students = [];
      for (let i = 2; i < rows.length; i++) {
        const r = rows[i];
        if (!r || !r[1]) continue;

        const name = String(r[1]).trim();
        const msv = String(r[2] || "").trim();
        const cls = String(r[3] || "").trim();
        const email = String(r[4] || "").trim();
        const phone = String(r[5] || "").trim();
        
        const grades = r.slice(6, 19).map(v => parseFloat(v));
        const gpa = calculateWeightedGPA(grades);

        students.push({ stt: students.length + 1, name, msv, cls, email, phone, gpa, bankAcc: "", bankName: "" });
      }

      globalStudents = students;
      displayStudentList(globalStudents);
    })
    .catch(err => {
      console.error("Lỗi nạp file Excel:", err);
      document.getElementById("studentTableBody").innerHTML = 
        `<tr><td colspan="6" class="text-center" style="color:red">Không thể tải file script.xlsx. Vui lòng kiểm tra lại file.</td></tr>`;
    });
}

// Hiển thị danh sách sinh viên lên bảng
function displayStudentList(list) {
  const tbody = document.getElementById("studentTableBody");
  document.getElementById("studentCount").textContent = `Tổng số: ${list.length} sinh viên`;

  if (list.length > 0 && isInitialLoad) {
    setProfile(list[0]);
    isInitialLoad = false;
  }

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center">Không tìm thấy sinh viên nào phù hợp.</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map((s, index) => `
    <tr onclick="selectStudent('${s.msv}')">
      <td class="text-center">${index + 1}</td>
      <td><span class="msv-text">${s.msv}</span></td>
      <td><strong>${s.name}</strong></td>
      <td>${s.cls}</td>
      <td class="text-center"><span class="gpa-badge">${s.gpa}</span></td>
      <td>${s.email}</td>
    </tr>
  `).join("");

  window.studentData = list;
}

// Chọn sinh viên và đẩy thông tin lên phần Hồ sơ
function selectStudent(msv) {
  const s = globalStudents.find(x => x.msv === msv);
  if (s) setProfile(s);
}

// Đổ dữ liệu hồ sơ cá nhân lên Form Admin
function setProfile(student) {
  currentEditingStudent = student;

  // Sidebar bên trái
  const sidebarList = document.getElementById("sidebarInfoList");
  sidebarList.innerHTML = `
    <div class="info-item"><span class="info-label">HỌ VÀ TÊN</span><strong class="info-val">${student.name}</strong></div>
    <div class="info-item"><span class="info-label">MÃ SV</span><span class="info-val">${student.msv}</span></div>
    <div class="info-item"><span class="info-label">LỚP</span><span class="info-val">${student.cls}</span></div>
    <div class="info-item"><span class="info-label">GPA</span><span class="info-val" style="color: #16a34a; font-weight:700;">${student.gpa}</span></div>
  `;

  // Bảng chi tiết bên phải
  const fields = [
    { key: "name", label: "Họ và tên" },
    { key: "msv", label: "Mã sinh viên" },
    { key: "cls", label: "Lớp quản lý" },
    { key: "email", label: "Email" },
    { key: "phone", label: "Số điện thoại" },
    { key: "bankAcc", label: "Số tài khoản NH" },
    { key: "bankName", label: "Tên ngân hàng" }
  ];

  const tbody = document.getElementById("adminProfileTableBody");
  tbody.innerHTML = fields.map(f => {
    const value = student[f.key] || "";
    return `
      <tr>
        <td class="form-label">${f.label}</td>
        <td>
          <input type="text" id="field_${f.key}" class="editable-field" value="${value}" disabled />
        </td>
        <td class="text-center">
          <button class="btn-edit-item" id="btn_edit_${f.key}" onclick="toggleEditRow('${f.key}')">Sửa</button>
        </td>
      </tr>
    `;
  }).join("");
}

// Bật / tắt chế độ chỉnh sửa từng ô input
function toggleEditRow(key) {
  const input = document.getElementById(`field_${key}`);
  const btn = document.getElementById(`btn_edit_${key}`);

  if (input.disabled) {
    input.disabled = false;
    input.focus();
    btn.textContent = "Khóa";
    btn.style.backgroundColor = "#fee2e2";
    btn.style.color = "#dc2626";
    btn.style.borderColor = "#fca5a5";
  } else {
    input.disabled = true;
    btn.textContent = "Sửa";
    btn.style.backgroundColor = "#f1f5f9";
    btn.style.color = "#2563eb";
    btn.style.borderColor = "#cbd5e1";
  }
}

// Lưu tất cả thay đổi trên Hồ sơ cá nhân
function saveAllAdminChanges() {
  if (!currentEditingStudent) return;

  currentEditingStudent.name = document.getElementById("field_name").value.trim();
  currentEditingStudent.msv = document.getElementById("field_msv").value.trim();
  currentEditingStudent.cls = document.getElementById("field_cls").value.trim();
  currentEditingStudent.email = document.getElementById("field_email").value.trim();
  currentEditingStudent.phone = document.getElementById("field_phone").value.trim();
  currentEditingStudent.bankAcc = document.getElementById("field_bankAcc").value.trim();
  currentEditingStudent.bankName = document.getElementById("field_bankName").value.trim();

  alert(`Đã cập nhật thông tin thành công cho sinh viên ${currentEditingStudent.name}!`);

  setProfile(currentEditingStudent);
  displayStudentList(globalStudents);
}

// Tìm kiếm sinh viên theo Tên hoặc Mã SV
function searchStudentByAdmin() {
  const query = document.getElementById("adminSearchInput").value.trim().toLowerCase();
  if (!query) {
    alert("Vui lòng nhập tên hoặc Mã sinh viên để tìm kiếm!");
    return;
  }

  const found = globalStudents.find(s => 
    s.msv.toLowerCase().includes(query) || 
    s.name.toLowerCase().includes(query)
  );

  if (found) {
    setProfile(found);
  } else {
    alert("Không tìm thấy sinh viên nào phù hợp!");
  }
}

// Lọc sinh viên theo GPA và sắp xếp giảm dần
function filterStudents(type, btnElement) {
  if (btnElement) {
    document.querySelectorAll('.btn-filter').forEach(b => b.classList.remove('active'));
    btnElement.classList.add('active');
  }

  let filtered = [...globalStudents];

  switch (type) {
    case 'gt36':
      filtered = globalStudents.filter(s => parseFloat(s.gpa) > 3.6);
      break;
    case '32_36':
      filtered = globalStudents.filter(s => parseFloat(s.gpa) > 3.2 && parseFloat(s.gpa) <= 3.6);
      break;
    case '25_32':
      filtered = globalStudents.filter(s => parseFloat(s.gpa) > 2.5 && parseFloat(s.gpa) <= 3.2);
      break;
    case 'lt25':
      filtered = globalStudents.filter(s => parseFloat(s.gpa) <= 2.5);
      break;
    case 'all':
    default:
      filtered = [...globalStudents];
      break;
  }

  // Sắp xếp điểm GPA giảm dần (từ cao xuống thấp)
  filtered.sort((a, b) => parseFloat(b.gpa) - parseFloat(a.gpa));

  displayStudentList(filtered);
}

// Lắng nghe phím Enter trên ô tìm kiếm
document.addEventListener("DOMContentLoaded", () => {
  loadData();

  const searchInput = document.getElementById("adminSearchInput");
  if (searchInput) {
    searchInput.addEventListener("keypress", (e) => {
      if (e.key === "Enter") searchStudentByAdmin();
    });
  }
});