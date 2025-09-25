let students = [];

// Đọc file Excel
document.getElementById("fileInput").addEventListener("change", function(e) {
  let file = e.target.files[0];
  let reader = new FileReader();
  reader.onload = function(e) {
    let data = new Uint8Array(e.target.result);
    let workbook = XLSX.read(data, { type: "array" });
    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    students = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })
                .slice(1) // bỏ dòng tiêu đề
                .map((row, i) => ({
                  stt: i + 1,
                  name: row[0] || "",  // Cột 1: Họ và tên
                  score: row[1] || ""  // Cột 2: Điểm
                }));
    renderTable(students);
  };
  reader.readAsArrayBuffer(file);
});

// Hiển thị bảng
function renderTable(data) {
  let tbody = document.querySelector("#studentTable tbody");
  tbody.innerHTML = "";
  data.forEach(s => {
    let tr = `<tr>
                <td>${s.stt}</td>
                <td>${s.name}</td>
                <td>${s.score}</td>
              </tr>`;
    tbody.innerHTML += tr;
  });
}

// Tìm kiếm
document.getElementById("searchInput").addEventListener("input", function() {
  let keyword = this.value.toLowerCase();
  let filtered = students.filter(s => s.name.toLowerCase().includes(keyword));
  renderTable(filtered);
});

// Xuất file Top 30 (theo thứ tự trong Excel)
document.getElementById("exportBtn").addEventListener("click", function() {
  let top30 = students.slice(0, 30);
  let ws = XLSX.utils.json_to_sheet(top30);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Top30");
  XLSX.writeFile(wb, "top30.xlsx");
});
