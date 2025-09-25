let students = [];

// Đọc file Excel
document.getElementById("fileInput").addEventListener("change", function(e) {
  let file = e.target.files[0];
  let reader = new FileReader();

  reader.onload = function(e) {
    let data = new Uint8Array(e.target.result);
    let workbook = XLSX.read(data, { type: "array" });

    // Lấy sheet đầu tiên
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    let rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    students = [];

    // Giả sử Excel có dạng: STT | Tên | Điểm
    for (let i = 1; i < rows.length; i++) {
      let row = rows[i];
      if (row[1] && row[2]) {
        students.push({ name: row[1], score: Number(row[2]) });
      }
    }

    renderTable(students);
  };

  reader.readAsArrayBuffer(file);
});

// Hiển thị bảng
function renderTable(data) {
  let sorted = [...data].sort((a, b) => b.score - a.score);
  sorted = sorted.slice(0, 30); // chỉ lấy top 30

  let result = "<h3>Top 30 học sinh</h3>";
  result += "<table><tr><th>STT</th><th>Tên</th><th>Điểm</th></tr>";
  sorted.forEach((s, i) => {
    result += `<tr><td>${i+1}</td><td>${s.name}</td><td>${s.score}</td></tr>`;
  });
  result += "</table>";
  document.getElementById("result").innerHTML = result;
}

// Lọc theo tên
function filterTable() {
  const keyword = document.getElementById("search").value.toLowerCase();
  const filtered = students.filter(s => s.name.toLowerCase().includes(keyword));
  renderTable(filtered);
}

// Xuất file Excel
function exportFile() {
  if (students.length === 0) {
    alert("Chưa có dữ liệu để xuất!");
    return;
  }

  let sorted = [...students].sort((a, b) => b.score - a.score);
  sorted = sorted.slice(0, 30);

  let data = [["STT", "Tên", "Điểm"]];
  sorted.forEach((s, i) => {
    data.push([i + 1, s.name, s.score]);
  });

  let worksheet = XLSX.utils.aoa_to_sheet(data);
  let workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Top30");

  XLSX.writeFile(workbook, "Top30HocSinh.xlsx");
}
