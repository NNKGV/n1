let students = [];
let top30Data = [];

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

function renderTable(data) {
  let sorted = [...data].sort((a, b) => b.score - a.score);
  top30Data = sorted.slice(0, 30);

  let result = "<h3>Top 30 học sinh điểm cao nhất</h3>";
  result += "<table><tr><th>STT</th><th>Tên</th><th>Điểm</th></tr>";
  top30Data.forEach((s, i) => {
    result += `<tr><td>${i+1}</td><td>${s.name}</td><td>${s.score}</td></tr>`;
  });
  result += "</table>";
  document.getElementById("result").innerHTML = result;
}

function filterTable() {
  const keyword = document.getElementById("search").value.toLowerCase();
  const filtered = students.filter(s => s.name.toLowerCase().includes(keyword));
  renderTable(filtered);
}

// 👉 Hàm xuất file Excel Top 30
function exportExcel() {
  if (top30Data.length === 0) {
    alert("Chưa có dữ liệu để xuất!");
    return;
  }

  // Chuẩn bị dữ liệu với STT
  const exportData = top30Data.map((s, i) => ({
    STT: i + 1,
    Tên: s.name,
    Điểm: s.score
  }));

  // Tạo sheet và workbook
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Top30");

  // Xuất file Excel
  XLSX.writeFile(wb, "Top30HocSinh.xlsx");
}
