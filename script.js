function addRow(containerId, className, fieldCount) {
  const container = document.getElementById(containerId);

  // Create a new input row
  const row = document.createElement("div");
  row.className = "input-row";

  // Create inputs
  for (let i = 0; i < fieldCount; i++) {
    const input = document.createElement("input");
    input.type = "text";
    input.className = className;
    row.appendChild(input);
  }

  // Add Delete button
  const deleteButton = document.createElement("button");
  deleteButton.textContent = "삭제";
  deleteButton.className = "delete-row-button";
  deleteButton.onclick = () => deleteRow(deleteButton);
  row.appendChild(deleteButton);

  container.appendChild(row);
}

function deleteRow(button) {
  const row = button.parentElement;
  row.parentElement.removeChild(row);
}

function getInputData(sheetClass) {
  const container = document.getElementById(sheetClass + "-container");
  const rows = Array.from(container.querySelectorAll(".input-row"));
  return rows.map((row) => {
    const inputs = Array.from(row.querySelectorAll("input"));
    return inputs.map((input) => input.value || "-");
  });
}

function generateCombinations() {
  const sheet1 = getInputData("sheet1");
  const sheet2 = getInputData("sheet2");
  const sheet3 = getInputData("sheet3");

  const resultTableBody = document.querySelector("#result-table tbody");
  resultTableBody.innerHTML = ""; // Clear previous results

  const combinations = [];

  sheet1.forEach((row1) => {
    sheet2.forEach((row2) => {
      sheet3.forEach((row3) => {
        const combinedRow = [...row1, ...row2, ...row3];
        combinations.push(combinedRow);

        const tr = document.createElement("tr");
        combinedRow.forEach((value) => {
          const cell = document.createElement("td");
          cell.textContent = value;
          tr.appendChild(cell);
        });
        resultTableBody.appendChild(tr);
      });
    });
  });

  return combinations;
}

function exportToExcel() {
  const combinations = generateCombinations();
  const ws = XLSX.utils.aoa_to_sheet(combinations);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Combinations");
  XLSX.writeFile(wb, "combinations.xlsx");
}

/**
 * Excel 파일 업로드 핸들러
 * @param {Event} event - file input change event
 * @param {string} containerId - e.g. 'sheet1-container'
 * @param {string} className - e.g. 'sheet1'
 * @param {number} fieldCount - Number of columns per row for that dataset
 */
function handleFileUpload(event, containerId, className, fieldCount) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // 첫 번째 시트를 가져온다고 가정
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // 기존 데이터 제거
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    // jsonData는 2차원 배열 형태: 각 배열의 요소는 한 행
    // 행마다 fieldCount개 컬럼만큼 값을 가져와서 Row 생성
    jsonData.forEach((rowData) => {
      const newRow = document.createElement("div");
      newRow.className = "input-row";

      // rowData를 fieldCount개까지 자르고, 부족하면 빈칸
      for (let i = 0; i < fieldCount; i++) {
        const input = document.createElement("input");
        input.type = "text";
        input.className = className;
        input.value = rowData[i] !== undefined ? rowData[i] : "";
        newRow.appendChild(input);
      }

      const deleteButton = document.createElement("button");
      deleteButton.textContent = "삭제";
      deleteButton.className = "delete-row-button";
      deleteButton.onclick = () => deleteRow(deleteButton);
      newRow.appendChild(deleteButton);

      container.appendChild(newRow);
    });
  };

  reader.readAsArrayBuffer(file);
}
