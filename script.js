const spreadSheetContainer = document.querySelector(".spreadsheet_container");
const ROWS = 10;
const COLS = 10;
const spreadsheet = [];
const exportBtn = document.querySelector("#export_btn");

const alphabets = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];

// 10 * 10 표 배열 만들기
const initSpreadsheet = () => {
  for (let i = 0; i < ROWS; i++) {
    let spreadsheetRow = [];
    for (let j = 0; j < COLS; j++) {
      let cellData = "";
      // header 만들기
      let isHeader = false;
      // disabled 속성 추가
      let disabled = false;

      // 모든 row 첫 column에 숫자 입력
      if (j == 0) {
        cellData = i;
        isHeader = true;
        disabled = true;
      }

      // 모든 column 첫 row에 알파벳 입력
      if (i == 0) {
        cellData = alphabets[j - 1];
        isHeader = true;
        disabled = true;
      }

      // 첫 row 첫 column은 "";
      if (!cellData) {
        cellData = "";
      }

      // 첫 row의 column은 "";
      if (cellData <= 0) {
        cellData = "";
      }

      const rowName = i;
      const columnName = alphabets[j - 1];

      const cell = new Cell(
        isHeader,
        disabled,
        cellData,
        i,
        j,
        rowName,
        columnName,
        false
      );
      spreadsheetRow.push(cell);
    }
    spreadsheet.push(spreadsheetRow);
  }
  drawSheet();
  console.log("spreadsheet", spreadsheet);
};

// 생성된 배열을 객체로 변환

class Cell {
  constructor(
    isHeader,
    disabled,
    data,
    row,
    column,
    rowName,
    columnName,
    active = false
  ) {
    this.isHeader = isHeader;
    this.disable = disabled;
    this.data = data;
    this.row = row;
    this.column = column;
    this.rowName = rowName;
    this.columnName = columnName;
    this.active = active;
  }
}

// cell 생성

const createCellEl = (cell) => {
  const cellEl = document.createElement("input");
  cellEl.className = "cell";
  cellEl.id = "cell_" + cell.row + cell.column;
  cellEl.value = cell.data;
  cellEl.disabled = cell.disabled;

  // css를 위한 isHeader == true -> 글래스 추가
  if (cell.isHeader) {
    cellEl.classList.add("header");
  }

  // 셀 클릭 상호작용
  cellEl.onclick = (e) => handleCellClick(cell);
  cellEl.onchange = (e) => handelOnChange(e.target.value, cell);
  return cellEl;
};

// cell 렌더링
// 10개의 셀을 하나의 div로 감싸기

const drawSheet = () => {
  for (let i = 0; i < spreadsheet.length; i++) {
    const rowContainerEl = document.createElement("div");
    rowContainerEl.className = "cell_row";

    for (let j = 0; j < spreadsheet[i].length; j++) {
      const cell = spreadsheet[i][j];
      rowContainerEl.append(createCellEl(cell));
    }
    spreadSheetContainer.append(rowContainerEl);
  }
};

// 하이라이트 지우기
const clearHeaderActiveStatus = () => {
  const headers = document.querySelectorAll(".header");

  headers.forEach((header) => {
    header.classList.remove("active");
  });
};

const handleCellClick = (cell) => {
  // 클릭 시 하이라이트 지우기 실행
  clearHeaderActiveStatus();
  const columnHeader = spreadsheet[0][cell.column];
  const rowHeader = spreadsheet[cell.row][0];

  // column header, row header 요소 가져오기
  const columnHeaderEl = getElFromRowCol(columnHeader.row, columnHeader.column);
  const rowHeaderEl = getElFromRowCol(rowHeader.row, rowHeader.column);

  // 하이라이트 추가
  columnHeaderEl.classList.toggle("active");
  rowHeaderEl.classList.toggle("active");
  document.querySelector("#cell_status").innerHTML =
    cell.columnName + "" + cell.rowName;
};

const getElFromRowCol = (row, col) => {
  return document.querySelector("#cell_" + row + col);
};

// 버튼 클릭 시 엑셀 데이터 생성
exportBtn.onclick = (e) => {
  let csv = "";
  for (let i = 0; i < spreadsheet.length; i++) {
    if (i === 0) continue;
    csv +=
      spreadsheet[i]
        .filter((item) => !item.isHeader)
        .map((item) => item.data)
        .join(",") + "\r\n";
  }

  // 엑셀 파일 다운로드
  // 미디어 데이터를 크기, 타입, 송수신을 위해 작은 객체로 나누는 작업
  const csvObj = new Blob([csv]);

  // Blob 객체를 가리키는 URL 생성
  const csvUrl = URL.createObjectURL(csvObj);

  const a = document.createElement("a");
  a.href = csvUrl;
  a.download = "Spreadsheet File Name.csv";
  a.click();
  // 생성한 URL 폐기
  window.URL.revokeObjectURL(csvUrl);
  // 다운로드 받은 파일 저장 경로는 브라우저에서 설정
};

const handelOnChange = (data, cell) => {
  cell.data = data;
};

initSpreadsheet();
