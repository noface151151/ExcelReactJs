import XLSX from "xlsx";
import { saveAs } from "file-saver";

const data = [
  [1, 2, 3],
  [true, false, null, "sheetjs"],
  ["foo    bar", "baz", new Date("2014-02-19T14:30Z"), "0.3"],
  ["baz", null, "\u0BEE", 3.14159],
  ["hidden"],
  ["visible"]
];

if (isNaN(data[2][2].getYear()))
  data[2][2] = new Date(Date.UTC(2014, 1, 19, 14, 30, 0));

const ws_name = "SheetJS";

var wscols = [
  { wch: 6 }, // "characters"
  { wpx: 50 }, // "pixels"
  { hidden: true } // hide column
];
var wsrows = [
  { hpt: 12 }, // "points"
  { hpx: 16 }, // "pixels"
  { hpx: 24, level: 3 },
  { hidden: true }, // hide row
  { hidden: false }
];

const CheckData = () => {
  console.log("Sheet Name: " + ws_name);
  console.log("Data: ");
  var i = 0;
  for (i = 0; i !== data.length; ++i) console.log(data[i]);
  console.log("Columns :");
  for (i = 0; i !== wscols.length; ++i) console.log(wscols[i]);
};

var wb = XLSX.utils.book_new();

var ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true });

const AddWorkSheetToWorkbook = () => {
  XLSX.utils.book_append_sheet(wb, ws, ws_name);
};

const SimpleFormula = () => {
  ws["C1"].f = "A1+B1";
  ws["C2"] = { t: "n", f: "A1+B1" };
};
const SingleCellArrayFormula = () => {
  XLSX.utils.sheet_set_array_formula(ws, "D1:D1", "SUM(A1:C1+A1:C1)");
};
const MultiCellArrayFormula = () => {
  XLSX.utils.sheet_set_array_formula(ws, "E1:E4", "TRANSPOSE(A1:D1)");
  ws["!ref"] = "A1:E6"; //khi insert dòng/cột thì phải khai báo lại range
};
const ColumnProps = () => {
  ws["!cols"] = wscols;
};

const RowProps = () => {
  ws["!rows"] = wsrows;
};
const Customformat = () => {
  var custfmt = '"This is "\\ 0.0';
  XLSX.utils.cell_set_number_format(ws["C2"], custfmt);
};
const MergeCell = () => {
  ws["!merges"] = [XLSX.utils.decode_range("A6:C6")];
  console.log("JSON Data:");
  console.log(XLSX.utils.sheet_to_json(ws, { header: 1 }));
};
const ExportExcel = () => {
  AddWorkSheetToWorkbook();
  //SimpleFormula();
  //SingleCellArrayFormula();
  //MultiCellArrayFormula();
  //ColumnProps();
  //RowProps();
  //Customformat();
  MergeCell();
  const fileExtension = "xlsx";
  const fileName = "test";
  const wbout = XLSX.write(wb, {
    bookType: fileExtension,
    bookSST: true,
    type: "binary"
  });
  saveAs(
    new Blob([strToArrBuffer(wbout)], { type: "application/octet-stream" }),
    getFileNameWithExtension(fileName, fileExtension)
  );
};

//util
const getFileNameWithExtension = (filename, extension) => {
  return `${filename}.${extension}`;
};

const strToArrBuffer = s => {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);

  for (var i = 0; i !== s.length; ++i) {
    view[i] = s.charCodeAt(i) & 0xff;
  }

  return buf;
};
const dateToNumber = (v, date1904) => {
  if (date1904) {
    v += 1462;
  }
  var epoch = Date.parse(v);

  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};
function getHeaderCell(v, cellRef, ws) {
  var cell = {};
  var headerCellStyle = { font: { bold: true } };
  cell.v = v;
  cell.t = "s";
  cell.s = headerCellStyle;
  ws[cellRef] = cell;
}

function getCell(v, cellRef, ws) {
  var cell = {};
  if (v === null) {
    return;
  }
  if (typeof v === "number") {
    cell.v = v;
    cell.t = "n";
  } else if (typeof v === "boolean") {
    cell.v = v;
    cell.t = "b";
  } else if (v instanceof Date) {
    cell.t = "n";
    cell.z = XLSX.SSF._table[14];
    cell.v = dateToNumber(cell.v);
  } else if (typeof v === "object") {
    cell.v = v.value;
    cell.s = v.style;
  } else {
    cell.v = v;
    cell.t = "s";
  }
  ws[cellRef] = cell;
}

function fixRange(range, R, C, rowCount, xSteps, ySteps) {
  if (range.s.r > R + rowCount) {
    range.s.r = R + rowCount;
  }

  if (range.s.c > C + xSteps) {
    range.s.c = C + xSteps;
  }

  if (range.e.r < R + rowCount) {
    range.e.r = R + rowCount;
  }

  if (range.e.c < C + xSteps) {
    range.e.c = C + xSteps;
  }
}

export { ExportExcel };
