import XLSX from "xlsx";
import { saveAs } from "file-saver";

/**
 * Khởi tạo một workbook mới
 */
const CreateNewWookBook = sheetname => {
  return {
    SheetNames: [sheetname],
    Sheets: {}
  };
};

/**
 * Thêm Sheet vào Workbook
 * @param {any} workbook - nếu chưa có workbook thì dùng hàm CreateNewWookBook để tạo mới 
 * @param {any} worksheet - worksheet
 * @param {string} sheetname - tên sheet
 */
const AddWorkSheetToWorkbook = (workbook, worksheet, sheetname) => {
  if (workbook === null || worksheet === null) {
    return;
  }
  workbook.Sheets["test"] = worksheet;
};

/**
 * Gán giá trị cho cell
 * @param {any} value - nếu có style thì truyền theo cấu trúc: {value:'...',style:{...}}
 * @param {string} column - tên cột (A,B,C,...)
 * @param {number} row 
 * @param {object} range
 * @param {any} worksheet 
 */
const SetCellValue = (value, column, row, range, worksheet, style) => {
  var cell = {};
  if (value === null) {
    return;
  }
  if (typeof value === "number") {
    cell.v = value;
    cell.t = "n";
    cell.z = "#,##0";
  } else if (typeof value === "boolean") {
    cell.v = value;
    cell.t = "b";
  } else if (value instanceof Date) {
    cell.t = "n";
    cell.z = XLSX.SSF._table[14];
    cell.v = dateToNumber(cell.v);
  } else {
    cell.v = value;
    cell.t = "s";
  }
  cell.s=style;
  fixRange(range, row - 1, XLSX.utils.decode_col(column));

  worksheet[column + row] = cell;
};

/**
 * Khai báo range cho sheet
 * @param {object} worksheet 
 * @param {object} range 
 */
const SetRangeWorksheet = (worksheet, range) => {
  worksheet["!ref"] = XLSX.utils.encode_range(range);
};

/**
 * MergeCell
 * @param {any} worksheet 
 * @param {string} rangeString - range cần merge, ví dụ "A6:C6"
 */
const MergeCell = (worksheet, rangeString) => {
  worksheet["!merges"] = [XLSX.utils.decode_range(rangeString)];
};

/**
 * 
 * @param {string} rangeString 
 */
const GetRange = rangeString => {
  return XLSX.utils.decode_range(rangeString);
};
/**
 * Export Excel
 * @param {string} fimeName 
 * @param {any} workbook 
 */
const ExportExcel = (fimeName, workbook) => {
  if (workbook === null) {
    return;
  }
  const fileExtension = "xlsx";
  const fileNameCheck = fimeName === null ? "" : fimeName;
  const wbout = XLSX.write(workbook, {
    bookType: fileExtension,
    bookSST: true,
    type: "binary",
    cellStyles: true
  });
  saveAs(
    new Blob([strToArrBuffer(wbout)], { type: "application/octet-stream" }),
    getFileNameWithExtension(fileNameCheck, fileExtension)
  );
};

//các hàm hỗ trợ
const getFileNameWithExtension = (filename, extension) => {
  return `${filename}.${extension}`;
};

const dateToNumber = (v, date1904) => {
  if (date1904) {
    v += 1462;
  }

  var epoch = Date.parse(v);

  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};

const strToArrBuffer = s => {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);

  for (var i = 0; i !== s.length; ++i) {
    view[i] = s.charCodeAt(i) & 0xff;
  }

  return buf;
};

const fixRange = (range, rowCount, column) => {
  if (range.e.r < rowCount) {
    range.e.r = range.e.r + 1;
  }

  if (range.e.c < column) {
    range.e.c = range.e.c + column;
  }
};

export {
  CreateNewWookBook,
  AddWorkSheetToWorkbook,
  SetCellValue,
  MergeCell,
  ExportExcel,
  SetRangeWorksheet,
  GetRange
};
