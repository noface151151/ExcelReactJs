import XLSX from "xlsx";
import { saveAs } from "file-saver";
import WriteLog from '../service/Logservice';

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
 * Khai báo range cho sheet
 * @param {object} worksheet 
 * @param {object} range 
 */
const SetRangeWorksheet = (worksheet, range) => {
  worksheet["!ref"] = XLSX.utils.encode_range(range);
};

/**
 *
 * @param {string} rangeString ví dụ "A6:C6"
 */
const GetRange = rangeString => {
  return XLSX.utils.decode_range(rangeString);
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
 * Insert tên cột
 * @param {stringArray} columnNames 
 * @param {object} worksheet 
 * @param {object} range 
 * @param {number} rowCount 
 */
const InsertColumnName = (columnNames, worksheet, range, rowCount, style) => {
  if (columnNames !== null && columnNames.length > 0) {
    columnNames.forEach((col, index) => {
      fixRange(range, rowCount - 1, index);
      SetCellValue(
        col,
        XLSX.utils.encode_col(index),
        rowCount,
        range,
        worksheet,
        style
      );
    });
  }
};

/**
 * Gán giá trị cho cell
 * @param {object} value -
 * @param {string} column - tên cột (A,B,C,...)
 * @param {number} row 
 * @param {object} range
 * @param {object} worksheet 
 * @param {object} style 
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
    // } else if (value instanceof Date) {
    //   cell.t = "n";
    //   cell.z = XLSX.SSF._table[14];
    //   cell.v = dateToNumber(cell.v);
  } else {
    cell.v = value;
    cell.t = "s";
  }
  if (style !== null) {
    cell.s = style;
  }
  fixRange(range, row - 1, XLSX.utils.decode_col(column));
  worksheet[column + row] = cell;
};

/**
 * Export Excel
 * @param {string} fimeName 
 * @param {any} workbook 
 */
const ExportExcel = (fimeName, workbook, worksheet, range, SheetName) => {
  if (workbook === null || worksheet === null) {
    return;
  }
  SetRangeWorksheet(worksheet, range);
  AddWorkSheetToWorkbook(workbook, worksheet, SheetName);
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
/**
 * Chuyển file Excel sang Json
 * @param {binary} Filebinary 
 */
const ConvertExcelToJson = (Filebinary,sheetName) => {
  var workbook = XLSX.read(Filebinary, { type: "binary" });
  var sheet_name_list = workbook.SheetNames;
  const index = sheet_name_list.indexOf(sheetName)
  WriteLog( XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[index]], {
    raw: true
  }));
};

/**
 * Chuyển đổi index sang column, ví dụ index:0 => column: A
 * @param {number} columnIndex 
 */
const IndexToColumn = columnIndex => {
  return XLSX.utils.encode_col(columnIndex);
};

/**
 * Chuyển đổi column sang index, ví dụ column A => index: 0
 * @param {string} columnName 
 */
const ColumnToIndex = columnName => {
  return XLSX.utils.decode_col(columnName);
};
//các hàm hỗ trợ
const getFileNameWithExtension = (filename, extension) => {
  return `${filename}.${extension}`;
};

// const dateToNumber = (v, date1904) => {
//   if (date1904) {
//     v += 1462;
//   }

//   var epoch = Date.parse(v);

//   return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
// };

const strToArrBuffer = s => {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);

  for (var i = 0; i !== s.length; ++i) {
    view[i] = s.charCodeAt(i) & 0xff;
  }

  return buf;
};

const fixRange = (range, rowCount, column) => {
  // if (range.e.r < rowCount) {
  //   range.e.r = range.e.r + 1;
  // }

  // if (range.e.c < column) {
  //   range.e.c = range.e.c + column;
  // }

  if (range.s.r > rowCount) range.s.r = rowCount;
  if (range.s.c > column) range.s.c = column;
  if (range.e.r < rowCount) range.e.r = rowCount;
  if (range.e.c < column) range.e.c = column;
};

export {
  CreateNewWookBook,
  AddWorkSheetToWorkbook,
  SetCellValue,
  MergeCell,
  ExportExcel,
  SetRangeWorksheet,
  GetRange,
  InsertColumnName,
  ConvertExcelToJson
};
