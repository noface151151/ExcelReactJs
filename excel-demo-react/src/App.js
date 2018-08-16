import React, { Component } from "react";
import "./App.css";
import * as ExcelLib from "./Excel/ExcelLib";

class App extends Component {
  downloadContract() {
    var oReq = new XMLHttpRequest();

    var URLToPDF = "http://localhost:25051/api/Home";

    oReq.open("POST", URLToPDF, true);

    oReq.responseType = "blob";

    oReq.onload = function() {
      // Once the file is downloaded, open a new window with the PDF
      // Remember to allow the POP-UPS in your browser
      const file = new Blob([oReq.response], {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const fileURL = URL.createObjectURL(file);

      window.open(fileURL, "_blank");
    };

    oReq.send();
  }

  ExportExcel = () => {
    const SheetName='test';
    const workbook = ExcelLib.CreateNewWookBook(SheetName);
    const worksheet = {};
    var range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
    for (let i = 1; i <= 10; i++) {
      ExcelLib.SetCellValue(
       100000000,
        "A",
        i,
        range,
        worksheet,
        {fill: {patternType: "solid", fgColor: {rgb: "FFFF0000"}}}
      );
    }
    ExcelLib.SetRangeWorksheet(worksheet, range);
    ExcelLib.AddWorkSheetToWorkbook(workbook, worksheet,SheetName);
    ExcelLib.ExportExcel("Test", workbook);
  };

  render() {
    return (
      <div className="App">
        <button onClick={this.ExportExcel}>Click</button>
      </div>
    );
  }
}

export default App;
