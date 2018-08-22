import React, { Component } from "react";
import "./App.css";
import TestPrint from "./PrintReport/PrintRePort";
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
    const SheetName = "test";
    const workbook = ExcelLib.CreateNewWookBook(SheetName);
    const worksheet = {};
    var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
    var rowcount = 1;
    const style = {
      font: {
        bold: true
      },
      alignment: {
        vertical: "center"
      },
      border: {
        top: {
          style: "thin"
        },
        bottom: {
          style: "thin"
        },
        left: {
          style: "thin"
        },
        right: {
          style: "thin"
        }
      }
    };
    // const title = "TEST EXPORT EXCEL";
    // ExcelLib.MergeCell(worksheet, "G1:K1");
    // ExcelLib.SetCellValue(title, "G", 1, range, worksheet, null);
    // rowcount += 1;
    const column = ["ProductName", "Price", "Quantity"];
    ExcelLib.InsertColumnName(column, worksheet, range, rowcount, style);
    rowcount += 1;
    const data = [
      {
        productName: "Banana",
        productPrice: 100000,
        quantity: 20
      },
      {
        productName: "Potato",
        productPrice: 100000,
        quantity: 20
      },
      {
        productName: "Apple",
        productPrice: 100000,
        quantity: 20
      }
    ];

    for (var i = 0; i < data.length; i++, rowcount++) {
      ExcelLib.SetCellValue(
        data[i].productName,
        "A",
        rowcount,
        range,
        worksheet,
        style
      );
      ExcelLib.SetCellValue(
        data[i].productPrice,
        "B",
        rowcount,
        range,
        worksheet,
        style
      );
      ExcelLib.SetCellValue(
        data[i].quantity,
        "C",
        rowcount,
        range,
        worksheet,
        style
      );
    }

    ExcelLib.ExportExcel("Test", workbook, worksheet, range, SheetName);
  };

  handleChange = e => {
    var files = e.target.files;
    var file;
    if (!files || files.length === 0) return;
    file = files[0];
    var value = file.name,
      ext = value.split(".").pop();
    if (ext !== "xlsx") {
      console.log("false");
      return;
    }
    var fileReader = new FileReader();
    fileReader.onload = function(e) {
      // call 'xlsx' to read the file
      var binary = "";
      var bytes = new Uint8Array(e.target.result);
      var length = bytes.byteLength;
      for (var i = 0; i < length; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      ExcelLib.ConvertExcelToJson(binary,'test');
      //  var oFile = XLSX.read(binary, {type: 'binary', cellDates:true, cellStyles:true});
    };
    fileReader.readAsArrayBuffer(file);
  };

  render() {
    return (
      <div className="App">
        <button onClick={this.ExportExcel}>Click</button>
        <input type="file" onChange={e => this.handleChange(e)} />
        <hr />
        <TestPrint />
      </div>
    );
  }
}

export default App;
