import React, { Component } from "react";
import "./App.css";
import  * as ExcelLib from './Excel/ExcelNew';
 


class App extends Component {
  downloadContract() {
    var oReq = new XMLHttpRequest();

    var URLToPDF = "http://localhost:25051/api/Home";

    oReq.open("POST", URLToPDF, true);

    oReq.responseType = "blob";

    oReq.onload = function() {
        // Once the file is downloaded, open a new window with the PDF
        // Remember to allow the POP-UPS in your browser
        const file = new Blob([oReq.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const fileURL = URL.createObjectURL(file);

        window.open(fileURL, "_blank");
    };

    oReq.send();
}

  render() {
    return (
      <div className="App">
       <button  onClick={ExcelLib.ExportExcel} >Click</button>
     
      </div>
    );
  }
}

export default App;
