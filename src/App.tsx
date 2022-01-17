import logo from "./logo.svg";
import "./App.css";
import xlsx from "xlsx";
import { data } from "./data";

function App() {
  const dataToSting = JSON.stringify(data, undefined, 4);
  const EXCEL_TYPE =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
  const EXCEL_EXTENSION = ".xlsx";

  const dlAsExcel = () => {
    const workSheet = xlsx.utils.json_to_sheet(data);
    const workBook = {
      Sheets: {
        'data': workSheet,
      },
      SheetNames: ['data'],
    };
    console.log(workBook);
    const excelBuffer = xlsx.write(workBook, {
      bookType: "xlsx",
      type: "array",
    });
    console.log(excelBuffer);
    saveAsExcel(excelBuffer, 'myFile')
  };

  const saveAsExcel = (buffer: BlobPart, filename: string) => {
    const data = new Blob([buffer], {type: EXCEL_TYPE});
      const link = document.createElement('a');
      link.href = window.URL.createObjectURL(data);
      const fileName = filename+EXCEL_EXTENSION;
      link.download = fileName;
      link.click();
  }

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <div>{dataToSting}</div>
        <button onClick={dlAsExcel}>Generate Xlsx File</button>
      </header>
    </div>
  );
}

export default App;
