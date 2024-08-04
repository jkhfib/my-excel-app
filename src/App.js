import React, { useRef, useState } from 'react';
import { HotTable } from '@handsontable/react';
import 'handsontable/dist/handsontable.full.css';
import * as XLSX from 'xlsx';
import './App.css';

const App = () => {
  const hotTableComponent = useRef(null);
  const [data, setData] = useState([[""]]);
  const [cellMeta, setCellMeta] = useState({});
  const [fileName, setFileName] = useState('edited-data.xlsx');
  const [inputFileName, setInputFileName] = useState('');

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return; // ファイルが選択されなかった場合は何もしない

    setInputFileName(file.name); // 入力ファイル名を保存

    const reader = new FileReader();

    reader.onload = (e) => {
      const arrayBuffer = e.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      setData(jsonSheet);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleAfterChange = (changes, source) => {
    if (source === 'loadData' || !changes) return;

    const newCellMeta = { ...cellMeta };

    changes.forEach(([row, col]) => {
      if (!newCellMeta[row]) newCellMeta[row] = {};
      newCellMeta[row][col] = true;
    });

    setCellMeta(newCellMeta);
  };

  const getCellMeta = () => {
    const cells = [];
    for (const row in cellMeta) {
      for (const col in cellMeta[row]) {
        if (cellMeta[row][col]) {
          cells.push({ row: Number(row), col: Number(col), className: 'cell-edited' });
        }
      }
    }
    return cells;
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.aoa_to_sheet(data, { header: 1 });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, fileName);
  };

  return (
    <div>
      <h1>Excel Editor</h1>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      {inputFileName && <p>Input File: {inputFileName}</p>}
      <HotTable
        ref={hotTableComponent}
        data={data}
        colHeaders={true}
        rowHeaders={true}
        width="600"
        height="300"
        licenseKey="non-commercial-and-evaluation"
        afterChange={handleAfterChange}
        cells={getCellMeta()}
      />
      <div>
        <input
          type="text"
          value={fileName}
          onChange={(e) => setFileName(e.target.value)}
          placeholder="Enter file name"
        />
        <button onClick={exportToExcel}>Export to Excel</button>
      </div>
    </div>
  );
};

export default App;
