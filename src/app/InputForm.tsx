// components/ExcelFileReader.tsx
"use client";
import React, { useState } from "react";
import { read, WorkBook, utils, WorkSheet } from "xlsx";
import { Workbook, Worksheet } from "exceljs";

function ExcelFileReader() {
  const [workbook, setWorkbook] = useState<Workbook>(new Workbook());
  const [csvFile, setCsvFile] = useState<File | null>(null);

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0] as File;
    if (file) {
      workbook.xlsx.load(await file.arrayBuffer());
      setWorkbook(workbook);
    }
  };

  const handleChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] as File;
    if (file) {
      setCsvFile(file);
    }
  };

  const handleSubmit = () => {
    const qty: {
      [key: string]: {
        qty: number;
        id?: string;
        name?: string;
      };
    } = {};

    let qtyCell = {
      col: "",
      row: "",
    };
    let cellEnd = {
      col: "",
      row: "",
    };

    let nameCell = {
      col: "",
      row: "",
    };

    let idCell = {
      col: "",
      row: "",
    };

    workbook.getWorksheet("Sheet1").eachRow((row) =>
      row.eachCell((cell) => {
        if (cell.text.trim() === "الرصيد") {
          qtyCell = {
            col: cell.col,
            row: cell.row,
          };
        }
        if (cell.text.trim() === "الـبـيـــان") {
          idCell = {
            col: cell.col,
            row: cell.row,
          };
        }
        if (cell.text.trim() === "رمز المادة") {
          nameCell = {
            col: cell.col,
            row: cell.row,
          };
        }
        if (cell.text.trim() === "المجاميع") {
          cellEnd = {
            col: cell.col,
            row: cell.row,
          };
        }
      })
    );

    workbook
      .getWorksheet("Sheet1")
      .getColumn(qtyCell.col)
      .eachCell((cell) => {
        if (
          cell.text.trim() &&
          parseInt(cell.row) > parseInt(qtyCell.row) &&
          parseInt(cell.row) < parseInt(cellEnd.row)
        ) {
          qty[cell.row] = {
            qty: parseInt(cell.text),
          };
        }
      });

    workbook
      .getWorksheet("Sheet1")
      .getColumn(nameCell.col)
      .eachCell((cell) => {
        if (
          cell.text.trim() &&
          parseInt(cell.row) > parseInt(qtyCell.row) &&
          parseInt(cell.row) < parseInt(cellEnd.row)
        ) {
          qty[cell.row].name = cell.text.trim();
        }
      });

    workbook
      .getWorksheet("Sheet1")
      .getColumn(idCell.col)
      .eachCell((cell) => {
        if (
          cell.text.trim() &&
          parseInt(cell.row) > parseInt(qtyCell.row) &&
          parseInt(cell.row) < parseInt(cellEnd.row)
        ) {
          qty[cell.row].id = cell.text.trim();
        }
      });

    console.log("qtyCell", csvFile?.toString());
  };

  return (
    <div className="p-10 flex justify-center">
      <h1>Excel File Reader</h1>
      <input type="file" onChange={handleFileChange} accept=".xlsx" />
      <br />
      <input type="file" onChange={handleChange} accept=".csv" />

      <button
        className="bg-green-600 text-yellow-50 px-2 border-md"
        onClick={handleSubmit}
      >
        submit
      </button>
    </div>
  );
}

export default ExcelFileReader;
