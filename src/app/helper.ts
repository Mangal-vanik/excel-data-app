import { Workbook } from "exceljs";
import Papa from "papaparse";
interface IQty {
  [key: string]: {
    qty: number;
    id?: string;
    name?: string;
  };
}

export interface ICsvResult {
  "Item Code": string;
  Quantity: string;
}

export const getDataFromExcel = (workbook: Workbook) => {
  const qty: IQty = {};

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
        nameCell = {
          col: cell.col,
          row: cell.row,
        };
      }
      if (cell.text.trim() === "رمز المادة") {
        idCell = {
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

  return qty;
};

export const getDataFromCsv = (file: any) =>
  new Promise((resolve) =>
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete(results) {
        resolve(results.data);
      },
    })
  );

export const getAccumulatedResults = async (
  qty: IQty,
  csvResults: ICsvResult[],
  workbook: Workbook
) => {
  const finalResults: {
    "Data Source": string;
    "Item Code": string;
    "Real Quantity": number;
    Description: string;
    Quantity: number;
    Difference: number;
  }[] = [];

  for (let item of csvResults) {
    const getItemFromQty = Object.values(qty).find(
      (singleItem) => singleItem.id === item["Item Code"]
    );

    if (getItemFromQty) {
      finalResults.push({
        "Data Source": `${item["Item Code"]},${item.Quantity}`,
        "Item Code": item["Item Code"],
        "Real Quantity": parseInt(item.Quantity),
        Description: getItemFromQty.name || "",
        Quantity: getItemFromQty.qty,
        Difference: getItemFromQty.qty - parseInt(item.Quantity),
      });
    }
  }

  let hashTagCell = {
    col: "",
    row: "",
  };

  workbook.getWorksheet("Sheet1").eachRow((row) =>
    row.eachCell((cell) => {
      if (cell.text.trim() === "#") {
        hashTagCell = {
          col: cell.col,
          row: cell.row,
        };
      }
    })
  );
  try {
    Object.keys(finalResults[0]).forEach((field, index) => {
      const cell = workbook
        .getWorksheet("Sheet1")
        .getCell(hashTagCell.row, hashTagCell.col + 3 + index);
      cell.value = field;
      cell.style = {
        font: {
          bold: true,
          size: 16,
        },
      };

      workbook
        .getWorksheet("Sheet1")
        .getColumn(hashTagCell.col + 3 + index).width = 20;
    });
    Object.values(finalResults).forEach((value, valueIndex) => {
      Object.keys(finalResults[0]).forEach((field, fieldIndex) => {
        workbook
          .getWorksheet("Sheet1")
          .getCell(
            hashTagCell.row + 1 + valueIndex,
            hashTagCell.col + 3 + fieldIndex
          ).value = (value as any)[field];
      });
    });
  } catch {}

  const blob = new Blob([await workbook.xlsx.writeBuffer()]);

  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "filename.xlsx";
  a.click();
  window.URL.revokeObjectURL(url);

  return finalResults;
};
