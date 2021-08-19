const Excel = require("exceljs");
const fs = require("fs");
const path = require("path");

const ignoreCellRules = /:|2021/;
const FOLDER = "nalozi";

function transform(amount, string) {
  if (string.formula && string.result) {
    return string.result
      .replace(/(\d+)gr/g, (a, b) => +b / 1000)
      .replace(/\./g, ",")
      .replace(/0?,?\d+/g, (m) => (+m.replace(",", ".") * amount).toFixed(2).toString() + "kg")
      .replace(/gr/g, "")
      .replace(/(\d)\.00kgkom/g, (m, c1) => `${c1} kom`)
      .replace(/(\d+)\.\d+kgml/g, (m, c1) => (+c1 >= 1000 ? `${+c1 / 1000} L` : `${c1} ml`))
      .replace(/0\.00kg(?=0\.)/g, "")
      .replace(/((jaja|e) (\d+\.?\d+?kg\/\d+\.?\d+?kg))/g, (m, c1, c2, c3, c4) => {
        const nums = c3.replace(/kg/g, "").split("/");
        return `${c2} ${Math.ceil(amount * (+nums[0] / +nums[1]))}`;
      });
  }
  if (typeof string === "string") {
    return string
      .replace(/(\d+)gr/g, (a, b) => +b / 1000)
      .replace(/\./g, ",")
      .replace(/0?,?\d+/g, (m) => (+m.replace(",", ".") * amount).toFixed(2).toString() + "kg")
      .replace(/gr/g, "")
      .replace(/(\d)\.00kgkom/g, (m, c1) => `${c1} kom`)
      .replace(/(\d+)\.\d+kgml/g, (m, c1) => (+c1 >= 1000 ? `${+c1 / 1000} L` : `${c1} ml`))
      .replace(/0\.00kg(?=0\.)/g, "")
      .replace(/((jaja|e) (\d+\.?\d+?kg\/\d+\.?\d+?kg))/g, (m, c1, c2, c3, c4) => {
        const nums = c3.replace(/kg/g, "").split("/");
        return `${c2} ${Math.ceil(amount * (+nums[0] / +nums[1]))}`;
      });
  }
}

fs.readdir(`${process.cwd()}/${FOLDER}`, (err, files) => {
  files.forEach((file) => {
    if (file.endsWith(".xlsx")) {
      (async () => {
        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(`${process.cwd()}/${FOLDER}/${file}`);

        const worksheet = workbook.getWorksheet("Sheet1");

        worksheet.eachRow(function (row) {
          let amount = null;
          let normative = null;
          row.eachCell(function (cell) {
            if (!ignoreCellRules.test(String(cell.value))) {
              if (typeof cell.value === "number") {
                amount = cell.value;
                normative = null;
              }

              if (typeof cell.value !== "number") {
                normative = cell.value;
              }

              if (normative && amount > 1) {
                cell.value = transform(amount, normative);
                amount = null;
                normative = null;
              }
            }
          });
        });

        await workbook.xlsx.writeFile(`${process.cwd()}/${FOLDER}/${file}`);
      })();
    }
  });
});
