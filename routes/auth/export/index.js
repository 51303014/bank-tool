const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const path = require("path");
const Excel = require("exceljs");

const DocType = {
  HDTG: 1,
  BM01: 2,
  BM02: 3,
  BM04: 4,
}

const getCellValue = (row, cellIndex) => {
  const cell = row.getCell(cellIndex);
  return cell && cell.value ? cell.value.toString() : '';
};

const exportFile = async (req, res) => {
  try {
    const docType = req.body.docType;
    const file = req.file;
    const workbook = new Excel.Workbook();
    const contentExcel = await workbook.xlsx.readFile(file.path);
    if (!file) {
      throw new Error('file.error.notFound');
    }
    switch (+docType) {
      case DocType.HDTG:
        const worksheet = contentExcel.getWorksheet(1);
        const rowStartIndex = 2;
        const numberOfRows = worksheet.rowCount - 1;
        const rows = worksheet.getRows(rowStartIndex, numberOfRows) ?? [];
        rows.map(async row => {
          try {
            // Load the docx file as binary content
            const content = fs.readFileSync(
                "./file-doc/BM02_KHTC_THU_THAP.doc",
                "binary"
            );
            // Unzip the content of the file
            const zip = new PizZip(content);
            // This will parse the template, and will throw an error if the template is
            // invalid, for example, if the template is "{user" (no closing tag)
            const doc = new Docxtemplater(zip, {
              paragraphLoop: true,
              linebreaks: true,
            });
            doc.render({
              CompanyName: getCellValue(row, 1),
              BusinessCode: getCellValue(row, 2),
              DateOfFirstIssue: getCellValue(row, 14),
              PlaceOfIssue: getCellValue(row, 16),
              CompanyAddress: getCellValue(row, 5),
              CompanyPhoneNumber: getCellValue(row, 7),
              CompanyEmail: getCellValue(row, 9),
              FullName: getCellValue(row, 19),
            });
            // Get the zip document and generate it as a nodebuffer
            const buf = doc.getZip().generate({
              type: "nodebuffer",
              // compression: DEFLATE adds a compression step.
              // For a 50MB output document, expect 500ms additional CPU time
              compression: "DEFLATE",
            });
            // buf is a nodejs Buffer, you can either write it to a
            // file or res.send it with express for example.
            fs.writeFileSync(path.resolve(__dirname, `output-${getCellValue(row, 2)}.docx`), buf);
            return res.status(200).send({ status: 200 });
          } catch (e) {
            return res.status(400).send({ status: 400, message: e.message });
          }
        });
        break;
      case DocType.BM01:
        const worksheetBM01 = contentExcel.getWorksheet(1);
        const rowStartIndexBM01 = 2;
        const numberOfRowsBM01 = worksheetBM01.rowCount - 1;
        const rowsBM01 = worksheetBM01.getRows(rowStartIndexBM01, numberOfRowsBM01) ?? [];
        rowsBM01.map(async row => {
          try {
            // Load the docx file as binary content
            const content = fs.readFileSync(
                "./file-doc/3502287961_BM01_COLLECT.docx",
                "binary"
            );
            console.log(3333);
            // Unzip the content of the file
            const zip = new PizZip(content);
            console.log(4444);
            // This will parse the template, and will throw an error if the template is
            // invalid, for example, if the template is "{user" (no closing tag)
            const doc = new Docxtemplater(zip, {
              paragraphLoop: true,
              linebreaks: true,
            });
            console.log(5555);
            doc.render({
              CompanyName: getCellValue(row, 1),
              // BusinessCode: getCellValue(row, 2),
              // DateOfFirstIssue: getCellValue(row, 14),
              // PlaceOfIssue: getCellValue(row, 16),
              // CompanyAddress: getCellValue(row, 5),
              // CompanyPhoneNumber: getCellValue(row, 7),
              // CompanyEmail: getCellValue(row, 9),
              // FullName: getCellValue(row, 19),
            });
            console.log(6666);
            // Get the zip document and generate it as a nodebuffer
            const buf = doc.getZip().generate({
              type: "nodebuffer",
              // compression: DEFLATE adds a compression step.
              // For a 50MB output document, expect 500ms additional CPU time
              compression: "DEFLATE",
            });
            // buf is a nodejs Buffer, you can either write it to a
            // file or res.send it with express for example.
            console.log(22222);
            fs.writeFileSync(path.resolve(__dirname, `output-${getCellValue(row, 2)}.docx`), buf);
            return res.status(200).send({ status: 200 });
          } catch (e) {
            return res.status(400).send({ status: 400, message: e.message });
          }
        });
        break;
      case DocType.BM02:
        const worksheetBM02 = contentExcel.getWorksheet(1);
        const rowStartIndexBM02 = 2;
        const numberOfRowsBM02 = worksheetBM02.rowCount - 1;
        const rowsBM02 = worksheetBM02.getRows(rowStartIndexBM02, numberOfRowsBM02) ?? [];
        rowsBM02.map(async row => {
          try {
            // Load the docx file as binary content
            const content = fs.readFileSync(
                "./file-doc/BM02_KHTC_THU_THAP.doc",
                "binary"
            );
            // Unzip the content of the file
            const zip = new PizZip(content);
            // This will parse the template, and will throw an error if the template is
            // invalid, for example, if the template is "{user" (no closing tag)
            const doc = new Docxtemplater(zip, {
              paragraphLoop: true,
              linebreaks: true,
            });
            doc.render({
              CompanyName: getCellValue(row, 1),
              BusinessCode: getCellValue(row, 2),
              DateOfFirstIssue: getCellValue(row, 14),
              PlaceOfIssue: getCellValue(row, 16),
              CompanyAddress: getCellValue(row, 5),
              CompanyPhoneNumber: getCellValue(row, 7),
              CompanyEmail: getCellValue(row, 9),
              FullName: getCellValue(row, 19),
            });
            // Get the zip document and generate it as a nodebuffer
            const buf = doc.getZip().generate({
              type: "nodebuffer",
              // compression: DEFLATE adds a compression step.
              // For a 50MB output document, expect 500ms additional CPU time
              compression: "DEFLATE",
            });
            // buf is a nodejs Buffer, you can either write it to a
            // file or res.send it with express for example.
            fs.writeFileSync(path.resolve(__dirname, `output-${getCellValue(row, 2)}.docx`), buf);
            return res.status(200).send({ status: 200 });
          } catch (e) {
            return res.status(400).send({ status: 400, message: e.message });
          }
        });
        break;
      case DocType.BM04:
        const worksheetBM04 = contentExcel.getWorksheet(1);
        const rowStartIndexBM04 = 2;
        const numberOfRowsBM04 = worksheetBM04.rowCount - 1;
        const rowsBM04 = worksheetBM04.getRows(rowStartIndexBM04, numberOfRowsBM04) ?? [];
        rowsBM04.map(async row => {
          try {
            // Load the docx file as binary content
            const content = fs.readFileSync(
                "./file-doc/BM02_KHTC_THU_THAP.doc",
                "binary"
            );
            // Unzip the content of the file
            const zip = new PizZip(content);
            // This will parse the template, and will throw an error if the template is
            // invalid, for example, if the template is "{user" (no closing tag)
            const doc = new Docxtemplater(zip, {
              paragraphLoop: true,
              linebreaks: true,
            });
            doc.render({
              CompanyName: getCellValue(row, 1),
              BusinessCode: getCellValue(row, 2),
              DateOfFirstIssue: getCellValue(row, 14),
              PlaceOfIssue: getCellValue(row, 16),
              CompanyAddress: getCellValue(row, 5),
              CompanyPhoneNumber: getCellValue(row, 7),
              CompanyEmail: getCellValue(row, 9),
              FullName: getCellValue(row, 19),
            });
            // Get the zip document and generate it as a nodebuffer
            const buf = doc.getZip().generate({
              type: "nodebuffer",
              // compression: DEFLATE adds a compression step.
              // For a 50MB output document, expect 500ms additional CPU time
              compression: "DEFLATE",
            });
            // buf is a nodejs Buffer, you can either write it to a
            // file or res.send it with express for example.
            fs.writeFileSync(path.resolve(__dirname, `output-${getCellValue(row, 2)}.docx`), buf);
            return res.status(200).send({ status: 200 });
          } catch (e) {
            return res.status(400).send({ status: 400, message: e.message });
          }
        });
        break;
    }
  } catch (e) {
    return res.status(400).send({ status: 400, message: e.message });
  }
};

module.exports = exportFile;
