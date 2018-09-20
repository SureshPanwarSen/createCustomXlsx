const Excel = require('exceljs');

function exportXlsx (data) {
    let sheetName = data.sheetName || 'statement1';
    var workbook = new Excel.Workbook();
    workbook.creator = 'Sankalp';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.views = [
        {
            x: 0, y: 0, width: 10000, height: 20000,
            firstSheet: 0, activeTab: 1, visibility: 'visible'
        }
    ];

    var worksheet = workbook.addWorksheet(sheetName, {
        pageSetup: {paperSize: 9, orientation: 'landscape'}
    });
    worksheet.pageSetup.margins = {
        left: 0.7, right: 0.7,
        top: 0.75, bottom: 0.75,
        header: 0.3, footer: 0.3
    };
    worksheet.pageSetup.printArea = 'A1:G20';
    worksheet.pageSetup.printTitlesRow = '1:3';

    let headRowArray = [
        ["Village Development Plan - Statement 1"],
        ["Village: Bhaisbor", null, null, null, null, "Gram Panchayat: Bhaisbor GP"],
        ["Statement 1: Summary of all Areas of Work"],
        [1, 2, 3, 4, 5, 6, 7, 8, 9],
        ["AOW ID", "Area Of Work (AOW)", "Sector", "Sub-Sector", "Funding\n(Scheme / Strategy)", "Line Department / Agency", "Budget\n(INR)", "Completion Time\n(In Months)", "Priority"],
        ["Group Name 1"],
        [1, "Activity Name 1", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"],
        [2, "Activity Name 2", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"],
        ["Group Name 2"],
        [3, "Activity Name 3", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"],
        [4, "Activity Name 4", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"]
    ];

    let rowsLength = headRowArray.length;
    let maxRowLength = 0;
    for (let i = 0; i < rowsLength; i++) {
        newLength = headRowArray[i].length;
        if (maxRowLength < newLength) {
            maxRowLength = newLength;
        }
        worksheet.addRow(headRowArray[i]);
    }

    if (maxRowLength === 9) {
        worksheet.mergeCells('A1', 'I1');
        worksheet.mergeCells('A2', 'E2');
        worksheet.mergeCells('F2', 'I2');
        worksheet.mergeCells('A3', 'I3');
    } else if (maxRowLength === 6) {
        worksheet.mergeCells('A1', 'F1');
        worksheet.mergeCells('A2', 'C2');
        worksheet.mergeCells('D2', 'F2');
        worksheet.mergeCells('A3', 'F3');
    }

    worksheet.getRow(1).font = {name: 'sans-serif', bold: true};
    worksheet.getRow(2).font = {name: 'sans-serif', bold: true};
    worksheet.getRow(3).font = {name: 'sans-serif', bold: true};

    let ArrayOfColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
    ArrayOfColumns.length = maxRowLength;
    headRowArray.push([]);

    headRowArray.forEach((row, index) => {
        ArrayOfColumns.forEach(element => {
            let cellName = element.toString() + index.toString();
            worksheet.getCell(cellName).alignment = {horizontal: 'center', wrapText: true, indent: 1, readingOrder: 'ltr'};
            worksheet.getCell(cellName).border = {
                top: {style: "medium"},
                left: {style: "medium"},
                bottom: {style: "medium"},
                right: {style: "medium"}
            };
        });
    });
    workbook.xlsx.writeFile('./sp.xlsx', 'utf-8').then((result) => {
        console.log('Result :-', result);
    });
}
let data = {};
exportXlsx(data);
