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

    let headRowArray = [[null, "Village Development Plan - Statement 1"], [null, "Village: Bhaisbor", null, null, null, null, "Gram Panchayat: Bhaisbor GP"], [null, "Statement 1: Summary of all Areas of Work"], [null, 1, 2, 3, 4, 5, 6, 7, 8, 9], [null, "AOW ID", "Area Of Work (AOW)", "Sector", "Sub-Sector", "Funding\n(Scheme / Strategy)", "Line Department / Agency", "Budget\n(INR)", "Completion Time\n(In Months)", "Priority"], [null, "Group Name 1"], [null, 1, "Activity Name 1", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"], [null, 2, "Activity Name 2", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"], [null, "Group Name 2"], [null, 3, "Activity Name 3", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"], [null, 4, "Activity Name 4", "Sector", "Sub-Sector", "Scheme Short Name", "Department", "Total Activity Budget", "Completion Time", "High, Med, Low"]];
    // let headRowArray = [{value: 'Village Development Plan - Statement 1'},
    // {villageName: 'Bhaisbor', gramPanchayatName: 'Gram Panchayat: Bhaisbor GP'},
    // {summary: 'Statement 1: Summary of all Area of Work'}];

    for (let i = 0; i < headRowArray.length; i++) {
        // let rowValues = [];
        // if (i === 0) {
        //     rowValues[0] = headRowArray[0].value;
        //     console.log(rowValues);
        worksheet.addRow(headRowArray[i]);
        // } else if (i === 1) {
        //     rowValues[0] = headRowArray[1].villageName;
        //     rowValues[5] = headRowArray[1].gramPanchayatName;
        //     console.log(rowValues);
        //     worksheet.addRow(rowValues);
        // } else if (i === 2) {
        //     rowValues[0] = headRowArray[2].summary;
        //     console.log(rowValues);
        //     worksheet.addRow(rowValues);
        // }
    }

    // worksheet.columns = [
    //     {header: 'AOW ID', key: 'aow_id', width: 15},
    //     {header: 'Area of Work (AOW)', key: 'work_id', width: 15},
    //     {header: 'Sector', key: 'sector_id', width: 15},
    //     {header: 'Sub-Sector', key: 'sub_sector_id', width: 15},
    //     {header: 'Funding (Scheme / Strategy)', key: 'funding_id', width: 15},
    //     {header: 'Line Department / Agency', key: 'department_id', width: 15},
    //     {header: 'Budget (INR)', key: 'budget_id', width: 15},
    //     {header: 'Completion Time (in Months)', key: 'completion_id', width: 15},
    //     {header: 'Prority', key: 'priority_id', width: 15},
    // ];

    workbook.xlsx.writeFile('./sp.xlsx', 'utf-8').then((result) => {
        console.log('Result :-', result);
    });

}
let data = {};
exportXlsx(data);
