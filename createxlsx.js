const Excel = require('exceljs');

let objectB = {
    "statementType": "2",
    "data": {
        "_id": "5ba48ac656aed728332f2914",
        "__v": 0,
        "name": "Plan",
        "year": "1",
        "status": "2",
        "vision": {
            "_id": "5ba35aa28e1d663240d21e91",
            "name": "Vision ",
            "color": "#3B9C9C",
            "description": "Desc",
            "totalGoalWeightage": 20,
            "deleted": false,
            "goals": [
                "5ba35ab38e1d663240d21e93",
                "5ba48adc56aed728332f2916"
            ],
            "objectives": ""
        },
        "country": "India",
        "state": "ANDHRA PRADESH",
        "district": "CHITTOOR",
        "tehsil": "Bangarupalem",
        "village": "Bodabafdla",
        "owner": "owner",
        "description": "desc",
        "nextActivityId": 1,
        "groups": [
            {
                "group": "group",
                "_id": "5ba48b0b56aed728332f2917",
                "activities": [
                    {
                        "_id": "5ba48b0b56aed728332f2918",
                        "plan": "5ba48ac656aed728332f2914",
                        "name": "act",
                        "group": "5ba48b0b56aed728332f2917",
                        "__v": 0,
                        "priority": "2",
                        "department": "1",
                        "subSector": "1",
                        "sector": "1",
                        "completionTime": "10 Months 3 Days",
                        "notes": "<p>notes act</p>",
                        "scheme": {
                            "_id": "5b1a0c58513f32224e49a242",
                            "schemeName": "Scheme"
                        },
                        "sequenceId": 1,
                        "totalBudget": 100,
                        "kpis": [
                            "5b9f58742cabc519984380c3"
                        ],
                        "milestones": [
                            {
                                "_id": "5ba48b7d56aed728332f2919",
                                "__v": 0,
                                "activity": "5ba48b0b56aed728332f2918",
                                "name": "Milestone",
                                "responsibility": "resp",
                                "startDate": "2018-09-16T18:30:00.000Z",
                                "outcome": "outcome",
                                "duration": 10,
                                "endDate": "2019-07-16T18:30:00.000Z",
                                "durationType": "Months",
                                "deleted": false,
                                "budget": 100
                            },
                            {
                                "_id": "5ba48b7d56aed728332f2919",
                                "__v": 0,
                                "activity": "5ba48b0b56aed728332f2918",
                                "name": "Milestone",
                                "responsibility": "resp",
                                "startDate": "2018-09-16T18:30:00.000Z",
                                "outcome": "outcome",
                                "duration": 10,
                                "endDate": "2019-07-16T18:30:00.000Z",
                                "durationType": "Months",
                                "deleted": false,
                                "budget": 100
                            }
                        ],
                        "goals": [],
                        "tags": [],
                        "endDateActivity": "2019-07-16T18:30:00.000Z",
                        "startDateActivity": "2018-09-16T18:30:00.000Z"
                    }
                ]
            }
        ],
        "deleted": false
    },
    "status": 1
}

let objectA = {
    "statementType": "1",
    "data": {
        "_id": "5b6bc146b3c1ef11b6457be5",
        "name": "devplannnnn",
        "status": "2",
        "year": "1",
        "projectName": "2",
        "country": "India",
        "state": "ANDAMAN AND NICOBAR ISLANDS",
        "district": "NORTH AND MIDDLE ANDAMAN",
        "tehsil": "Mayabunder",
        "village": "Asha Nagar (EFA)",
        "nextActivityId": 15,
        "groups": [
            {
                "group": "Group",
                "_id": "5b6bc175b3c1ef11b6457be6",
                "activities": [
                    {
                        "_id": "5b9b64b2ed13452dd1ecf8e3",
                        "plan": "5b6bc146b3c1ef11b6457be5",
                        "name": "Act for delete",
                        "group": "5b6bc175b3c1ef11b6457be6",
                        "__v": 0,
                        "sequenceId": 12,
                        "totalBudget": 0,
                        "kpis": [],
                        "scheme": {
                            "_id": '',
                            "schemeName": "Scheme 1"
                        },
                        "milestones": [],
                        "goals": [],
                        "tags": [],
                        "endDateActivity": "1969-12-31T18:30:00.000Z",
                        "startDateActivity": "2099-12-31T18:30:00.000Z"
                    },
                    {
                        "_id": "5b9b75d37082d7409cfce216",
                        "plan": "5b6bc146b3c1ef11b6457be5",
                        "name": "Act for delete 2",
                        "group": "5b6bc175b3c1ef11b6457be6",
                        "__v": 0,
                        "sequenceId": 13,
                        "totalBudget": 0,
                        "kpis": [],
                        "milestones": [],
                        "goals": [], "scheme": {
                            "_id": '',
                            "schemeName": "Scheme 1"
                        },
                        "tags": [],
                        "endDateActivity": "1969-12-31T18:30:00.000Z",
                        "startDateActivity": "2099-12-31T18:30:00.000Z"
                    },
                    {
                        "_id": "5b9f57fd2cabc519984380c1",
                        "plan": "5b6bc146b3c1ef11b6457be5",
                        "name": "act statement",
                        "group": "5b6bc175b3c1ef11b6457be6",
                        "__v": 0,
                        "sequenceId": 14,
                        "totalBudget": 0,
                        "kpis": [],
                        "milestones": [],
                        "goals": [],
                        "tags": [],
                        "scheme": {
                            "_id": '',
                            "schemeName": "Scheme 1"
                        },
                        "endDateActivity": "1969-12-31T18:30:00.000Z",
                        "startDateActivity": "2099-12-31T18:30:00.000Z"
                    },
                    {
                        "_id": "5b9f587f2cabc519984380c4",
                        "plan": "5b6bc146b3c1ef11b6457be5",
                        "name": "act statement",
                        "group": "5b6bc175b3c1ef11b6457be6",
                        "__v": 0,
                        "sequenceId": 15,
                        "totalBudget": 0,
                        "kpis": [
                            "5b9f58742cabc519984380c3"
                        ],
                        "milestones": [],
                        "scheme": {
                            "_id": '',
                            "schemeName": "Scheme 1"
                        },
                        "goals": [],
                        "tags": [],
                        "endDateActivity": "1969-12-31T18:30:00.000Z",
                        "startDateActivity": "2099-12-31T18:30:00.000Z"
                    }
                ]
            }
        ],
        "deleted": false
    },
    "status": 1
};

function exportXlsx (inputObject) {
    // inputObject = JSON.parse(inputObject);
    // let sheetName = data.sheetName || 'statement';
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

    var worksheet = workbook.addWorksheet('statement', {
        pageSetup: {paperSize: 9, orientation: 'landscape'}
    });
    worksheet.pageSetup.margins = {
        left: 0.7, right: 0.7,
        top: 0.75, bottom: 0.75,
        header: 0.3, footer: 0.3
    };
    worksheet.pageSetup.printArea = 'A1:G20';
    worksheet.pageSetup.printTitlesRow = '1:3';

    if (inputObject.statementType == 1) {
        let headRowArray = [];
        headRowArray[0] = ['Plan Name: ' + inputObject.data.name];
        let row2 = [];
        row2[0] = 'Village : ' + inputObject.data.village;
        if (inputObject.data.gramPanchayat) {
            row2[5] = 'GramPanchayat : ' + inputObject.data.gramPanchayat;
        } else {
            row2[5] = 'GramPanchayat : ' + '-----';
        }
        headRowArray[1] = row2;
        headRowArray[2] = ["Statement 1: Summary of all Areas of Work"];
        headRowArray[3] = [1, 2, 3, 4, 5, 6, 7, 8, 9];
        headRowArray[4] = ["AOW ID", "Area Of Work (AOW)", "Sector", "Sub-Sector", "Funding\n(Scheme / Strategy)", "Line Department / Agency", "Budget\n(INR)", "Completion Time\n(In Months)", "Priority"];

        let i = 5;
        inputObject.data.groups.forEach(element => {
            let arr = [];
            arr[0] = element.group;
            headRowArray[i] = arr;
            console.log('HD ', i, headRowArray[i]);
            i++;
            // headRowArray[i] = [];
            element.activities.forEach((activity, index) => {
                let actArr = [];
                actArr[0] = index + 1;
                actArr[1] = activity.name || null;
                actArr[2] = activity.sector || null;
                actArr[3] = activity.subSector || null;
                actArr[4] = activity.scheme.schemeName || null;
                actArr[5] = activity.department || null;
                actArr[6] = activity.totalBudget || null;
                actArr[7] = activity.completionTime || null;
                actArr[8] = activity.priority || null;
                // console.log('ActArray :', actArr);
                headRowArray[i] = actArr;
                i++;
            });
        });
        // console.log('Head Array :- ', JSON.stringify(headRowArray));

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

        worksheet.getRow(1).font = {name: 'sans-serif', bold: true, size: 18};
        worksheet.getRow(2).font = {name: 'sans-serif', bold: true, size: 14};
        worksheet.getRow(3).font = {name: 'sans-serif', bold: false, size: 16};

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
    } else if (inputObject.statementType == 2) {
        let headRowArray = [];
        headRowArray[0] = ['Plan Name: ' + inputObject.data.name];
        let row2 = [];
        row2[0] = 'Village : ' + inputObject.data.village;
        if (inputObject.data.gramPanchayat) {
            row2[3] = 'GramPanchayat : ' + inputObject.data.gramPanchayat;
        } else {
            row2[3] = 'GramPanchayat : ' + '-----';
        }
        headRowArray[1] = row2;
        headRowArray[2] = ["Statement 2: : Details of each Area of Work"];
        headRowArray[3] = [1, 2, 3, 4, 5, 6];
        headRowArray[4] = ["AOW ID", "Area Of Work(AOW)", "Milestones", "Milesstone Timeline", "Responsibility", "Outcome"];

        let i = 5;
        inputObject.data.groups.forEach(element => {
            let arr = [];
            arr[0] = element.group;
            headRowArray[i] = arr;
            i++;
            element.activities.forEach((activity, index) => {
                let actName = activity.name;
                activity.milestones.forEach(miles => {
                    let actArr = [];
                    actArr[0] = index + 1;
                    actArr[1] = actName || null
                    actArr[2] = miles.name || null;
                    actArr[3] = miles.endDate || null;
                    actArr[4] = miles.responsibility || null;
                    actArr[5] = miles.outcome || null;
                    headRowArray[i] = actArr;
                    i++;
                });
            });
        });
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

        worksheet.getRow(1).font = {name: 'sans-serif', bold: true, size: 18};
        worksheet.getRow(2).font = {name: 'sans-serif', bold: true, size: 14};
        worksheet.getRow(3).font = {name: 'sans-serif', bold: false, size: 16};

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


}

exportXlsx(objectB);
