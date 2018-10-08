const path = require('path');
const XLSX = require('xlsx');

const subsectors = [{
    value: '0',
    viewValue: 'Afforestation'
},
{
    value: '1',
    viewValue: 'Communication'
},
{
    value: '3',
    viewValue: 'Horticulture'
},
{
    value: '4',
    viewValue: 'Public Transport'
},
{
    value: '5',
    viewValue: 'Village Institution'
},
{
    value: '6',
    viewValue: 'Water Security'
},
{
    value: '7',
    viewValue: 'Land Development'
},
{
    value: '8',
    viewValue: 'Rural Sanitation'
},
{
    value: '9',
    viewValue: 'Strengthening Democracy'
},
{
    value: '10',
    viewValue: 'Health'
},
{
    value: '11',
    viewValue: 'E-governance'
},
{
    value: '12',
    viewValue: 'Education'
},
{
    value: '13',
    viewValue: 'Vulnerable Section'
},
{
    value: '14',
    viewValue: 'Food Security'
},
{
    value: '15',
    viewValue: 'Well Being'
}
];

const priorities = [{
    value: '0',
    viewValue: 'Low'
},
{
    value: '1',
    viewValue: 'Medium'
},
{
    value: '2',
    viewValue: 'High'
}
];

const sectors = [{
    value: '0',
    viewValue: 'Ecology & Environment Development'
},
{
    value: '1',
    viewValue: 'Basic Amenities'
},
{
    value: '3',
    viewValue: 'Economic Development'
},
{
    value: '4',
    viewValue: 'Infrastructure'
},
{
    value: '5',
    viewValue: 'Social Development'
},
{
    value: '6',
    viewValue: 'Governance'
},
{
    value: '7',
    viewValue: 'Human Development'
}
];

const departments = [{
    value: '0',
    viewValue: 'Forest Environment and Wildlife Dept.'
},
{
    value: '1',
    viewValue: 'BSNL'
},
{
    value: '3',
    viewValue: 'HCCDD'
},
{
    value: '4',
    viewValue: 'Road & Bridges Deptt.'
},
{
    value: '5',
    viewValue: 'State RMDD'
},
{
    value: '6',
    viewValue: 'State Health Dept.'
},
{
    value: '7',
    viewValue: 'State HRDD'
}
];


function getWorkBook() {

    // exports.exportXlsx = function (inputObject, cb) {

    if (inputObject.statementType == 1) {

        const filename = (__dirname + '/st1.xlsx');
        let workbook = XLSX.readFile(filename, {
            cellStyles: true, cellNF: true,
        });

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
            i++;
            element.activities.forEach((activity, index) => {
                let actArr = [];
                actArr[0] = index + 1;
                actArr[1] = activity.name || null;
                actArr[2] = sectors.filter(s => s.value === activity.sector)[0] ? sectors.filter(s => s.value === activity.sector)[0].viewValue : null || null;
                actArr[3] = subsectors.filter(s => s.value === activity.subSector)[0] ? subsectors.filter(s => s.value === activity.subSector)[0].viewValue : null || null;
                if (activity.scheme && activity.anyOtherSource) {
                    actArr[4] = activity.scheme.schemeName || null;
                } else if (activity.scheme) {
                    actArr[4] = activity.scheme.schemeName || null;
                } else if (activity.anyOtherSource) {
                    actArr[4] = activity.anyOtherSource || null;
                } else {
                    actArr[4] = '--';
                }
                actArr[5] = departments.filter(d => d.value === activity.department)[0] ? departments.filter(d => d.value === activity.department)[0].viewValue : null || null;
                actArr[6] = activity.totalBudget || null;
                actArr[7] = activity.completionTime || null;
                actArr[8] = priorities.filter(p => p.value === activity.priority)[0] ? priorities.filter(p => p.value === activity.priority)[0].viewValue : null || null;
                headRowArray[i] = actArr;
                i++;
            });
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        XLSX.utils.sheet_add_aoa(sheet, headRowArray);
        workbook.Workbook.Sheets[0].Hidden = 1;

        // output format determined by filename *
        XLSX.writeFile(workbook, 'out.xlsx', {
            type: 'utf-8',
            cellStyles: true
        });


    } else if (inputObject.statementType == 2) {
        const filename = (__dirname + '/st2.xlsx');
        let workbook = XLSX.readFile(filename, {
            cellStyles: true, cellNF: true,
        });

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


        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        XLSX.utils.sheet_add_aoa(sheet, headRowArray);
        workbook.Workbook.Sheets[0].Hidden = 1;

        // output format determined by filename *
        XLSX.writeFile(workbook, 'out.xlsx', {
            type: 'utf-8',
            cellStyles: true
        });

        // workbook.xlsx.writeFile(urlForExcel, 'utf-8').then(result => {
        //     // cb(urlForExcel);
        // });

    }

}



const inputObject = {
    "statementType": 1,
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
                        "goals": [],
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
                        "goals": [],
                        "tags": [],
                        "endDateActivity": "1969-12-31T18:30:00.000Z",
                        "startDateActivity": "2099-12-31T18:30:00.000Z"
                    }
                ]
            },
            {
                "group": "Group 2",
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
                        "goals": [],
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

getWorkBook(inputObject);
