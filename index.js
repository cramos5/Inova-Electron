/**
 * Created by Carlos Ramos on 5/23/2017.
 */
const electron = require('electron');
const { app, BrowserWindow, Menu, ipcMain } = electron;
const path = require('path');

const fs = require('fs');

let Excel;
let workbook;
var sheet;
let id;
let tempTime;
let tempDate;
let fileName;
let format;
let formatUp;
let child;

Excel = require('exceljs');
workbook = new Excel.Workbook();
sheet = workbook.addWorksheet('My Sheet');




format = [
    {key:"clinic", header:"Clinic ID", width:10},
    {key:"date", header:"Date", width:10},
    {key:"fname", header:"First Name", width:10},
    {key: "lname", header: "Last Name", width:10},
    {key: "id", header: "Employee ID", width:10},
    {key: "dob", header: "Date of Birth", width:10},
    {key: "height", header: "Height", width:10},
    {key: "ibs", header: "Weight(Ibs)", width:10},
    {key: "waist", header: "Waist Measurement", width:10},
    {key: "bp", header: "Blood Pressure", width:10},
    {key: "bcomp", header: "Body Composition", width:10},
    {key: "bmi", header: "Body Mass Index", width:10},
    {key: "preg", header: "Pregnant", width:10},
    {key: "fast", header: "Fasting", width:10},
    {key: "pace", header: "Pacemaker", width:10}
];

formatUp = [
    {key:"clinic", header:"Clinic ID", width:10},
    {key:"date", header:"Date", width:10},
    {key:"fname", header:"First Name", width:10},
    {key: "lname", header: "Last Name", width:10},
    {key: "id", header: "Employee ID", width:10},
    {key: "dob", header: "Date of Birth", width:10},
    {key: "height", header: "Height", width:10},
    {key: "ibs", header: "Weight(Ibs)", width:10},
    {key: "waist", header: "Waist Measurement", width:10},
    {key: "bp", header: "Blood Pressure", width:10},
    {key: "bcomp", header: "Body Composition", width:10},
    {key: "bmi", header: "Body Mass Index", width:10},
    {key: "preg", header: "Pregnant", width:10},
    {key: "fast", header: "Fasting", width:10},
    {key: "pace", header: "Pacemaker", width:10},
    {key: "upheight", header: "Ht (cm) - Map to HEIGHT", width:10},
    {key: "upwaist", header: "Wt (kg) - Map to Weight", width:10},
    {key: "upweight", header: "Waist (cm) - Map to WAIST", width:10}
];


//Header must be in below format
sheet.columns = format;
sheet.getRow(1).font = {bold: true};
let mainWindow;
let clinicWindow;
let addWindow;

tempTime = getTime();
tempDate = getDate();


function entryCheck(bio) {
    var i;
    var mes = 'full';
    for (i = 0; i < bio.length; i++ ) {
        if(i == 5 ){
            continue;
        }

        if(bio[i] == ""){
            mes = 'empty';
        }
    };
    return mes;
}

function message(msg) {
    child = new BrowserWindow(
        {parent: mainWindow, modal : true, show: false, height: 270,
        width: 600,
        resizeable: false,
        frame: false,});
    child.loadURL('file://'+__dirname+'/message.html');
    child.once('ready-to-show', () => {
        child.webContents.send('sendmsg', msg);
        child.show();
    });

    child.on('close', () => {
        child = null;
    })

}



function getDate() {

    var date = new Date();

    var year = date.getFullYear();

    var month = date.getMonth() + 1;
    month = (month < 10 ? "0" : "") + month;

    var day  = date.getDate();
    day = (day < 10 ? "0" : "") + day;

    return month + '-' + day + '-' + year;
}

function getTime() {

    var date = new Date();

    var hour = date.getHours();
    hour = (hour < 10 ? "0" : "") + hour;

    var min  = date.getMinutes();
    min = (min < 10 ? "0" : "") + min;

    var sec  = date.getSeconds();
    sec = (sec < 10 ? "0" : "") + sec;

    return hour + '.' + min + '.' + sec;
}

function finishXL(){

};

function createMaster(){
    let masterbook = new Excel.Workbook();
    let masterfile = "./Excel_Files/Master_Lists/Master List on " + tempDate+ " at " + tempTime +".xlsx";
    let msheet = masterbook.addWorksheet('All Clinics');
    msheet.columns = formatUp;
    msheet.getRow(1).font = {bold: true};
    msheet.getRow(1).alignment = {textRotation: 90};
    msheet.getRow(1).height = 200;
    fs.readdir('./Excel_Files/Individual_Clinics/', (err, dir) => {
        for (var i = 0, path; path = dir[i]; i++) {
            //console.log('./Excel_Files/Individual_Clinics/'+path);
            let loc = path;
            let tempwork = new Excel.Workbook();
            tempwork.xlsx.readFile('./Excel_Files/Individual_Clinics/'+loc)
                .then(function() {
                    //console.log("getting data from each excel, opened" + loc);
                    var tempsheet = tempwork.getWorksheet("My Sheet");
                    var x = 0;
                    tempsheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
                        //console.log("Row " + rowNumber);
                        if(rowNumber != 1){
                            var mdata = {
                                clinic: Number(row.values[1]),
                                date: row.values[2],
                                fname: row.values[3],
                                lname: row.values[4],
                                id: Number(row.values[5]),
                                dob: row.values[6],
                                height: Number(row.values[7]),
                                ibs: Number(row.values[8]),
                                waist: Number(row.values[9]),
                                bp:row.values[10],
                                bcomp: Number(row.values[11]),
                                bmi: Number(row.values[12]),
                                preg: row.values[13],
                                fast: row.values[14],
                                pace: row.values[15],
                                upheight: Number((Number(row.values[7]) * 2.54).toFixed(2)),
                                upwaist: Number((Number(row.values[9]) * 2.54).toFixed(2)),
                                upweight: Number((Number(row.values[8])/2.2).toFixed(2)),
                            };
                           // console.log(mdata);
                            msheet.addRow(mdata);
                            masterbook.xlsx.writeFile(masterfile).then(function() {});
                        }

                        //console.log("tryin to read row");
                    });
                   // console.log("ENd of Sheet")
                })
        }
    });
    message("Master Excel File Created")
}

app.on('ready', () =>{
    mainWindow = new BrowserWindow({
        height: 900,
        width: 900,
        minWidth: 900,
        minHeight: 900,
        resizeable: false,
        show: false,
        frame: false,
        icon: path.join(__dirname + '/app/photo.jpg')
    });
    mainWindow.loadURL('file://'+__dirname+'/Layout.html');
    mainWindow.on('closed', () =>{
        app.quit(); //Lets us close all windows once main is closed
    } );
    const mainMenu = Menu.buildFromTemplate(menuTemplate)//Sets up Menu
    Menu.setApplicationMenu(mainMenu)//Implements Menu

    clinicWindow = new BrowserWindow({
        height: 270,
        width: 600,
        resizeable: false,
        show: true,
        frame: false,
        icon: path.join(__dirname + '/app/photo.jpg')
    });
    clinicWindow.loadURL('file://'+__dirname+'/clinicid.html');
    clinicWindow.on('close', () => {
        clinicWindow = null;
    }); //This allows garbage collecter to free memory, since referance is
    //maintained until reassinged
});

ipcMain.on('IDdata', (event, data) => {
    id = data;
    fileName = "./Excel_Files/Individual_Clinics/Clinic " + id + " on " + tempDate+ " at " + tempTime +".xlsx";
    clinicWindow.close();
    mainWindow.webContents.send('sendID', id);
    mainWindow.show()
});

ipcMain.on('killmsg', (event) => {
    child.close();
});

ipcMain.on('createmaster', (event) => {
    createMaster();
   // message("Master Excel Sheet Created");
});

ipcMain.on('killapp', (event) => {
    clinicWindow.close();
    app.quit();
});



ipcMain.on('data', (event, biometrics) => {



    var empty = entryCheck(biometrics);

    if (empty == 'empty') {
        message("Please Enter Values for All Entries")
        return;
    };

    var data = {
        clinic: id,
        date: tempDate,
        fname: biometrics[0],
        lname: biometrics[1],
        id: biometrics[2],
        dob: biometrics[3],
        height: Number(biometrics[4])*12 + Number(biometrics[5]),
        ibs: biometrics[6],
        waist: biometrics[7],
        bp: biometrics[8]+"/" + biometrics[9],
        bcomp: biometrics[10],
        bmi: biometrics[11],
        preg: biometrics[12],
        fast: biometrics[13],
        pace: biometrics[14]
    };
    sheet.addRow(data);
    workbook.xlsx.writeFile(fileName).then(function() {});
    mainWindow.webContents.send('clearform');
});

ipcMain.on('quitmain', (event) => {
    app.quit();
});



const menuTemplate = [
    {
        label: 'File',
        submenu: [

            {
                label:'Quit',

                accelerator: process.platform === 'darwin' ? 'Command+Q' : 'Ctrl+Q', // conddtion ? True : False

                click() {
                    app.quit();
                }
            }
        ]

    } //Each {} Object Corresponds to the label of drop down sub menue
]

if (process.platform === 'darwin') {
    menuTemplate.unshift({}); //This if statement handles the menu differences with OIX
}

if(process.env.NODE_ENV !== 'production') {
    menuTemplate.push({//Add menu option to left
        label: 'Debug',
        submenu: [
            {//Shortcut Method to reinstate default menu options
                role: 'reload' //Shortcut for reloading page
            },
            {
                label: 'Toggle Developer Tools',

                accelerator:  process.platform === 'darwin' ? 'Command+I' : 'Ctrl+Shift+I',

                click(item, focusedWindow){
                    focusedWindow.toggleDevTools();
                }
            }
        ]

    });
    //production
    //development
    //staging
    //test
}