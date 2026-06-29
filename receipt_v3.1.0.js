// JavaScript source code
//新レシート表ver.3.1.0
//assembled on 2026/6/8    last update 2026/6/13
//assembled by youichi ohno

var ss = SpreadsheetApp.getActiveSpreadsheet();
let activeSheet;

function sheetCheck() {
    //初めにシートの存在チェック
    try {
        const sheetname1 = "シート1";
        const sheetname2 = "057200前熊センター";
        const sheetname3 = "057201砂子センター";
        const sheetname4 = "057419長久手営業所";

        let sheet = [];
        sheet[0] = ss.getSheetByName(sheetname1);
        sheet[1] = ss.getSheetByName(sheetname2);
        sheet[2] = ss.getSheetByName(sheetname3);
        ss.setActiveSheet(sheet[0]);
        ss.setActiveSheet(sheet[1]);
        ss.setActiveSheet(sheet[2]);

    } catch (e) {
        Logger.log("シートがありません :" + e.message);
    }
}

//**************************************** */
//ボタンクリックしたときに各ステータスを読み込む
//**************************************** */
function makepdfButtunClick() {
    const sheetname = "シート1";

    try {
        activeSheet = ss.getSheetByName(sheetname);
        ss.setActiveSheet(activeSheet);
    } catch (e) {
        Logger.log("シートがありません :" + e.message);
    }
    activeSheet.getRange(1, 3, 2, 10).merge().clearContent();
    let values = activeSheet.getRange('a1:a5').getValues();

    //値のチェック
    let check0 = Object.prototype.toString.call(new Date(values[0]));
    let check1 = Object.prototype.toString.call(new Date(values[1]));
    let check2 = Number.isInteger(Number(values[2]));
    let check3 = typeof Boolean(values[3]);
    let check4 = typeof String(values[4]);

    if (check0 == "[object Date]" && check1 == "[object Date]" && check2 == true && check3 == "boolean" && check4 == "string") {
        if (values[0] != "" && values[1] != "" && values[3] != "" && values[4] != "") {
            if (values[4] == "057200前熊センター" || values[4] == "057201砂子センター" || values[4] == "057419長久手営業所") {
                let check11 = islatterdaylater(check0, check1);
                if (Boolean(check11 === false)) {
                    //処理本体に引き渡す
                    receiptSheetMain(values);
                } else {
                    Logger.log("日付が逆転しています");
                }
            } else {
                Logger.log("値が異常です");
            }
        } else {
            Logger.log("値が異常です");
        }
    }
}

//************************************* */
//処理本体
//************************************* */
function receiptSheetMain(argEnterValues) {
    let begindate, enddate, dateinvalid, print, center, i = 0;
    let enterValues = [];
    argEnterValues.map(function (e) {
        [[begindate], [enddate], [print], [dateinvalid], [center]] = argEnterValues;

        if (enterValues[i] == "") {
            Logger.log("値が異常です");
        }
        [begindate, enddate, print, dateinvalid, center].flat(1);
    });

    //引数に従ってスプレッドシートから順次名前を読み込む
    let sheetname = center;

    let sheet;
    try {
        sheet = ss.getSheetByName(sheetname);
        ss.setActiveSheet(sheet);
    } catch (e) {
        Logger.log("シートがありません :" + e.message);
    }

    //********************************* */
    //スプレッドシートから各値を取得
    //********************************* */
    const firstRow = 3;
    const employeeRow = 1;
    const colLength = 5;
    var rowLength;
    var employees;

    if (sheet != null) {
        rowLength = sheet.getLastRow() - (firstRow - 1);

        //社員集合体
        const cache = makeUserCache();
        if (cache.get('employees') == null) {
            employees = sheet.getRange(firstRow, employeeRow, rowLength, colLength).getValues();
            cache.put('employees', employees);
        } else if (cache.get('employees') != null) {
            if ((cache.get('emploees').length > 0)) {
                employees = cache.get('employees');
            }
        }


        //********************************** */
        //日付の設定
        //********************************** */
        let nowDate = new Date();

        let folder;
        var folders = [];
        folder = makeFolder1();
        folders = makeFolder2(nowDate, folder);

        //receiptsheetフォルダIDを取得
        let folderId = DriveApp.getFoldersByName('receiptsheet').next().getId();
        var folderId1 = folders[0].getId();
        var folderId2 = folders[1].getId();
        var folderId3 = folders[2].getId();

        let file = DriveApp.getFileById(ss.getId());

        let currentFolders = file.getParents();

        let moveFolder, folderId11;

        if (currentFolders.hasNext()) {
            moveFolder = currentFolders.next();
            folderId11 = moveFolder.getId();
        }

        eachValueProcess(begindate, enddate, print, dateinvalid, center, employees);
    } else {
        Logger.log("エラー");
    }

    function eachValueProcess(begindate, enddate, print, dateinvalid, center, employees2) {
        let beginDate1 = new Date(begindate);
        let endDate2 = new Date(enddate);
        let showDate = Utilities.formatDate(endDate2, 'JST', 'yyyy/M/d');
        let beginDate = Utilities.formatDate(beginDate1, 'JST', 'yyyy年M月d日');
        let endDate = Utilities.formatDate(endDate2, 'JST', 'yyyy年M月d日');
        let loopp = 3;

        let statusvalues2 = [beginDate, endDate, print, dateinvalid, center, loopp, showDate, beginDate1, endDate2];

        //***************************************************** */
        //印刷するスプレッドシートページ数を数える
        //***************************************************** */
        let employeePage = valueAnalysis1(employees2);

        //************************************* */
        //印刷する日にちを数える
        //************************************* */
        let dateNumber = isdatediffdate(begindate, enddate);
        //結果スプレッドシートファイルの命名
        let fileNaming1 = [], fileNaming2 = [];

        //********************************* */
        //ファイル命名
        //********************************* */
        for (let i = 0; i <= dateNumber; i++) {
            if (i == 0) {
                fileNaming1.push(Utilities.formatDate(new Date(begindate), 'JST', 'yyyyMMdd'));
            } else if (i > 0) {
                begindate.setDate(begindate.getDate() + 1);
                fileNaming1.push(Utilities.formatDate(begindate, 'JST', 'yyyyMMdd'));
            }

            for (let j = 0; j < employeePage; j++) {
                fileNaming2.push(String(fileNaming1[i]) + String(j));
            }

        }

        //****************************** */
        //配列から必要社員名簿を抽出
        //****************************** */

        let associationEmployees3;
        let filesId = [];
        filesId = makeSpreadSheet(folderId1, fileNaming2);

        //************************** */
        //各関数の呼び出し
        //************************** */
        associationEmployees3 = getEmployeesAssociationValues(employees);
        drawBody(associationEmployees3, statusvalues2, filesId, folders);
    }

    //******************************************* */
    //******************************************* */
    //ここから各関数
    //******************************************* */
    //******************************************* */

    //***************************************** */
    //***************************************** */
    //スプレッドシート組み立て
    //スプレッドシート作成
    //***************************************** */
    //***************************************** */  
    function makeSpreadSheet(folderId, fileNames = []) {
        let sheet, file = [], files = [], fileId, fileId11 = [], filesId2 = [], sheetId = [], sheetID;

        fileNames.map(function (e) {
            try {

                let fileMetadata = {
                    name: e,
                    mimeType: 'application/vnd.google-apps.spreadsheet',
                    parents: [folderId]
                };

                file[i] = Drive.Files.create(fileMetadata);

                sheet = SpreadsheetApp.openById(file[i].id);
                fileId = sheet.getSheets()[0];
                sheetID = sheet.getSheetId();

                fileId11.push(file[i].getId());
                filesId2.push(file[i].getId());
                sheetId.push(sheetID);

                files.push(file[i])

            } catch (e) {
                e.message;
            }
        });

        return [filesId2, fileId11, fileNames, files, sheetId];
    }

    //****************************************** */
    //連想配列にする
    //****************************************** */
    function getEmployeesAssociationValues(employeesValues = []) {
        //値の配列を作成
        let keys = ["employeeNo", "employeeName", "groupChief", "receipt", "order"];
        let employees10 = [];

        employees10 = employeesValues.map(function (values) {
            let obj = {};
            keys.forEach(function (key, i) {
                obj[key] = values[i];
            });

            return obj;
        });

        return employees10;

    }

    //****************************** */
    //各フォルダ作成
    //****************************** */
    function makeFolder1() {
        //スプレッドシートが格納されているフォルダ自身を取得
        let ssId = ss.getId();
        let parentFolder, folder1;

        try {
            parentFolder = DriveApp.getFileById(ssId).getParents();
            folder1 = parentFolder.next();
        } catch (e) {
            e.message;
        }

        let folders, folder;
        //receiptsheetフォルダを作成
        try {
            folders = folder1.getFoldersByName('receiptsheet');
            folder = folders.next();
        } catch (error) {
            folder = folder1.createFolder('receiptsheet');
        }

        return folder;
    }

    function makeFolder2(nowDate, folder) {
        let folders3, folders4, folders5;
        let folder4, folder5, folders = [];

        //スプレッドシート・pdfファイル保存用フォルダ名指定
        let nowDate6 = Utilities.formatDate(nowDate, 'JST', 'yyyyMMddHHmmss');
        let folder3;
        //receiptsheet以下スプレッドシート・pdfファイル保存用・結果pdfファイル保存フォルダを作成
        try {
            folders3 = DriveApp.getFoldersByName(nowDate6);
            folder3 = folders3.next();
        } catch (error) {
            folder3 = folder.createFolder(nowDate6);
        }
        try {
            folders4 = folder3.getFoldersByName('pdf');
            folder4 = folders4.next();
        } catch (error) {
            folder4 = folder3.createFolder('pdf');
        }
        try {
            folders5 = folder4.getFoldersByName('resultpdf');
            folder5 = folders5.next();
        } catch (error) {
            folder5 = folder4.createFolder('resultpdf');
        }

        folders[0] = folder3;
        folders[1] = folder4;
        folders[2] = folder5;

        return folders;
    }

}

//******************************************** */
//******************************************** */
//罫線・各値一括設定
//******************************************** */
//******************************************** */
function drawBody(writeValues = [], printValues = [], filesId = [], folders = []) {
    let fileId;
    let writeValuesLength = writeValues.length;

    //*************************************** */
    //昇順にソート
    //*************************************** */
    let newWriteValues3 = writeValues.sort(descOrder);
    let newWriteValues = basicSort(writeValues);

    //************************************** */
    //ここから各ソート
    //************************************** */
    function basicSort(writeValues = []) {
        let newwritevalues = [], newwritevalues0 = [], newwritevalues1 = [];
        let val11, val2, val3, val4, val5, val6, i = 0;
        let newVal2 = "", newVal3 = "", newVal4 = "", newVal5 = "", newVal6 = "";
        let writeValuesLength = writeValues.length;
        writeValues.map(function (key, val) {
            val11 = JSON.parse(JSON.stringify(writeValues[val]));
            val2 = val11.employeeNo;
            val3 = val11.employeeName;
            val4 = val11.groupChief;
            val5 = val11.receipt;
            val6 = val11.order;

            if (val5 != "") {
                newwritevalues1.push([val2, val3, val4, val5, val6]);
                i++;
            } else if (val4 != "") {
                newVal2 = val2, newVal3 = val3, newVal4 = val4, newVal5 = val5, newVal6 = val6;
                i++;
            } else {
                newwritevalues0.push([val2, val3, val4, val5, val6]);
                i++;
            }
            if (writeValuesLength <= i) {
                if (newwritevalues0.length != 0) {
                    newwritevalues = newwritevalues0.concat(newwritevalues1);
                } else if (newwritevalues1.length != 0) {
                    newwritevalues = newwritevalues1;
                }
                if (newVal2 != "") {
                    newwritevalues.unshift([newVal2, newVal3, newVal4, newVal5, newVal6]);
                }
            }
        });
        return newwritevalues;
    }

    let receiptCheck = checkElement(newWriteValues);

    function checkElement(writeValues = []) {
        let judge;
        writeValues.some(function (val) {
            judge = val.includes("受付");
        });
        return judge;
    }

    let newWriteValues2 = receiptSort(newWriteValues, receiptCheck);
    writeValuesLength = writeValues.length;
    var page = Math.ceil(writeValuesLength / 9);
    function receiptSort(writeValues = [], receiptCheck) {
        let newwritevalues = [], newwritevalues0 = [], newwritevalues1 = [];
        let val11, val2, val3, val4, val5, val6;
        let newVal2 = "", newVal3 = "", newVal4 = "", newVal5 = "", newVal6 = "";
        writeValuesLength = writeValues.length;
        let i = 0, j = 1, k = 0;

        let page = Math.ceil(writeValuesLength / 9);
        let remainder = writeValuesLength % 9;
        let perfectNumber = page * 9 - 1;
        let diff = perfectNumber - writeValuesLength;
        writeValues.map(function (val) {
            val11 = JSON.parse(JSON.stringify(val));
            val2 = val11[0];//.employeeName;
            val3 = val11[1];//.employeeNo;
            val4 = val11[2];//.groupChief;
            val5 = val11[3];//.receipt;
            val6 = val11[4];//.order;

            if (writeValuesLength >= 9) {
                if (page > j) {
                    if (receiptCheck) {
                        if (val2 != "") {
                            newwritevalues0.push([val2, val3, val4, val5, val6]);
                            k++;
                        }
                    } else if (page >= j) {
                        if (val2 != "") {
                            newwritevalues0.push([val2, val3, val4, val5, val6]);
                            k++;
                        }
                    }
                    i++;

                } else if (page <= j) {
                    if (receiptCheck) {
                        if (val5 == "") {
                            newwritevalues0.push([val2, val3, val4, val5, val6]);
                            k++;
                        } else if (val5 != "" && page == j) {
                            if (remainder != 0) {
                                newVal2 = val2, newVal3 = val3, newVal4 = val4, newVal5 = val5, newVal6 = val6;
                                while (perfectNumber > k) {
                                    newwritevalues1.push([" ", " ", " ", " ", " "]);
                                    k++;
                                }
                                newwritevalues1.push([newVal2, newVal3, newVal4, newVal5, newVal6]);
                                k++;
                            }

                        }
                    } else if (page == j) {
                        if (receiptCheck) {
                            if (val2 != "") {
                                newwritevalues0.push([val2, val3, val4, val5, val6]);
                                k++;
                            }
                        } else if (page >= j) {
                            if (val2 != "") {
                                newwritevalues0.push([val2, val3, val4, val5, val6]);
                                k++;
                            }
                        }
                        if (val2 != "") {
                            newwritevalues0.push([val2, val3, val4, val5, val6]);
                            k++;
                        }
                    }
                    i++;

                }

                if (writeValuesLength <= k + 1 && !receiptCheck) {
                    while (perfectNumber > k) {
                        newwritevalues1.push([" ", " ", " ", " ", " "]);
                        k++;
                    }
                    i++;
                }
            } else if (writeValuesLength < 9) {
                newwritevalues0.push([val2, val3, val4, val5, val6]);
                k++;
                if (writeValuesLength <= k + 1) {
                    while (perfectNumber > k) {
                        newwritevalues1.push([" ", " ", " ", " ", " "]);
                        k++;
                    }
                }

                i++;
            }

            if (i > 8) {
                i = 0;
                j++;
            }
            newwritevalues = newwritevalues0.concat(newwritevalues1);
        });
        return newwritevalues;
    }

    let i = 0, j = 0, k = 0, val11, val2, val3, val4, val5, val6, writeText11 = [], writeText12 = [], writeText13 = [], writeText111, writeText121, writeText131;

    if (printValues[2] == 0) {
        const date1 = new Date(printValues[7]);
        const date2 = new Date(printValues[8]);
        const cache1 = makeUserCache();
        for (let date = date1; date <= date2; date.setDate(date.getDate() + 1)) {
            for (let val of newWriteValues2) {
                [employeeName, employeeNo, groupChief, receipt, order] = val;

                val2 = val[0];
                val3 = val[1];
                val4 = val[2];
                val5 = val[3];
                val6 = val[4];

                if (cache1.get('writeText11') == null || cache1.get('writeText12') == null || cache1.get('writeText13') == null) {
                    if (i < 3) {
                        writeText11.push(val2, val3);
                    } else if (i >= 3 && i < 6) {
                        writeText12.push(val2, val3);
                    } else if (i >= 6 && i < 9) {
                        writeText13.push(val2, val3);
                    }
                }

                i++;

                if (i > 8) {
                    i = 0;
                    k++;
                    fileId = filesId[0].shift();

                    writeText111 = 'writeText11' + String(j);
                    writeText121 = 'writeText12' + String(j);
                    writeText131 = 'writeText13' + String(j);

                    if (cache1.get(writeText111) != null && cache1.get(writeText121) != null && cache1.get(writeText131) != null) {
                        writeText11 = cache1.get(writeText111);
                        writeText12 = cache1.get(writeText121);
                        writeText13 = cache1.get(writeText131);
                    }
                    if (cache1.get('writeText11') == null) {
                        cache1.put(writeText111, writeText11);
                    }
                    if (cache1.get('writeText12') == null) {
                        cache1.put(writeText121, writeText12);
                    }
                    if (cache1.get('writeText13') == null) {
                        cache1.put(writeText131, writeText13);
                    }

                    draw(writeText11, writeText12, writeText13, printValues, fileId, k, date, filesId[1], filesId);
                    if (k >= page) {
                        k = 0;
                    }
                    writeText11 = [];
                    writeText12 = [];
                    writeText13 = [];
                    j++;
                }
            }
        }
        merge(folders[1].getId(), folders[2].getId());
    }

    function draw(writeText0 = [], writeText1 = [], writeText2 = [], printValues = [], fileId, print, date, fileId11 = [], filesId) {
        let ss1, ss2;
        ss1 = SpreadsheetApp.openById(fileId);
        ss2 = ss1.getActiveSheet();
        sheetId = ss2.getSheetId();
        //一括で書き込み
        try {
            ss2.getRange("c1:d3").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("20");
            ss2.getRange("e3:f3").mergeAcross().setHorizontalAlignment("right").setVerticalAlignment("middle");
            ss2.getRange("f1:f2").setHorizontalAlignment("right").setVerticalAlignment("middle");

            ss2.getRange("a4:f4").setHorizontalAlignments([["left", "center", "left", "center", "left", "center"]]).setVerticalAlignment("middle");
            ss2.getRange("a5:f5").setHorizontalAlignments([["right", "center", "right", "center", "right", "center"]]).setVerticalAlignment("middle").setFontSize("14");
            ss2.getRange("a7:b9").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("c7:d9").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("e7:f9").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("a11:b11").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("c11:d11").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("e11:f11").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("a13:b13").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("c13:d13").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("e13:f13").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");

            ss2.getRange("a16:f16").setHorizontalAlignments([["left", "center", "left", "center", "left", "center"]]).setVerticalAlignment("middle");
            ss2.getRange("a17:f17").setHorizontalAlignments([["right", "center", "right", "center", "right", "center"]]).setVerticalAlignment("middle").setFontSize("14");
            ss2.getRange("a19:b21").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("c19:d21").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("e19:f21").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("a23:b23").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("c23:d23").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("e23:f23").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("a25:b25").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("c25:d25").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("e25:f25").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");

            ss2.getRange("a28:f28").setHorizontalAlignments([["left", "center", "left", "center", "left", "center"]]).setVerticalAlignment("middle");
            ss2.getRange("a29:f29").setHorizontalAlignments([["right", "center", "right", "center", "right", "center"]]).setVerticalAlignment("middle").setFontSize("14");
            ss2.getRange("a31:b33").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("c31:d33").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("e31:f33").merge().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("36").setFontColor('magenta');
            ss2.getRange("a35:b35").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("c35:d35").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("e35:f35").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize("16");
            ss2.getRange("a37:b37").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("c37:d37").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");
            ss2.getRange("e37:f37").mergeAcross().setHorizontalAlignment("center").setVerticalAlignment("middle");


            ss2.getRangeList(["a5:f5"]).setBorder(true, false, true, false, null, null);
            ss2.getRangeList([["a4:a5"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["c4:c5"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["e4:e5"]]).setBorder(true, true, true, true, null, null);

            ss2.getRangeList(["a16:f16"]).setBorder(true, false, true, false, null, null);
            ss2.getRangeList(["a17:f17"]).setBorder(true, false, true, false, null, null);
            ss2.getRangeList([["a16:a17"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["c16:c17"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["e16:e17"]]).setBorder(true, true, true, true, null, null);

            ss2.getRangeList(["a28:f28"]).setBorder(true, false, true, false, null, null);
            ss2.getRangeList(["a29:f29"]).setBorder(true, false, true, false, null, null);
            ss2.getRangeList([["a28:a29"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["c28:c29"]]).setBorder(true, true, true, true, null, null);
            ss2.getRangeList([["e28:e29"]]).setBorder(true, true, true, true, null, null);

            ss2.getRangeList(["c4:d39"]).setBorder(true, true, true, true, null, null);

            ss2.getRangeList(["a4:f39"]).setBorder(true, true, true, true, null, null, "", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

            ss2.setColumnWidths(1, 6, 130);
            ss2.setRowHeights(1, 39, 28);

            SpreadsheetApp.flush();

        } catch (e) {
            e.message;
        }


        const personalNo = "個人No.";
        const personalName = "氏名";
        const holidaydeposit = "公休・誤入金";
        const attention = "↑レシートのない理由に○をする";

        const str1 = [personalNo, personalName, personalNo, personalName, personalNo, personalName];
        const str2 = [holidaydeposit, , holidaydeposit, , holidaydeposit, ,];
        const str3 = [attention, , attention, , attention, ,];

        write();

        function write() {
            let date11;
            if (printValues[3] === Boolean(true)) {
                date11 = "年      月      日";
            } else {
                //文字列型変換ここポイント！！
                date11 = String(Utilities.formatDate(new Date(date), 'JST', 'yyyy年M月d日'));
            }

            try {
                ss2.getRange(1, 1, 39, 6).setValues([["", "", "釣銭準備金持ち出し簿", "", "", date11,], ["", "", "", "", "", "No." + print,], ["", "", "", "", printValues[4], ""],
                    str1,
                    writeText0,
                ["", "", "", "", "", ""],
                [printValues[6], "", printValues[6], "", printValues[6], ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                    str2,
                ["", "", "", "", "", ""],
                    str3,
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                    str1,
                    writeText1,
                ["", "", "", "", "", ""],
                [printValues[6], "", printValues[6], "", printValues[6], ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                    str2,
                ["", "", "", "", "", ""],
                    str3,
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                    str1,
                    writeText2,
                ["", "", "", "", "", ""],
                [printValues[6], "", printValues[6], "", printValues[6], ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                    str2,
                ["", "", "", "", "", ""],
                    str3,
                ["", "", "", "", "", ""],
                ["", "", "", "", "", ""],
                ],
                );

                SpreadsheetApp.flush();
                let fileName = filesId[2].shift();
                let ssId = filesId[1].shift();
                let shId = filesId[4].shift();

                makepdf(folders[1].getId(), fileName, ssId, shId);

            } catch (e) {
                e.message;
            }

        }

    }

}


//**************************************** */
//pdf関連
//**************************************** */

//pdfファイル作成関数
async function makepdf(pdfSaveFolderId, pdfFileName, ssId, shId) {
    await savePdf(pdfSaveFolderId, pdfFileName, ssId, shId);
    function savePdf(pdfSaveFolderId, pdfFileName, ssId, shId) {
        //PDFの保存先
        //★★★フォルダーIDを入力してください★★★
        let folderId = pdfSaveFolderId;

        //アクティブなスプレッドシートを取得する
        // let ss = SpreadsheetApp.getActiveSpreadsheet();

        //スプレッドシートIDを取得する
        // let ssId = ss.getId();

        //シートIDを取得する
        // let shId = ss.getActiveSheet().getSheetId();

        //★★★PDFのファイル名を入力してください★★★
        //※ポイント: ファイル名が重複しないようにしましょう
        let fileName = pdfFileName;

        //関数createPdfを実行し、PDFを作成して保存する
        createPdf(folderId, ssId, shId, fileName);
    }

    //PDFを作成し指定したフォルダーに保存する関数
    //以下4つの引数を指定する必要がある
    //1: フォルダーID (folderId)
    //2: スプレッドシートID (ssId)
    //3: シートID (shId)
    //4: ファイル名 (fileName)
    function createPdf(folderId, ssId, shId, fileName) {
        //PDFを作成するためのベースとなるURL
        let baseUrl = "https://docs.google.com/spreadsheets/d/"
            + ssId
            + "/export?gid="
            + shId;

        //★★★自由にカスタマイズしてください★★★
        //PDFのオプションを指定
        let pdfOptions = "&exportFormat=pdf&format=pdf"
            + "&size=A4" //用紙サイズ (A4)
            + "&portrait=true"  //用紙の向き true: 縦向き / false: 横向き
            + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
            + "&top_margin=0.65" //上の余白
            + "&right_margin=0.25" //右の余白
            + "&bottom_margin=0.25" //下の余白
            + "&left_margin=0.25" //左の余白
            + "&horizontal_alignment=CENTER" //水平方向の位置
            + "&vertical_alignment=MIDDLE" //垂直方向の位置
            + "&printtitle=false" //スプレッドシート名の表示有無
            + "&sheetnames=false" //シート名の表示有無
            + "&gridlines=false" //グリッドラインの表示有無
            + "&fzr=false" //固定行の表示有無
            + "&fzc=false" //固定列の表示有無
            //+ "&range: 'A2' + '%3A' + 'B2'" // セル範囲を指定（この場合だとA2:B2）
            + "&printnotes=false"  // メモを表示（falseは非表示）
            + "&scale=4"; // 印刷の拡大/縮小（4はページに合わせる）

        //PDFを作成するためのURL
        let url = baseUrl + pdfOptions;

        //アクセストークンを取得する
        let token = ScriptApp.getOAuthToken();

        //headersにアクセストークンを格納する
        let options = {
            headers: {
                'Authorization': 'Bearer ' + token
            }
        };

        //PDFを作成する
        let blob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName);

        //PDFの保存先フォルダー
        //フォルダーIDは引数のfolderIdを使用します
        let folder = DriveApp.getFolderById(folderId);

        //PDFを指定したフォルダに保存する
        folder.createFile(blob);
    }
}


//pdfファイル結合関数
async function merge(targetFolderId, exportFolderId) {
    //まとめて結合の対象フォルダ
    let target = targetFolderId;
    let exportman = exportFolderId;

    //eval関数で外部JSを取得させると導入できる
    eval(UrlFetchApp.fetch("https://unpkg.com/pdf-lib/dist/pdf-lib.js").getContentText());

    //setTimeoutをUtilities.sleepに置き換えてしまう
    setTimeout = (func, sleep) => (Utilities.sleep(sleep), func())
    await MergeFiles(target, exportman);
    //指定のフォルダ内のPDFをまとめて結合するメイン関数
    async function mergeAllPDFs(blobs, fileName) {
        //PDFLibsでパッケージ
        const pdf = await PDFLib.PDFDocument.create();

        //取得したPDFデータを次々に結合する
        for (let i = 0; i < blobs.length; i++) {
            //PDFデータをロードする
            const tempBytes = await new Uint8Array(blobs[i].getBytes());
            const tempPdf = await PDFLib.PDFDocument.load(tempBytes);

            //PDFページのページカウントを取得する
            const pages = tempPdf.getPageCount();

            //空のPDFに対して取得データを追加していく
            for (let p = 0; p < pages; p++) {
                const [tempPage] = await pdf.copyPages(tempPdf, [p]);
                pdf.addPage(tempPage);
            }
        }

        //PDFを保存する
        const pdfDoc = await pdf.save()

        //結合PDFを返却する
        return Utilities.newBlob(pdfDoc).setName(fileName);
    }

    //PDFをまとめて結合する関数
    async function MergeFiles(target, exportman) {
        //指定のフォルダ内のPDFを検索
        let folder = DriveApp.getFolderById(target);
        let files = folder.getFiles();
        let blobs = [];
        //blobsに対象のPDFをすべて配列に追加
        while (files.hasNext()) {
            let file = files.next();
            blobs.unshift(file.getBlob());
        }
        //PDF結合を実行
        let myPDF = await mergeAllPDFs(blobs, "result");
        //出力先フォルダに出力
        let exportFolder = DriveApp.getFolderById(exportman);
        let resultPdf = exportFolder.createFile(myPDF);
        let pdfUrl = resultPdf.getDownloadUrl();
        activeSheet.getRange(1, 3, 2, 10).merge().setFontColor("blue").setFontSize(18).setValue(pdfUrl);
        let sheet = ss.getSheetByName("シート1");
        sheet.activate();
    }
}


//************************************ */
//キャッシュサービスを使う
//************************************ */
function makeUserCache() {
    const cache = CacheService.getUserCache();
    return {
        get: function (key) {
            return JSON.parse(cache.get(key));
        },
        put: function (key, value, sec) {

            cache.put(key, JSON.stringify(value), (sec === undefined) ? 600 : sec);
            return value;
        }
    };
}

function makeDocumentCache() {
    const cache = CacheService.getDocumentCache();
    return {
        get: function (key) {
            return JSON.parse(cache.get(key));
        },
        put: function (key, value, sec) {

            cache.put(key, JSON.stringify(value), (sec === undefined) ? 600 : sec);
            return value;
        }
    };
}

function makeScriptCache() {
    const cache = CacheService.getScriptCache();
    return {
        get: function (key) {
            return JSON.parse(cache.get(key));
        },
        put: function (key, value, sec) {

            cache.put(key, JSON.stringify(value), (sec === undefined) ? 600 : sec);
            return value;
        }
    };
}

//****************************** */
//開いたとき・値を変えたとき
//****************************** */
function onOpen() {
    const sheetname = 'シート1';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetname);
    ss.setActiveSheet(sheet);

    let nowdate = new Date();
    let nowdate2 = new Date(nowdate);
    let formatteddate = Utilities.formatDate(nowdate2, 'JST', 'yyyy/MM/dd');

    sheet.getRange('a1:a4').setValues([[formatteddate], [formatteddate], [""], [false],]);
    sheet.getRange('a1:a2')
        .setFontFamily('Google Sans Mono')
        .setFontSize(11);
    const activeSheet = ss.getName();
    if (activeSheet != sheetname) {
        sheet.activate();
    }
}

function onEdit(e) {
    const sheetname = 'シート1';

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName(sheetname);

    if (e.range.getRow() == 1 && e.range.getColumn() == 1) {
        if (sheet.getRange("a1").getValue() > sheet.getRange("a2").getValue()) {
            sheet.getRange("a2").setValue(sheet.getRange("a1").getValue());
        }
    }
    if (e.range.getRow() == 2 && e.range.getColumn() == 1) {
        if (sheet.getRange("a1").getValue() > sheet.getRange("a2").getValue()) {
            sheet.getRange("a1").setValue(sheet.getRange("a2").getValue());
        }
    }

}

function isdatediffdate(formerday, latterday) {
    formerday = new Date(formerday);
    latterday = new Date(latterday);
    let timestamp1 = Math.floor(formerday.getTime() / 1000);
    let timestamp2 = Math.floor(latterday.getTime() / 1000);
    const margin = 10 * 1000;

    let datediff = Math.abs((timestamp2 - timestamp1) / (60 * 60 * 24));

    return datediff;
}

function isdatediff(formerday, latterday, targetdiff) {
    formerday = new Date(formerday);
    latterday = new Date(latterday);
    let timestamp1 = Math.floor(formerday.getTime() / 1000);
    let timestamp2 = Math.floor(latterday.getTime() / 1000);
    const margin = 10 * 1000;

    if (Math.abs(timestamp2 - timestamp1 >= 60 * 60 * 24 * targetdiff - margin)) {
        return true;
    }
    return false;
}

function islatterdaylater(formerday, latterday) {
    formerday = new Date(formerday);
    latterday = new Date(latterday);
    let timestamp1 = Math.floor(formerday.getTime() / 1000);
    let timestamp2 = Math.floor(latterday.getTime() / 1000);

    if (timestamp2 - timestamp1 < 0) {
        return true;
    }
    return false;
}

//昇順に並び替え
function descOrder(val1, val2) {
    if (val1.order > val2.order) {
        return 1;
    } else if (val1.order < val2.order) {
        return -1;
    } else {
        return 0;
    }
}

//*************************************** */
//印刷するページ数を分析
//*************************************** */
function valueAnalysis1(writeValues = []) {
    let employeesCount = writeValues.length;
    let employeesCount1 = employeesCount / 9;
    let employeesCountstrArray = [], num;
    employeesCountstrArray = String(employeesCount).split('');
    let employeesCountTotal = employeesCountstrArray.map(function (e) {
        num = Number(e);
        num++;
        return num;
    });

    let page1;

    if (employeesCount1 <= 1) {
        page1 = 1;
    } else if (employeesCount1 > 1 && employeesCountTotal != "") {
        page1 = Math.ceil(employeesCount1);
    }

    return page1;
}

//強制終了
function exit_sub() {
    return false;
}