/*程式說明2*/
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

try{
    var inDoc = app.activeDocument;
    s1 = inDoc.name & "開啟中";
    s2 = "True";
}
catch(e){
    s1 = "沒有開啟文件!";
    s2 = "False";
}

// return message
retAry[0] = s1;
retAry[1] = s2;
retAry; // return caller