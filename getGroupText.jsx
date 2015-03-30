/*取出錨定群組內的文字 v1.01 */

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var symStart = "〈";
var symEnd = "〉";
var symStyleName = "-"; // 附加樣式名稱區隔符
var symTfStart = "s-TextFrame"; // 文字框前符
var symTfEnd = "e-TextFrame"; // 文字框後符
var symParghStart = "s-Paragraph"; // 段落前符
var symParghEnd = "e-Paragraph"; // 段落後符
var symPargh = "//" //分段符號

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */            
            getGroupContents();
            s1 = "取出群組內的文字，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "取出群組內的文字，有錯誤發生!" + e.toString();
            s2 = "False"; // 執行失敗 
        }
    }else{
        s1 =  "沒有開啟的文件";
        s2 = "False"; // 執行失敗         
    }
}else{
    // 外部的呼叫程式沒有給予必要參數時
	s1 = " 沒有足夠必要的參數! ";
	s2 = "False";
}

// 最後把結果字串陣列回傳給外部呼叫程式
retAry[0] = s1;
retAry[1] = s2;
retAry; // 丟了

// ========================= Function ============================

// 取出群組內的文字
function getGroupContents(){
    var pageObjs = inDoc.allPageItems;
    var pGroupAry = new Array();
    //var ss = "";
    for( pv1=0 ; pv1<pageObjs.length ; pv1++ ){
        if( pageObjs[pv1].constructor.name == "Group" ){
            pGroupAry.push (pageObjs[pv1]); // 找出是群組的物件
        }
    }    
    for( pv2=pGroupAry.length-1 ; pv2>=0 ; pv2-- ){ 
        //ss += "Group:" + pGroupAry[pv2].id.toString() + "   " + pGroupAry[pv2].parent.constructor.name + "\n";        
        if( pGroupAry[pv2].parent.constructor.name == "Character" ){ // 如果爸爸是文字流中的一分子
            pGroupAry[pv2].locked = false; // 確定鎖定是解除的
            var parentChar = pGroupAry[pv2].parent;
            var norStr = normalization( pGroupAry[pv2] );
            parentChar.contents = norStr; // 以內容文字取代群組物件
            //ss += "      Group contents: " + norStr + "\n";
        }
    //ss += "\n================================\n";
    }
    //writeStrToFile( ss ,  "E:\\"+formatDate2Str()+".txt" );
}

// ====================== Sub Function ===========================

//檢查參數是否有輸入
function checkParas(){
    return true; // 不需參數
}

// 正規化群組物件內文字
function normalization( groupObj ){
    var tmpStr = "";
    for( v1=groupObj.allPageItems.length-1 ; v1>=0 ; v1-- ){
        // 群組內的每個物件
        var subObj = groupObj.allPageItems[v1];
        if( subObj.constructor.name == "TextFrame" ){ // 只處裡文字框內容
            var tfOStyleName = "Null"
            if( subObj.appliedObjectStyle != null ){
                tfOStyleName = subObj.appliedObjectStyle.name.toString();
            }
            tmpStr += symStart + symTfStart + symEnd; // 插入文框開始記號  
            var insideStory = subObj.parentStory;//文框的內文
            for( v2=0 ; v2<insideStory.paragraphs.length ; v2++ ){ // 文字框內段落可能不只一個
                var parghContents = insideStory.paragraphs[v2].contents.toString();
                var parghLen = parghContents.length; // 段落長度
                if( parghContents.charCodeAt(parghLen-1) == "13" ){ // 判斷段落結尾是不是分段符號
                    parghContents = parghContents.replace ( parghContents[parghLen-1] , symPargh )
                }
                tmpStr += symStart + symParghStart + symEnd + parghContents + symStart + symParghEnd + symEnd; // 段落內容
            }
            tmpStr += symStart + symTfEnd + symEnd; // 文框結束
        }
    }
    return tmpStr;
}

// 日期字串格式化
function formatDate2Str(){
    nowDate = new Date();
    strYear = nowDate.getFullYear().toString(); // 年
    strMonth = (nowDate.getMonth()+1).toString(); // 月
    if ( strMonth.length < 2 ){ // 補零
        strMonth = "0" + strMonth;
    }
    strDate = nowDate.getDate().toString(); // 日
    if ( strDate.length < 2 ){ // 補零
        strDate = "0" + strDate;
    }
    strHour = nowDate.getHours().toString(); // 時
    if ( strHour.length < 2 ){ // 補零
        strHour = "0" + strHour;
    }
    strMin = nowDate.getMinutes().toString(); // 分
    if ( strMin.length < 2 ){ // 補零
        strMin = "0" + strMin;
    }
    strSec = nowDate.getSeconds().toString(); // 秒
    if ( strSec.length < 2 ){ // 補零
        strSec = "0" + strSec;
    }
    strMSec = nowDate.getMilliseconds().toString(); // 微秒

    var str = strYear + strMonth + strDate + strHour + strMin + strSec + "_" + strMSec;
    return str;
}

// 寫文字檔
function writeStrToFile( str , filePathStr ){
	var txtFile = new File (filePathStr);
	txtFile.open("w");
    if( str!=null ){
        txtFile.write ( str.toString() );
    }else{
        txtFile.write ( "NULL" );
    }
    txtFile.close();
}

