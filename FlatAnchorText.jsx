/*錨定文字平整*/
/* 2011-01-12 v0.1 */
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============
//$迴圈上限:=loopmax
====================================================================================*/


//app.scriptArgs.set(  "loopmax" , "6" );
var loopmax = "6";


// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var symStart = "〈";
var symEnd = "〉";
var symStyleName = "-"; // 附加樣式名稱區隔符
var symTfStart = "s-TextFrame"; // 錨定框前符
var symTfEnd = "e-TextFrame"; // 錨定框前符


//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
        //var txtFileName = "E:\\" + inDoc.name + "_" + formatDate2Str() + ".txt"; // 匯出路徑檔名 正式要殺掉
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */            
            flatAnchorText();
            s1 = "錨定文字平整，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "錨定文字平整，有錯誤發生!" + e.toString();
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

//============================= Main Function ==============================================

function flatAnchorText(){
    clearFindChangePrefs(); // 乾淨搜尋選項
    // 設定搜尋條件
    var loopCount = 0; // 計算迴圈次數
    while( countTfWithTf().length !=0 && loopCount<loopmax ){
        var tfAry = getTfCollect(); // 尋找內容不含物件的文字框
        //alert( tfAry.length );
        // alert( tfAry.length );
        for( var pv2=tfAry.length-1 ; pv2>=0 ; pv2-- ){
            replCharByCode( tfAry[pv2] );
            var pinChar = tfAry[pv2].parent;
            var pinStory = pinChar.parentStory;
            var parChar = pinStory.characters[pinChar.index];// 錨定點文字
            try{
                insPointslen = parChar.insertionPoints.length; // 測試屬性是否能正常存取
            }catch(e){
                parChar = pinStory.characters[pinStory.characters.length-1]; // 插入文字點物件無法解析，擺在文本最後面
            }
            var pinsert = parChar.insertionPoints[parChar.insertionPoints.length-1] // 以最後後面為插入點
            var OStyleName = getApplyOStyleName(tfAry[pv2]);
            parChar.contents = symStart + symTfStart + symEnd 
                                        + tfAry[pv2].parentStory.contents.toString() 
                                        + symStart + symTfEnd + symEnd; // 錨定框內容塞入文字流內
//~             pinsert.contents= symStart + symTfStart + symStyleName + OStyleName + symEnd 
//~                                         + tfAry[pv2].parentStory.contents.toString() 
//~                                         + symStart + symTfEnd + symEnd; // 錨定框內容塞入文字流內
//~             try{
//~                 tfAry[pv2].remove();
//~             }catch(e){
//~             }            
        }
        loopCount += 1; // 迴圈累加
    }
//~     for( var pv3=tfAry.length-1 ; pv3>=0 ; pv3-- ){ // 刪除已平整化的錨定物件
//~         //tfAry[pv3].locked = false;
//~         //tfAry[pv3].remove(); // 原先錨定框刪除
//~     }
}

//============================= Sub Function ==============================================

//檢查參數是否有輸入
function checkParas(){
    if( app.scriptArgs.isDefined( "loopmax" )){
        loopmax = app.scriptArgs.getValue( "loopmax" );
        return true; // 外部有定義參數
    }else{
        return false; // 外部沒有定義參數
    }
}

// 清除 搜尋選項
function clearFindChangePrefs(){
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;
}

// 找出要處理的文字框
function getTfCollect(){
	var tfObjs = new Array();
	for( var v2=inDoc.allPageItems.length-1 ; v2>=0 ; v2-- ){
        var inPageItem = inDoc.allPageItems[v2];
        var typeName = inPageItem.constructor.name;
        switch( typeName ){
            case "TextFrame":
                if( inPageItem.parent.constructor.name=="Character" ){ // 是不是錨定物件
                    var nowTf = inPageItem;
                    if( nowTf.id==nowTf.startTextFrame.id ){ // 只取串鍊的第一個
                        if( nowTf.allPageItems.length == 0 ){ // 如果裡面沒有其他物件了
                            // nowTf.parentStory.contents = "";
                            tfObjs.push( nowTf ); // 文框內沒有包含物件了才納入收藏
                        }                        
                    }
                }else{
                    //writeStrToFile( inPageItem.parent.constructor.name , txtFileName );
                }
            break;
        }
	}
	return tfObjs; // 返回收藏
}

// 檢查文件中還包含物件的文字框數量
function countTfWithObj(){
    var tfWithObjAry = new Array();
    for( v1=0 ; v1<inDoc.allPageItems.length ; v1++ ){ // 從文件的所有頁面物件開始
        var inObj = inDoc.allPageItems[v1];
        var typeName = inObj.constructor.name; // 取得物件型別
        if( typeName=="TextFrame" ){
            if( inObj.allPageItems.length ){
                tfWithObjAry.push( inObj );// 尚有包含物件的算進去
            }
        }
    }
    return tfWithObjAry; // 回傳有包含物件文框的數量
}

// 文框內還有文框的數量
function countTfWithTf(){
    var tfWithTfAry = new Array();
    for( v1=0 ; v1<inDoc.allPageItems.length ; v1++ ){ // 從文件的所有頁面物件開始
        var inObj = inDoc.allPageItems[v1];
        var typeName = inObj.constructor.name; // 取得物件型別
        if( typeName=="TextFrame" ){
            if( inObj.textFrames.length ){
                tfWithTfAry.push( inObj );// 尚有包含物件的算進去
            }
        }
    }
    return tfWithTfAry; // 回傳有包含物件文框的數量
}

// 把文框內的分段符號都換掉
function replCharByCode( inTf ){
    var inStory = inTf.parentStory;
    if( inStory.paragraphs.length==0 ){
        return;
    }
    for( v1=inStory.paragraphs.length-1 ; v1>=0 ; v1-- ){
        var parghLen = inStory.paragraphs[v1].characters.length;
        if( inStory.paragraphs[v1].contents.charCodeAt( parghLen-1 ) == "13" ){
            inStory.paragraphs[v1].characters[parghLen-1].contents = "//";
        }
    }
}

// 取得物件樣式
function getApplyOStyleName( inObj ){
    var OStyle=inObj.appliedObjectStyle;
    if( OStyle != null ){
        return OStyle.name;
    }else{
        return "Null";
    }
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
	txtFile.open("a");
    if( str!=null ){
        txtFile.write ( str.toString() );
    }else{
        txtFile.write ( "NULL" );
    }
    txtFile.close();
}