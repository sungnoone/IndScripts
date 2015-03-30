/*取出矩形多邊形框內含的文字框*/ 
/* 2011-01-12 v0.1 */
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============

====================================================================================*/

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
        //var txtFileName = "E:\\" + inDoc.name + "_" + formatDate2Str() + ".txt"; // 匯出路徑檔名 正式要殺掉
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */
            go();
            s1 = "取出矩形多邊形框內含的文字框，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "取出矩形多邊形框內含的文字框，有錯誤發生!" + e.toString();
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


/*====================== Function ===================================*/

// 針對 Rectangle、Polygon物件內，若內含文字框，剪出貼在原父物件所在文字流插入點( 僅適用於父物件具備錨定，若父物件未錨定將產生錯誤 )
function go(){
    for( pv1=inDoc.allPageItems.length-1 ; pv1>=0 ; pv1-- ){
        if( inDoc.allPageItems[pv1].constructor.name == "TextFrame" ){
            var inTf = inDoc.allPageItems[pv1];
            var pType = inTf.parent.constructor.name;
            switch( pType ){
                case "Rectangle":
                    var inRect = inTf.parent;
                    if( inRect.parent.constructor.name == "Character" ){
                        var pChar = inRect.parent;
                        var pInsert =null;
                        if(  pChar.insertionPoints.length != 0 ){                        
                            pInsert =  pChar.insertionPoints[ pChar.insertionPoints.length-1 ];
                        }else{
                            alert( "Rectangle父插入點不明" );
                        }
                        inDoc.select( inTf , SelectionOptions.REPLACE_WITH);
                        app.cut();
                        inDoc.select( pInsert , SelectionOptions.REPLACE_WITH );
                        app.paste();
                    }
                break;
                case "Polygon":
                    var inPoly = inTf.parent;
                    if( inPoly.parent.constructor.name == "Character" ){
                        var pChar = inPoly.parent;
                        var pInsert =null;
                        if(  pChar.insertionPoints.length != 0 ){                        
                            pInsert =  pChar.insertionPoints[ pChar.insertionPoints.length-1 ];
                        }else{
                            alert( "Polygon父插入點不明" );
                        }
                        inDoc.select( inTf , SelectionOptions.REPLACE_WITH);
                        app.cut();
                        inDoc.select( pInsert , SelectionOptions.REPLACE_WITH );
                        app.paste();
                    }
                break;
            }
        }
    }
}

//................................................................................................................................

//檢查參數是否有輸入
function checkParas(){
    return true; // 外部沒有定義參數
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

//................................................................................................................................

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
