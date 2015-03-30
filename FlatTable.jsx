/*表格轉文字 v1.2*/ 
/* 2011-01-12 v0.1 */
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============

====================================================================================*/


//測試時期預設變數值( 正式使用時，必須註解掉，否則外部傳進來的參數會被覆蓋掉 )
//app.scriptArgs.set(  "loopmax" , "10" );
var loopmax = "5";

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var symStart = "〈";
var symEnd = "〉";
var symStyleName = "-"; // 附加樣式名稱區隔符
var symTbStart = "s-table"; // 物件前符
var symTbEnd = "e-table"; // 物件尾符
var symCStyleName = "Script-SymbolStyle";

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
        //var txtFileName = "E:\\" + inDoc.name + "_" + formatDate2Str() + ".txt"; // 匯出路徑檔名 正式要殺掉
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */
            flatTable();
            //標記字串以字元樣式定型，避免被後續字體替換程式誤換
            applyRegularCStyle(symStart+symTbStart+symEnd); //前符定型
            applyRegularCStyle(symStart+symTbEnd+symEnd); //後符定型
            s1 = "表格轉換，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "表格轉換，有錯誤發生!" + e.toString();
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


//================================== Function & Sub ==========================================

// 表格平面化
function flatTable(){
    for( pv1=inDoc.stories.length-1 ; pv1>=0 ; pv1-- ){
        var inStory = inDoc.stories[pv1]; // 從文章所有內文開始
        var loopCount = 0;
        while( inStory.tables.length != 0 && loopCount<loopmax ){
            for( pv2=0 ; pv2<inStory.tables.length ; pv2++ ){
                var inTable = inStory.tables[pv2];
                if( inTable.parent.constructor.name=="TextFrame" ){
                    var tableStyle = inTable.appliedTableStyle;
                    // 插入表格開始記號                    
                    var firstCell = inTable.cells[0]; // 表格第一個儲存格
                    var firstChar = firstCell.characters[0]; // 表格第一個儲存格的第一個字
                    if( inTable.cells[0].characters.length == 0 ){ // 若儲存格內沒有字
                        firstCell.contents =  symStart + symTbStart + symEnd ; // 儲存格內容完全指定為表格起始符號;
                    }else{ // 若儲存格內有字
                        var firstInsert = inTable.cells[0].insertionPoints[0]; // 前符插入點
                        firstInsert.contents = symStart + symTbStart  + symEnd; // 在儲存格第一字的第一個插入點，塞入表格起始符號
                    }
                    // 插入表格結束記號                    
                    var lastCell = inTable.cells[inTable.cells.length-1]; // 最後儲存格                    
                    var lastChar = lastCell.characters[lastCell.characters.length-1]; // 最後儲存格的最後一字
                    if( lastCell.characters.length == 0 ){ // 如果最後一個儲存格沒有字
                        lastCell.contents = symStart + symTbEnd + symEnd; //  儲存格內容完全指定為表格結束符號;
                    }else{ // 若儲存格內有字
                        var lastInsert = lastChar.insertionPoints[lastChar.insertionPoints.length-1]; // 後符插入點
                        lastInsert.contents =  symStart + symTbEnd + symEnd; // 塞入表格結束符號
                    }
                    inTable.convertToText( "\t" , "\r" ); // 表格轉文字
                }
            }
            loopCount +=1;
        }
    }
}

//檢查參數是否有輸入
function checkParas(){
    return true
//~     if( app.scriptArgs.isDefined( "loopmax" )){
//~         loopmax = app.scriptArgs.getValue( "loopmax" );
//~         return true; // 外部有定義參數
//~     }else{
//~         return false; // 外部沒有定義參數
//~     }    
}

//指定字串套規則字體
function applyRegularCStyle( searchString ){
    var myCStyle = symbolCStyleCheck();
    
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;
    var findPrefs = app.findTextPreferences;
    var changePrefs = app.changeTextPreferences;
    
    findPrefs.findWhat = searchString;
    changePrefs.appliedCharacterStyle = myCStyle;
    var result = inDoc.changeText();
    
}

//=================================================================================

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

// 標記的字元樣式
function symbolCStyleCheck(){
	// style exist flag
	var flag ="0";
	// check if assign style exists
    var symCStyle = null;
	for ( cci=0 ; cci<inDoc.characterStyles.length ; cci++ ){
		symCStyle = inDoc.characterStyles[cci];
		if ( symCStyle.name == symCStyleName ){
			flag = "1";  // exist
		}
	}
	// let's create character style if flag value is "1"
	if ( flag =="0" ){
		symCStyle = inDoc.characterStyles.add();
		symCStyle.name =  symCStyleName;
	}
    symCStyle.appliedFont = "新細明體";
    return symCStyle;
}


