/*特殊字體加標*/ 
/* 2011-01-12 v0.2 */
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============
//$:碼轉換定義檔:=fontmapfile
========================*/

//測試時期預設變數值( 正式使用時，必須註解掉，否則外部傳進來的參數會被覆蓋掉 )
//app.scriptArgs.set(  "fontmapfile" , "\\\\idServer3\\FontMapRef\\CodeMap\\E00078\\物理備課.txt" ); //  字體對應表

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var fontmapfile = ""; // 碼轉換定義檔
var fontmapSep1 = ";" // 碼轉換定義行字串切割符
var fontmapSep2 = "," // 碼轉換定義 替換字串切割符
var codebookfile = "\\\\idserver3\\FontMapRef\\CodeBook\\codebook.txt"; //鍵盤碼簿

var txtFileName = "E:\\" + formatDate2Str() + ".txt"; // 匯出路徑檔名

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值
        if( app.documents.length != 0 ){ // 要有開啟的文件
            var inDoc = app.activeDocument;
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
            try{
                /*        程式內容     */
                tranFont()
                s1 = "特殊字體加標，完成!";
                s2 = "True"; // 執行成功
             }catch(e){
                s1 =  "特殊字體加標，失敗! " + e.toString();
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

// ======================== main function ============================

// 轉字
function tranFont( ){    
    // 讀入鍵盤碼
    var codeAry = new Array(); // 鍵盤碼
    var codefile = new File( codebookfile );
    codefile.open ( "r" );
    while( codefile.eof == false ) {
        var lineStr = codefile.readln();
        codeAry.push(lineStr); // 塞入鍵盤碼陣列        
    }
    //alert( codeAry.length );
    codefile.close();
    // 讀入對應表
    var mapfile = new File( fontmapfile );
    mapfile.open ( "r" );
    var tmpAry = new Array(); // 對應碼行
    while( mapfile.eof == false ){
        var lineStr = mapfile.readln();        
        tmpAry.push(lineStr); // 切割字串
    }
    mapfile.close();
    // 拆對應碼行
    for( pv1=0 ; pv1<tmpAry.length ; pv1++ ){
        replaceFontChar( codeAry , tmpAry[pv1] ); // 鍵盤碼 與 字體對應碼 結合
    }
}

//================================== Function & Sub ==========================================

// 
function replaceFontChar( codeAry , mapLine ){
    //ss = "";
    var fontAry = new Array();
    fontAry = mapLine.split( fontmapSep1 ); // 切割字串
    if( fontAry.length<4 ){
        return;
    }
    var fontName = fontAry[0]; // 字體名稱
    var replSymStart = fontAry[1]; // 前符
    var replSymEnd = fontAry[2]; // 後符
    var replSymAry = new Array(); // 替換文字陣列
    replSymAry = fontAry[3].split( fontmapSep2 );
    //ss += fontName + replSymStart + replSymEnd + replSymAry + "\n";
    //writeStrToFile( ss , txtFileName );
    
    if( checkFontExist(fontName) == false ){ // 該指定字體在文件中不存在
        //alert( fontName + " 不存在" );
        return false;
    }

    //尋找字體
    //ss += fontName + " \n"
    for( v1=0 ; v1<replSymAry.length ; v1++ ){
        app.findTextPreferences = NothingEnum.nothing;
        app.changeTextPreferences = NothingEnum.nothing;
        with ( app.findTextPreferences ){
            appliedFont = fontName;
            findWhat = codeAry[v1];
        }        
        var objs = inDoc.findText(); // 找到物件
        //ss += "  " + codeAry[v1] + "  數量：" + objs.length + "\n";
        for( v2=objs.length-1 ; v2>=0 ; v2-- ){
            //ss += "   find Object: " + objs[v2].contents.toString() + "\n";
            objs[v2].appliedFont = "新細明體";
            if( replSymAry[v1]!="" ){
                objs[v2].contents = replSymStart + replSymAry[v1].toString() + replSymEnd; // 替換內容
            }else{
                //對應替換字若為空，表示不替換原字碼
                objs[v2].contents = replSymStart +codeAry[v1].toString() + replSymEnd; // 只加前後符
            }            
        }
    }
     //writeStrToFile( ss , txtFileName );
}

// 檢查文件是否有指定字體
function checkFontExist( fontNameStr ){
    var docFonts = inDoc.fonts;
    for( v1=0 ; v1<docFonts.length ; v1++ ){
        if( docFonts[v1].name == fontNameStr ){
            return true;
        }
    }
    return false; // 一直沒找到
}

//檢查參數是否有輸入
function checkParas(){    
    if( app.scriptArgs.isDefined( "fontmapfile" ) )
    {
        fontmapfile = app.scriptArgs.getValue( "fontmapfile" ); // 如果有給就取出
        return true;
    }else{
        return false;
    }
}

// =================== 通用函數 ==========================

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


