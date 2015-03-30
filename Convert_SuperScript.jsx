/*轉換上標字*/ 
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============
//$前標字串:=startStr
//$後標字串:=endStr
========================*/

//測試時期預設變數值( 正式使用時，必須註解掉，否則外部傳進來的參數會被覆蓋掉 )
//app.scriptArgs.set(  "startStr" , "<s-SuperScript>" );
//app.scriptArgs.set(  "endStr" , "<e-SuperScript>" );

var startStr = "^";//前標字串"<s-SuperScript>"
var endStr = "";//後標字串"<e-SuperScript>"

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
        // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */            
            Convert_SubScript();
            s1 = "轉換上標文字，執行完畢!";
            s2 = "True"; // 執行成功
         }catch(e){
            s1 =  "轉換上標文字，有錯誤發生!" + e.toString();
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

//================================== Main Function  ==========================================

//轉換上標字
function Convert_SubScript(){
    clearFindChangePrefs(); // 乾淨搜尋選項
    var findTextPrefs = app.findTextPreferences; // 建立搜尋條件    
    with( findTextPrefs ){
        position = Position.SUPERSCRIPT; // 上標字        
    }
    var rltObjs = app.findText();
    //alert( rltObjs.length );
    var str = "";
    for( var pv1=rltObjs.length-1 ; pv1>=0 ; pv1-- ){
        var findObj = rltObjs[pv1];
        findObj.position = Position.NORMAL;//恢復正常位置(避免當指令重複執行時，重複加標得錯誤)
        var lastChar = findObj.characters[findObj.characters.length-1];//最後一個字
        var lastInsert = lastChar.insertionPoints[lastChar.insertionPoints.length-1];//最後一個字最後的插入點        
        lastInsert.contents = endStr;
        var firstChar = findObj.characters[0];//第一個字
        var firstInsert = firstChar.insertionPoints[0];//第一個字的第一個插入點
        firstInsert.contents = startStr;
        //str += "開始插入點:" +  firstInsert.contents + "內容:" + findObj.contents + "最後插入點:"  + lastInsert.contents + "\n";
    }
    writeStrToFile( str , "E:\\上標字.txt" );
}

//檢查參數是否有輸入......................................................................
function checkParas(){
    if( app.scriptArgs.isDefined( "startStr" ) &&
         app.scriptArgs.isDefined( "endStr" )   ){ // 判斷外部呼叫程式是否給了程式所需參數的值
        startStr = app.scriptArgs.getValue( "startStr" ); // 如果有給就取出
        endStr = app.scriptArgs.getValue( "endStr" );
        return true;
    }else{
        return false;
    }
}

// 寫文字檔......................................................................
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

//================================== Sub Function  ==========================================

// 清除 搜尋選項......................................................................
function clearFindChangePrefs(){
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;
}

