/*複合字內含指定名稱字類轉換 */
/*2011-01-12 v0.1 支援字類： 數字、英文、羅馬字 */
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============

====================================================================================*/

var userDefineFontName1 = "數字";
var userDefineFontName2 = "英文";
var userDefineFontName3 = "羅馬字";

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
                replSingleFont( userDefineFontName1 );
                replSingleFont( userDefineFontName2 );
                replSingleFont( userDefineFontName3 );
                s1 = "複合字內含指定名稱字類轉換，完成!";
                s2 = "True"; // 執行成功
             }catch(e){
                s1 =  "複合字內含指定名稱字類轉換，失敗! " + e.toString();
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

/*======================= Master Function =================================*/

// 在文件中所有複合體內含的字類名稱尋找，若有符合名稱的字類，便在文件中搜尋有套用到該複合字的文字，套用在字類中所指定的字體
// 例如: 複合字為 [HLA01+標楷] ，其中定義 [英文字]這一類的字，便套用HLA01這個字體，但是在文件搜尋時，若以HLA01搜尋是找不到該字的。
// 必須以複合字體為搜尋條件才找的到，若將凡是套到該複合字的英文字都改套成HLA01，則可在搜尋時以HLA01為條件
// 找到套用該複合字的英文字，並做一些特處理
function replSingleFont( userDefineFontName ){
    var inFonts = inDoc.fonts;
    var inCFonts = inDoc.compositeFonts;
    var logDate = new Date();
    var mystr = "====================" + logDate.toString() + "====================\n\r"; // 紀錄文字 開頭 註記時間
    for( var pv1=0 ; pv1<inFonts.length ; pv1++ ){
        var inFont = inDoc.fonts[pv1];
        if( validCompositeFont( inFont )==true ){ // 驗證是否為合法有效字體
             if( inFont.fontType=="1718894932" ){  // 複合字
                var entFont = getCFontEnt( inFont , userDefineFontName ); // 取得複合字指定字類套用字體
                if( entFont != null ){ // 如果傳回不為null
                    mystr += inFont.index + "：\t" + inFont.name + "：\t內含字類：" + userDefineFontName + "\t被指定為：" + entFont + "\n\r";
                    var findResults = findByCompositeFont( inFont.name ); // 找到套用複合字的內容物
                    for( var pv2=0 ; pv2<findResults.length ; pv2++ ){
                        // 每個內容特定字類套用獨特字體
                        //str += "\t" + findResults[pv2].constructor.name + "\n\r";
                        var findObjs = replFont( findResults[pv2] , userDefineFontName , entFont ); // 複合字指定字類替換字體
                    }
                }
            }
        }
    }
    writeStrToFile( mystr , Folder.desktop.fsName + "\\extractCompositeFonts紀錄.txt" );
}



/*...................................................... Sub Function .........................................................*/

// 找出複合字內，指定字類字體為何?
function getCFontEnt( CFont , charTypeName ){
    var inCFont = inDoc.compositeFonts.itemByName(CFont.fullNameNative);  // 傳入的 Font 要轉為 compositeFont
    //alert( inCFont.name );
    for( var v1=0 ; v1<inCFont.compositeFontEntries.length ; v1++ ){
        var ent = inCFont.compositeFontEntries[v1];
        if( ent.name==charTypeName ){ // 比對名稱
            return ent.appliedFont; // 是就回傳其指定套用字體
        }
    }
    return null;// 都沒有
}



// 尋找文件中套用到指定複合字體的物件
function findByCompositeFont( cfontName ){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.appliedFont = cfontName;
    return objs = inDoc.findText(); // 找到物件
}

// 搜尋指定字類文字，套用指定字體
function replFont( inObj , charTypeName , fontName ){
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    switch( charTypeName ){
        case "數字":
            app.findTextPreferences.findWhat = "^9";
            app.changeTextPreferences.appliedFont = fontName;
        break;
        case "英文":
            app.findTextPreferences.findWhat = "^$";
            app.changeTextPreferences.appliedFont = fontName;
        break;
        case "羅馬字":
            app.findTextPreferences.findWhat = "^$";
            app.changeTextPreferences.appliedFont = fontName;
        break;
    }
    try{
        var findObjs = inObj.changeText();//inObj.findText();
        return findObjs;
    }catch(e){
        return "";
    }    
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

// 驗證複合字屬性正常取用，為合格正常的複合字
function validCompositeFont( font ){
    // 無法正確取出指定屬性之複合字體，視為無效字，傳回 false
    try{
        var attr1 = font.fontType; // 判斷字體類別
        return true;
    }catch(e){
        return false;
    }
}

//檢查參數是否有輸入
function checkParas(){
        return true;
}
