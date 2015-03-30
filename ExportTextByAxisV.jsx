/*根據座標匯出文字框文字(直稿)*/
/* 2011-01-05 v1.0 產出桌面文檔與原檔同名*/
/* 2012-04-25 v2.0 結果改存放網路 存檔格式改為UTF8*/
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============

====================================================================================*/
var resultFileRootPath = "\\\\idserver3.hanlin.com.tw\\NetTemp";

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
            export2Text();
            s1 = "匯出文字檔，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "匯出文字檔，有錯誤發生!" + e.toString();
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


//================================ Main Function =====================================

// 匯出文字
function export2Text(){
    // 建立存放資料夾
    nowDate = new Date();
    var outputFolderName = resultFileRootPath + "\\" + inDoc.name + "_output";//Folder.desktop.fsName + "\\" + inDoc.name + "_output"
    var outputFolder = new Folder( outputFolderName );
    if( outputFolder.exists==false ){
        // 資料夾不存在
        var createBool = outputFolder.create();
        if( createBool==false ){
            return; // 建立夾子失敗
        }
    }
    var txtFileName = outputFolderName + "\\" + inDoc.name + "_" + formatDate2Str() + ".txt"; // 匯出路徑檔名
    //ss="\n"
    //writeStrToFile( "\n" , txtFileName );
    for( pv1=0 ; pv1<inDoc.pages.length ; pv1++ ){
        var pinPage=inDoc.pages[pv1];
        var pgSection = pinPage.appliedSection; // 章節
        var pgName = pgSection.name + pinPage.name;// 拼湊頁俗名
        writeStrToFile( "\r\n========== Page "+(pv1+1)+ "\tName:" + pgName +" ==========\n\r" , txtFileName );
        var pageObjs = new Array();
        for( pv3=0 ; pv3<pinPage.allPageItems.length ; pv3++ ){
            var objType =pinPage.allPageItems[pv3].constructor.name;
            if( objType=="TextFrame" ){ // 只針對文字框做處理
                pageObjs.push ( pinPage.allPageItems[pv3] );
            }
        }
        if( pageObjs != null ){
            var validObjs = filtersObjs( pageObjs ); // 過濾掉不能正常讀取座標物件            
            var objsAry = V_Sort( validObjs ); // 排序頁面物件( 直稿 )
            if( objsAry != null ){
                for( pv2=0 ; pv2<objsAry.length ; pv2++ ){
                    var typeInfo = objsAry[pv2].constructor.name;
                    // 排物件順序
                    var locationAry = new Array ();
                    switch( typeInfo ){
                        case "TextFrame":
                            var pinTf = objsAry[pv2];
                            if( pinTf.id==pinTf.startTextFrame.id ){ // 只要串鍊的第一個
                                writeStrToFile( pinTf.parentStory.contents.toString() + "\r\n" , txtFileName );
                            }
                        break;
                        case "Image":
                            var pinImg = objsAry[pv2];
                            writeStrToFile( pinImg.itemLink.name.toString() + "\r\n" , txtFileName );
                            //ss += "Image: " + pinImg.itemLink.name.toString() + "\n";
                        break;
                        case "PDF":
                            var pinPdf = objsAry[pv2];
                            writeStrToFile( pinPdf.itemLink.name.toString() + "\r\n" , txtFileName );
                            //ss += "PDF: " + pinPdf.itemLink.name.toString() + "\n";
                        break;
                        default: appendStrToFile( typeInfo + "\n\r" ,  resultFileRootPath + "\\ExportTextByAxis__例外物件清單.txt" );
                    }
                }
            }            
           
        }
    }
}

/*............................................................................ Child Function ....................................................................................*/

// 過濾物件合法性
function filtersObjs( objsAry ){
	if( objsAry.length == 0 ){ return null; } // 傳入的物件陣列為空
	var tmpAry = new Array(objsAry.length);
    var validObjs = new Array(); // 可正常取用座標屬性物件
    var invalidObjs = new Array(); // 無法正常讀取座標屬性的物件
    // 區別可以正常取得座標物件 及 無法正常取得座標物件
    for( mv1=0 ; mv1<objsAry.length ; mv1++ ){
        try{
            var get_Y1 = objsAry[mv1].geometricBounds[0];
            var get_X1 = objsAry[mv1].geometricBounds[1];
            var get_Y2 = objsAry[mv1].geometricBounds[2];
            var get_X2 = objsAry[mv1].geometricBounds[3];
            validObjs.push (objsAry[mv1]);
        }catch(e){
            invalidObjs.push (objsAry[mv1]);
        }
    }
    return validObjs; // 返回可以正常取得座標的物件    
}

// 排直稿s
function V_Sort( objs ){
    if( objs == null){
        return null;
    }
    var sortObjsByX2 = new Array(objs.length);// 存放以 x2 座標排列的物件陣列
    sortObjsByX2 = smallSort( objs , 3 );// 座標 x2 排序

    // 找出 x2 相同的數值陣列
    var sameValInX2 = new Array();
    var val = "0";
    var bool = false;
    for( var cv1=0 ; cv1<sortObjsByX2.length ; cv1++ ){
        if( sortObjsByX2[cv1].geometricBounds[3].toString()==val.toString() ){
            if( bool==false ){
                sameValInX2.push( sortObjsByX2[cv1].geometricBounds[3] );
                bool = true;
            }
        }else{
            val = sortObjsByX2[cv1].geometricBounds[3];
            bool = false;
        }
    }
    
    // Re-Sort  sortObjsByX2 By y1 根據清單找出物件排序，根據 y1
    for( var cv2=0 ; cv2<sameValInX2.length ; cv2++ ){    
        var x2SameObjs = new Array();
        for( var cv3=0 ; cv3<sortObjsByX2.length ; cv3++ ){            
            if( sortObjsByX2[cv3].geometricBounds[3].toString()==sameValInX2[cv2].toString() ){
                var insertPoint = cv3;
                x2SameObjs.push(sortObjsByX2[cv3]);                
            }
        }
        insertPoint = insertPoint - x2SameObjs.length + 1
        var sortByY1 = smallSort( x2SameObjs , 0 ); // 根據 y1 排序
        for( var cv4=0 ; cv4<sortByY1.length ; cv4++ ){
            sortObjsByX2.splice( insertPoint+cv4  , 1 , sortByY1[cv4] ); // 代回原陣列
        }        
    }
    //viewObjPosition( sortObjsByX2 ); // 偵錯紀錄
    return sortObjsByX2; // 回傳排序完的物件陣列
    
}

/*............................................................................. Sub Child Function .....................................................................................................*/

// objs 要排序的物件陣列        sortkey  指定要用哪個座標值排序( y1 , x1 , y2 , x2 )
function smallSort( objs , sortkey ){
    // sortkey  0=y1  1=x1   2=y2   3=x2
    var sortResultObjs = new Array();
    switch( sortkey ){
        case 0: // y1
            for( var v1=0 ; v1<objs.length ; v1++ ){
                sortResultObjs = objs.sort( function(a,b){ return a.geometricBounds[0]-b.geometricBounds[0]; } );// 座標 y1 排序
            }
        break;
        case 1: // x1
            for( var v1=0 ; v1<objs.length ; v1++ ){
                sortResultObjs = objs.sort( function(a,b){ return b.geometricBounds[1]-a.geometricBounds[1]; } );// 座標 x1 排序
            }
        break;
        case 2: // y2
            for( var v1=0 ; v1<objs.length ; v1++ ){
                sortResultObjs = objs.sort( function(a,b){ return a.geometricBounds[2]-b.geometricBounds[2]; } );// 座標 y2 排序
            }
        break;
        case 3: // x2
            for( var v1=0 ; v1<objs.length ; v1++ ){
                sortResultObjs = objs.sort( function(a,b){ return b.geometricBounds[3]-a.geometricBounds[3]; } );// 座標 x2 排序
            }
        break;
    }
    return sortResultObjs;// 返回排序後的物件陣列    
}

//檢查參數是否有輸入
function checkParas(){
//~     if( app.scriptArgs.isDefined( "exportpath" )){
//~         exportpath = app.scriptArgs.getValue( "exportpath" ); // 如果有給就取出
//~         return true;
//~     }else{
//~         return false;
//~     }
    return true;
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

// 依據陣列物件順序列出座標清單( 觀察用 )
function viewObjPosition( objs ){
    var posCol = "";
    var nowDate = new Date();
    posCol = "==============" + nowDate.toString() + "=============\n\r"; // 加上時間戳記
    for( var v1=0 ; v1<objs.length ; v1++ ){
        posCol += objs[v1].geometricBounds + "\n\r";
    }
    appendStrToFile( posCol , resultFileRootPath + "\\" + inDoc.name + "_物件座標清單.txt" ); //Folder.desktop.fsName + "\\" + inDoc.name + "_物件座標清單.txt"
}

// 紀錄正異常物件
function logFilterObjsResult( objs ){

}

// 寫文字檔
function writeStrToFile( str , filePathStr ){
	var txtFile = new File (filePathStr);
    txtFile.encoding = "UTF8";
	txtFile.open("a");
    if( str!=null ){
        txtFile.write ( str.toString() );
    }else{
        txtFile.write ( "NULL" );
    }
    txtFile.close();
}

// 寫文字檔( 附加 )
function appendStrToFile( str , filePathStr ){
	var txtFile = new File (filePathStr);
    txtFile.encoding = "UTF8";
	txtFile.open("a");
    if( str!=null ){
        txtFile.write ( str.toString() );
    }else{
        txtFile.write ( "NULL" );
    }
    txtFile.close();
}
