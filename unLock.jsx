/*取消物件鎖定*/
/*2011-01-12 v0.1*/
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============
//$解鎖次數:=round
========================*/

//測試時期預設變數值( 正式使用時，必須註解掉，否則外部傳進來的參數會被覆蓋掉 )
//app.scriptArgs.set(  "round" , "20" );

// 每個指令碼都要建立回傳陣列 陣列內容為字串
// [0] 要告訴爸爸的話    [1] 只允許True或False代表成功或失敗
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否

var round = ""; // 解次數

//  從外部呼叫程式傳過來的參數要把值取出來
if( checkParas() == true ){ // 判斷外部呼叫程式是否給了程式所需參數的值   
    if( app.documents.length != 0 ){ // 要有開啟的文件
        var inDoc = app.activeDocument;
            // 務必加上錯誤擷取機制，以防指令碼無法將錯誤回傳
        try{
            /*        程式內容     */
            unLockObj();
            s1 = "取消物件鎖定，執行完畢!";
            s2 = "True"; // 執行成功
        }catch(e){
            s1 =  "取消物件鎖定，有錯誤發生!" + e.toString();
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

// 解散群組
function unLockObj(){
    var win
    for( pv1=0 ; pv1<round ; pv1++ ){ // 重複做滿指定的次數        
        for( pv2=inDoc.pageItems.length-1 ; pv2>=0 ; pv2-- ){            
            inDoc.pageItems[pv2].locked = false; // 取消鎖定
        }
    }
}

// ====================== Sub Function ===========================

//檢查參數是否有輸入
function checkParas(){
    if( app.scriptArgs.isDefined( "round" ) ){ // 判斷外部呼叫程式是否給了程式所需參數的值
        round = app.scriptArgs.getValue( "round" ); // 如果有給就取出
        return true;
    }else{
        return false;
    }
}


