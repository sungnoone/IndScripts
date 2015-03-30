/*取得文件色票(文件必須事先開啟)*/ 
/*====== 參數區( 開頭以 //$ 開始行且在程式內前30行者，均視為參數宣告，供外部主程式擷取 )============
========================*/

//測試時期預設變數值( 正式使用時，必須註解掉，否則外部傳進來的參數會被覆蓋掉 )
//app.scriptArgs.set(  參數名稱 , 參數預設值 );

// 必須有開啟的文件
if( app.documents.length != 0 ){
    var idDoc = app.activeDocument;
    var colorAry = getColors();
    //alert( colorAry.length );
}

// 最後把結果字串陣列回傳給外部呼叫程式
colorAry; // 丟了

//================================== Function & Sub ==========================================

function getColors(){
    var idColors = idDoc.colors;
    var colorAry = new Array();
    var str = "";
    var i = 0;
    for( v1=0 ; v1<idColors.length ; v1++ ){
        if( trim(idColors[v1].name) !="" ){
            colorAry[v1] = idColors[v1].name;
            //str += colorAry[v1] + "\n";
            i++;
        }
    }
    return colorAry;
}

//去除頭尾空白
function trim( str ){
    return str.replace(/(^\s*)|(\s*$)/g, ""); 
} 
