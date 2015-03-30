/*抽取特定範圍、條件文字 */
//$起始頁:=setPageStartRange
//$結束頁:=setPageEndRange
//$大於(Y1):=setStart_y1
//$小於(Y2):=setStart_y2
//$字體大小:=setFontSize
//$字體顏色名稱:=setFontColorName<<getDocColors.jsx

//~ app.scriptArgs.set( "setPageStartRange" , "7" );
//~ app.scriptArgs.set( "setPageEndRange" , "10" );
//~ app.scriptArgs.set( "setStart_y1" , "204" );
//~ app.scriptArgs.set( "setStart_y2" , "844" );
//~ app.scriptArgs.set( "setFontSize" , "16" );
//~ app.scriptArgs.set( "setFontColorName" , "" );


var setPageStartRange; // 起始頁
var setPageEndtRange; // 結束頁
var setStart_y1; // Y上
var setStart_x1; // X左
var setStart_y2; // Y下
var setStart_x2; // X右
var setFontSize; // 字體大小
var setPStyleName; // 段落樣式(暫無)
var setFontColorName; // 色票名稱

var necessaryBool1 = false; // 是否加入字體大小條件
var necessaryBool3 = false; // 是否加入字體顏色條件
//var outputFileName = Folder.desktop.fsName+"\\catchText.txt";

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
            catchText();            
            s1 = "抽取特定範圍條件文字，執行完畢!";
            s2 = "True"; // 執行成功
         }catch(e){
            s1 =  "抽取特定範圍條件文字，有錯誤發生!" + e.toString();
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

function catchText(){
    nowDate = new Date();
    var outputFileName = Folder.desktop.fsName+"\\"+inDoc.name + "_"  + formatDate2Str().toString() +".txt"; // 轉出文字檔以InDesign文件名稱為預設檔名
    clearFindChangePrefs(); // 乾淨搜尋選項
    // 設定搜尋條件
    // 字體大小
    if( setFontSize=="" || setFontSize==null ){
        necessaryBool1 = false; // 不要字體大小條件
        //alert( "未輸入字體大小" );
    }else{
        necessaryBool1 = true; // 要字體大小條件
        setFindPrefs_FontSize();// 設定字體大小條件
    }
    // 色票
    if( setFontColorName=="" || setFontColorName==null ){
        necessaryBool3 = false; // 不要字體顏色條件
    }else{
        setFindPrefs_FillColor( setFontColorName ); // 設定 色票 搜尋條件
        removeAppColor( setFontColorName ); // 色票必須在App層級加入才能搜尋，故作完之後要移除在App層級的色票
    }
    

    // 只在指定頁面範圍內		
    // 逐頁
    var longStr = "";
    for( var v1=setPageStartRange-1 ; v1<setPageEndRange ; v1++ ){
        longStr += "\n========== Page "+(v1+1)+" ==========\n";
        var tfAry = getTfCollect( inDoc.pages[v1] ); // 取得頁內應該要處理的文字框
        // alert( "第"+i+"頁" );
        tfAry = checkLocationStandard(tfAry); // 檢查是否座標範圍內
        // alert( tfAry.length );
        tfAry = sortTextFrames( tfAry ); // 排序文字框
        if( necessaryBool1==false && necessaryBool3==false ){
            // 都沒設定條件 範圍內就抓
            //var storyAry = new Array();
            for( var v2=0 ; v2<tfAry.length ; v2++ ){
                // storyAry.push(tfAry[j].parentStory);
                // 文框內容直接抓
                longStr += tfAry[v2].parentStory.contents+"\n"; // 收集字串
            }
        }
        else{
            // 範圍內 還要符合條件 才抓
            var resultAry = findTexts( tfAry ); // 執行
            if( resultAry != null ){ // 必須有找到東西才能做
                for( var v3=0 ; v3<resultAry.length ; v3++ ){
                    // 文框內 搜尋符合的
                    longStr += resultAry[v3].contents+"\n";// 收集字串
                }
            }
        } 
        // 寫入結果到文字檔
        // 一次寫入
        writeStrToFile(longStr , outputFileName);
    }

}

//============================= Sub Function ==============================================

//檢查參數
function checkParas(){
    if( app.scriptArgs.isDefined( "setPageStartRange" ) 
    && app.scriptArgs.isDefined( "setPageEndRange" ) 
    && app.scriptArgs.isDefined( "setStart_y1" ) 
    && app.scriptArgs.isDefined( "setStart_y2" ) 
    && app.scriptArgs.isDefined( "setFontSize" ) 
    && app.scriptArgs.isDefined( "setFontColorName" ) )
    {
        setPageStartRange = app.scriptArgs.getValue( "setPageStartRange" ); // 如果有給就取出
        setPageEndRange = app.scriptArgs.getValue( "setPageEndRange" );
        setStart_y1 = app.scriptArgs.getValue( "setStart_y1" );
        setStart_y2 = app.scriptArgs.getValue( "setStart_y2" );
        setFontSize = app.scriptArgs.getValue( "setFontSize" );
        setFontColorName = app.scriptArgs.getValue( "setFontColorName" );
        return true;
    }
    else
    {
        return false;
    }
}

// 取樣式名稱陣列
function getAryOfPStyle(){
	var PStyleAry = new Array();
	for( var v1=0 ; v1<inDoc.paragraphStyles.length ; v1++ ){
		var tmpPStyle = inDoc.paragraphStyles[v1];
		PStyleAry.push( tmpPStyle.name );
	}
	return PStyleAry;
}

// 取色票
function getAryOfColor(){
	var SwatchAry = new Array();
	for( var v1=0 ; v1<inDoc.swatches.length ; v1++ ){
		var tmpSwatch = inDoc.swatches[v1];
		SwatchAry.push( tmpSwatch.name );
	}
	return SwatchAry;
}

// 刪除在Application中的色票
function removeAppColor( ColorName ){
	try{
		app.colors.itemByName ( ColorName ).remove();
	}
	catch(e){}
}

// 找出要處理的文字框
function getTfCollect( pageObj ){
	var tfObjs = new Array();
	for( var v2=0 ; v2<pageObj.allPageItems.length ; v2++ ){
		var inPageItem = pageObj.allPageItems[v2];
		if( inPageItem.constructor.name=="TextFrame" ){ // 判斷是不是文字框
			var nowTf = inPageItem;
			if( nowTf.id==nowTf.startTextFrame.id ){ // 只取串鍊的第一個
				tfObjs.push( nowTf ); // 納入收藏
			}
		}
	}
	return tfObjs; // 返回收藏
}

// 檢查框座標位置
function checkLocationStandard( tfObjs ){
	var tfAry = new Array();
	for( var v1=0 ; v1<tfObjs.length ; v1++ ){
		var inTf = tfObjs[v1];
		if( inTf.geometricBounds[0]>=setStart_y1 && inTf.geometricBounds[2]<=setStart_y2 ){
			tfAry.push( inTf );
		}
	}
	return tfAry;
}

// 排列文字框順序
function sortTextFrames( tfObjs ){
	if( tfObjs.length == 0 ){ return null; }
	var tmpAry = new Array();
	// 先排 X 軸	
	for( var v1=0 ; v1<tfObjs.length ; v1++ ){
		tmpAry[v1] = {};
		tmpAry[v1].txt = tfObjs[v1].geometricBounds[3];
		tmpAry[v1].val = tfObjs[v1];
	}
	var tfAry_OrderX = tmpAry.sort( function(a,b){ return b.txt-a.txt; } ); // X 軸重新排序後結果	
	// 再排 Y 軸
	var nowX = tfAry_OrderX[0].val.geometricBounds[3]; // 紀錄最新 x2
	var tfAry_OrderY = new Array(); // 擺 Y 排序後新結果
	var sameOfx = new Array(); // 放 X 相同物件
	for( var v2=0 ; v2<tfAry_OrderX.length ; v2++ ){
		if( tfAry_OrderX[v2].val.geometricBounds[3]!=nowX ){ // X 不同
			// 排 Y
			sameOfx = sameOfx.sort( function(a,b){ return a.txt-b.txt; } );
			// 重置前 將已佇列的元素寫入新陣列
			for( var v3=0 ; v3<sameOfx.length ; v3++ ){
				// 寫回是帶 x2 回去
				tfAry_OrderY.push( {txt:sameOfx[v3].val.geometricBounds[0] , val:sameOfx[v3].val} );
			}
			nowX = tfAry_OrderX[v2].val.geometricBounds[3] // 重新登記 目前 X
			// 重置陣列
			sameOfx = new Array(); 		
			sameOfx[0] = {};		
			sameOfx[0].txt = tfAry_OrderX[v2].val.geometricBounds[0]; // 登記 y1 位置
			sameOfx[0].val = tfAry_OrderX[v2].val;
		}
		else{ // X 相同
			var num = sameOfx.length;
			sameOfx[num] = {};
			sameOfx[num].txt = tfAry_OrderX[v2].val.geometricBounds[0];
			sameOfx[num].val = tfAry_OrderX[v2].val;
		}	
	}
	// 最後一次的 sameOfx 要補入陣列
	sameOfx = sameOfx.sort ( function(a,b){ return a.txt-b.txt; } );
	// 帶入新陣列
	for( var v3=0 ; v3<sameOfx.length ; v3++ ){
		tfAry_OrderY.push( {txt:sameOfx[v3].val.geometricBounds[0] , val:sameOfx[v3].val} );
	}
	// 轉回單值陣列
	var resultAry = new Array();
	for( var v4=0 ; v4<tfAry_OrderY.length ; v4++ ){
		resultAry.push( tfAry_OrderY[v4].val );
	}
	return resultAry;
}

// 清除 搜尋選項
function clearFindChangePrefs(){
	app.findTextPreferences = NothingEnum.NOTHING;
	app.changeTextPreferences = NothingEnum.NOTHING;
}

// 設定字體大小 搜尋條件
function setFindPrefs_FontSize(){
	var findTextPrefs = app.findTextPreferences;
	with( findTextPrefs ){
		pointSize = setFontSize+"Q";
	}
}

// 設定 色票 搜尋條件
function setFindPrefs_FillColor( ColorName ){
	// 在Application建立一相同色票
	// 此應該是Adobe 物件 Bug 因為目前無法指定文件的色票給搜尋格式中的 fillColor , 會出現不在同一文件或工作區錯誤 ( CS4也有相同問題 )
	// 先取得文件中的色票
	var docColor = inDoc.swatches.itemByName(ColorName);
	var appColor = null;
	// App中有同名的先殺掉
	try{
		// 如果已經存在就殺掉
		appColor = app.colors.itemByName(ColorName);
		appColor.remove();
	}catch(e){}
	//appColor = app.colors.add( { name:docColor.name , colorValue:docColor.colorValue , space:docColor.space , model:docColor.model , label:docColor.label } );
	appColor = app.colors.add( docColor.properties ); // 增加色票
	var findTextPrefs = app.findTextPreferences;
	with( findTextPrefs ){
		fillColor = appColor; // 設定搜尋色票
	}
}

// 搜尋
function findTexts( searchObjs ){
    if( searchObjs==null ){
        return null;
    }
	// 來源是文字框
	var objAry = new Array(); // 存放找到物件
	// 每個文字框找
	for( var v1=0 ; v1<searchObjs.length ; v1++ ){
		var nowStory = searchObjs[v1].parentStory; // 要用內文找
		var results = nowStory.findText(); // 搜尋結果
		for( var v2=0 ; v2<results.length ; v2++ ){
			objAry.push(results[v2]);
		}
	}
	return objAry; // 回傳找到的物件集合陣列
}

// 目前文件單位偏好
function fontUnit(){
	
}

//~ // 寫文字檔( 未用到 )
//~ function writeFile( contentObjs , filePathStr ){
//~ 	var txtFile = new File (filePathStr);
//~ 	txtFile.open("w");
//~ 	for( var v1=0 ; v1<contentObjs.length ; v1++ ){
//~ 		txtFile.writeln ( contentObjs[v1].contents.toString() );
//~ 	}
//~ 	txtFile.close();
//~ }

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
