/*程式說明1

*/ 
//$處裡檔案:=inputFile
//$色票名稱:=swatchName
//$字體大小:=fontSize

//app.scriptArgs.set( "inputFile" , "\\\\idServer3\\Export\\20107313164087-生字詞語卡.indd" );
var retAry = new Array(); // 回傳陣列
var s1 = ""; // 回傳給外部呼叫程式的文字訊息
var s2 = "False"; // 回傳給外部呼叫程式，本指令碼執行成功與否
// get parameters
if( app.scriptArgs.isDefined( "inputFile" ) ){
	var inputFileStr = app.scriptArgs.getValue( "inputFile" );
	if( checkFileExists( inputFileStr ) ){
		//var inDoc = app.open( new File ( inputFileStr ) );
        
        //inDoc.close();
		s1 = inputFileStr + "  存在!";
		s2 = "True";
		}
	else{
		s1 =  inputFileStr + "  不存在!";
		s2 = "False";
		}
	}
else{
	s1 = " 沒有足夠必要的參數! ";
	s2 = "False";
	}

// return message
retAry[0] = s1;
retAry[1] = s2;
retAry; // return caller

//================================== Function & Sub ==========================================

// Checking for the file if exists
function checkFileExists( fileName ){
	var myFile = new File( fileName );
	if( myFile.exists ){
		return true;
		}
	else{
		return false;
		}	
	}
