<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" language="java" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>테스트화면</title>
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<link rel="shortcut icon" href="#">

<style>
	* {
	    font-size: 16px;
	    font-family: Consolas, sans-serif;
  	}
  	textarea {
	    width: 80%;
	    height: 100px;
	}
</style>

<script type="text/javascript">
	function excelDownload(){
		var jsonData = $("#jsonData").val();
	
		if(!validationCheck(jsonData)){
			alert("입력값 확인");
			return;
		}
		doExcelDownloadProcess();
	}
	
	function validationCheck(jsonData){
		if(jsonData != "" && jsonData != null){
			return true;
		}else{
			return false;
		}
	}
	
	function doExcelDownloadProcess(){
		var f = document.form1;
        f.action = "/downloadExcelFile";
        f.submit();
	}
</script>

</head>

<body>

<form id="form1" name="form1" method="post" enctype="multipart/form-data">
	<textarea id = "jsonData" name="jsonData" ></textarea>
    <button type="button" onclick="excelDownload()">엑셀다운로드 작업</button>
</form>
</body>
</html>