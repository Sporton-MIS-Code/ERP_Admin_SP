<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title>資產管理系統</title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 10pt; line-height: 12pt; LETTER-SPACING: 1px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 3px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
-->
</style>

<script language=javascript>
        function chkdata()
        {
                 if (document.login.EmpID.value == '')
                 {
                     alert('請輸入「帳號」！');
                     document.login.EmpID.focus();
                     return false;
                 }

                 //--------------------------------------------------------------------------------------------------------------
	             if (document.login.EmpID.value != '')
	             {
                     var str = document.login.EmpID.value;
	                 var a = str.indexOf('<');
			         var b = str.indexOf('>');
			         var c = str.indexOf("'");
			         var d = str.indexOf('"');

			         if ((a != -1) || (b != -1) || (c != -1) || (d != -1))
			         {
			              alert('帳號\n\n請勿  含有『HTML語法』或『單引號』或『雙引號』喔！');
 			              document.login.EmpID.focus();
			              return false;
	                  }
	             }

                 //--------------------------------------------------------------------------------------------------------------
                 if (document.login.Pwd.value == '')
                 {
                     alert('請輸入「密碼」！');
                     document.login.Pwd.focus();
                     return false;
                 }

                 //--------------------------------------------------------------------------------------------------------------
	             if (document.login.Pwd.value != '')
	             {
                     var str = document.login.Pwd.value;
	                 var a = str.indexOf('<');
			         var b = str.indexOf('>');
			         var c = str.indexOf("'");
			         var d = str.indexOf('"');

			         if ((a != -1) || (b != -1) || (c != -1) || (d != -1))
			         {
			              alert('密碼\n\n請勿  含有『HTML語法』或『單引號』或『雙引號』喔！');
 			              document.login.Pwd.focus();
			              return false;
	                  }
	             }
        }
</script>
</head>
<body onload="login.EmpID.focus();" background="0153.gif">
<center>
<FORM name="login" method=post action="Login.asp" onsubmit="return chkdata();">
      <table border=3 bgcolor=slateblue class="size03"><th><center>SPORTON(資產管理系統)</center></table><br>
      <table cellspacing="3" border="0" cellpadding="3" align="center" class="size02">
             <tr>
                 <td>帳號：<input type=text name="EmpID" size="10" maxlength="10" class="size01"></td>
             </tr>
             <tr>
                 <td>密碼：<input type=password name="Pwd" size="10" maxlength="10" class="size01"></td>
             </tr>
      </table><br>
      <input type=reset value="清 除" class="size02">　
      <input type=submit value="登 入" class="size02">
</FORM>
<table class="size01"><td bgcolor=pink>注 意 事 項</table>
<table class="size01"><tr><td bgcolor=deepskyblue>使用上如有問題，請與資訊部聯絡(分機219、336、338)！謝謝！</table>
</center>
</body>
</html>