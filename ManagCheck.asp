<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title><%=Trim(Request.QueryString("Title"))%></title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 10pt; line-height: 12pt; LETTER-SPACING: 0px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 2px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
-->
</style>
<script language="javascript">
        function CheckInput()
        {
                 Check=true;
                 if (!Menu.Business.checked&&!Menu.Unit.checked&&!Menu.AssNumber.checked&&!Menu.Obtain.checked&&!Menu.State.checked)
                 {
                     alert('上列查詢條件並未勾選!!');Check=false;
                 }
                 else
                 {
                    if(Menu.ObtainBegin.value==''&&Check==true&&Menu.Obtain.checked)
                    {
                       alert('請填寫欲查詢的日期.謝謝!!');Check=false;
                    }
                 } 
                 return Check;
        }
</script>
</head>
<script language="vbscript"></script>
<body background="0153.gif" class="size02" onload="if(('<%=Request.QueryString("CheckManag")%>'=='FixAssetsRegist') ||('<%=Request.QueryString("CheckManag")%>'== 'PrintAssetsRegist')) {document.Menu.FileNumber.focus()};"> 
<%Response.Write("<center><table border=3 bgcolor=slateblue class=size03><th>SPORTON(資產管理程式)</table>" &_
                 "<table><td bgcolor=pink class=size02>" & Request.QueryString("Title") & "</table></center>")

  select case Request.QueryString("CheckManag")
         '******************************************************************************************************************************************************
         case "CheckDetailedList","PrintMark":
              Response.Write("<form name=Menu method=get action=ManagShow.asp>" &_     
                             "<center><fieldset style='width:500;'></center>" &_
                             "<legend>請輸入欲查詢的條件</legend>" &_
                             "<li><input type=checkbox name=Business checked>依事業部查詢：<select name=BusinessName class=size01><option>Component<option>EMC<option>Safety<option>BTL</select>" &_
                             "<li><input type=checkbox name=Unit checked>依部門查詢：" &_
                             "<select name=UnitName class=size01>" &_
                             "<option>認證一部華亞實驗室" &_
                             "<option>認證一部技術支援部" &_
                             "<option>認證一部研發部" &_
                             "<option>認證一部品管部" &_
                             "<option>認證一部文件部" &_
                             "<option>認證一部案件管理部" &_
                             "<option>認證一部安規工程部" &_
                             "<option>認證一部汐止工程部" &_
                             "<option>認證一部東湖工程部" &_
                             "<option>認證一部林口工程部" &_
                             "<option>認證一部EMC業務部" &_
                             "<option>認證一部Safety業務部" &_
                             "<option>認證一部新竹營運處業務部" &_
                             "<option>認證一部新竹營運處工程部" &_
                             "<option>認證二部業務部" &_
                             "<option>認證二部安規工程部" &_
                             "<option>認證二部EMC工程部" &_
                             "<option>認證三部業務部" &_
                             "<option>認證三部工程組" &_
                             "<option>認證三部修改組" &_
                             "<option>認證三部文件組" &_
                             "<option>認證五部天線設計部" &_
                             "<option>認證五部RFI工程部" &_
                             "<option>認證五部RF工程部" &_
                             "<option>認證五部OTA工程部" &_
                             "<option>零件事業部業務部" &_
                             "<option>零件事業部資材部" &_
                             "<option>總經理室" &_
                             "<option>財務部" &_
                             "<option>資訊部" &_
                             "<option>客服部" &_
                             "<option>稽核室" &_
                             "<option>公共行政部" &_
                             "</select>")
              if Request.QueryString("CheckManag") = "PrintMark" then
                 Response.Write("<li>依資產編號查詢：<input type=text name=FileNo size=15 maxlength=15 class=size01>")
                 Response.Write("<li>輸入盤點年份：<input type=text name=PrintYear size=4 maxlength=4 class=size01>")
              end if
              Response.Write("</fieldset>")
         '******************************************************************************************************************************************************
         case "FixAssetsRegist","PrintAssetsRegist":
              Response.Write("<form name=Menu method=get action=ManagShow.asp>" &_
                             "<center><fieldset style='width:500;'>" &_
                             "<legend>請輸入欲修改的資產編號</legend>" &_
                             "資產編號：<input type=text name=FileNumber size=15 maxlength=15 class=size01>" &_
                             "</fieldset></center>")

         '******************************************************************************************************************************************************
         case "FixAssetsRegistQ"
              Response.Write("<form name=Menu method=get action=ManagShowQ.asp>" &_
                             "<center><fieldset style='width:500;'>" &_
                             "<legend>請輸入欲查詢的資產編號</legend>" &_
                             "資產編號：<input type=text name=FileNumber size=15 maxlength=15 class=size01>" &_
                             "</fieldset></center>")

         '******************************************************************************************************************************************************
         case "DetailedList":
              Response.Write("<form name=Menu method=get action=ManagShow.asp onsubmit='return CheckInput();'>" &_     
                             "<center><fieldset style='width:500;'></center>" &_
                             "<legend>請輸入欲查詢的條件</legend>" &_
                             "<li><input type=checkbox name=Business>依事業部查詢：<select name=BusinessName class=size01><option>Component<option>EMC<option>Safety<option>BTL</select>" &_
                             "<li><input type=checkbox name=Unit>依部門查詢：" &_
                             "<select name=UnitName class=size01>" &_
                             "<option>認證一部華亞實驗室" &_
                             "<option>認證一部技術支援部" &_
                             "<option>認證一部研發部" &_
                             "<option>認證一部品管部" &_
                             "<option>認證一部文件部" &_
                             "<option>認證一部案件管理部" &_
                             "<option>認證一部安規工程部" &_
                             "<option>認證一部汐止工程部" &_
                             "<option>認證一部東湖工程部" &_
                             "<option>認證一部林口工程部" &_
                             "<option>認證一部EMC業務部" &_
                             "<option>認證一部Safety業務部" &_
                             "<option>認證一部新竹營運處業務部" &_
                             "<option>認證一部新竹營運處工程部" &_
                             "<option>認證二部業務部" &_
                             "<option>認證二部安規工程部" &_
                             "<option>認證二部EMC工程部" &_
                             "<option>認證三部業務部" &_
                             "<option>認證三部工程組" &_
                             "<option>認證三部修改組" &_
                             "<option>認證三部文件組" &_
                             "<option>認證五部天線設計部" &_
                             "<option>認證五部RFI工程部" &_
                             "<option>認證五部RF工程部" &_
                             "<option>認證五部OTA工程部" &_
                             "<option>零件事業部業務部" &_
                             "<option>零件事業部資材部" &_
                             "<option>總經理室" &_
                             "<option>財務部" &_
                             "<option>資訊部" &_
                             "<option>客服部" &_
                             "<option>稽核室" &_
                             "<option>公共行政部" &_
                             "</select>" &_
                             "<li><input type=checkbox name=State>依狀態查詢：<select name=StateSel class=size01><option>報廢<option>出售<option>閑置<option>正常<option>投底<option>抵押<option>租賃<option>盤盈<option>盤虧</select>" &_
                             "<li><input type=checkbox name=AssNumber>依資產分類查詢：<select name=AssSel class=size01><option>土地<option>房屋及建築<option>機器設備<option>運輸設備<option>辦公設備<option>其它設備<option>租賃改良</select>" &_
                             "<li><input type=checkbox name=Obtain>依取得日期：<input type=text name=ObtainBegin value='"&date&"' size=10 class=size01> ~ <input type=text name=ObtainEnd value='"&date&"' size=10 class=size01>" &_
                             "</fieldset>")    
  end select
    
  Response.Write("<input type=hidden name=CheckManag value='"& Request.QueryString("CheckManag") &"'>" &_
                 "<input type=hidden name=Title value='"& Request.QueryString("Title") &"'>" &_
                 "<input type=hidden name=CheckC value='"& Request.QueryString("CheckC") &"'>" &_
                 "<input type=hidden name=Check value='"& Request.QueryString("Check") &"'>" &_
                 "<center><br><input type=submit name=button value='查詢' class=size01><br>" &_
                 "<br><br><fieldset style='width:500;'>注意事項</fieldset><fieldset style='width:500;'><li>如有問題，請洽資訊部(分機219、336、338)。謝謝！</fieldset></center>" &_
                 "</form></body></html>")
%>