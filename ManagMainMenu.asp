<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->

<%'-----------------------------------------------------------------------------------------------------------------------------------------------------------
  Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

  '-----------------------------------------------------------------------------------------------------------------------------------------------------------
  Str1 = "SELECT Admin_ShowPower.ProgID, ProgList.ProgName " &_
         "FROM Admin_ShowPower INNER JOIN ProgList ON Admin_ShowPower.ProgID = ProgList.ProgID " &_
         "WHERE (Admin_ShowPower.CName = '"& Session("Cname") &"') AND (Admin_ShowPower.RootLevel = 1) " &_
         "ORDER BY Admin_ShowPower.ProgID"
  Set rs1 = GF_GetRS(Str1,ManagCn)
  If rs1.EOF Then%>
     <script language=javascript>
             alert('您尚無任何權限！');
             document.location.href = 'Default.asp';
     </script>
   <%response.End
  End If%>

<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title>資產管理系統-主功能選單</title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 10pt; line-height: 12pt; LETTER-SPACING: 1px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 3px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
-->
</style>
</head>
<body name="Menu" background="0153.gif">
<center>
<table border=3 bgcolor=slateblue class="size03"><th>SPORTON(資產管理系統)</table>
<%Do While Not rs1.EOF
           RS1_ProgID = Trim(rs1("ProgID"))
           RS1_ProgName = Trim(rs1("ProgName"))
           
           Str2 = "SELECT ProgList.ProgName, ISNULL(ProgList.URL, '') AS URL, ISNULL(ProgList.CheckManag, '') AS CheckManag, " &_
                  "ISNULL(ProgList.Check1, '') AS Check1, ISNULL(ProgList.Check2, '') AS Check2 " &_
                  "FROM Admin_ShowPower INNER JOIN ProgList ON Admin_ShowPower.ProgID = ProgList.ProgID " &_
                  "WHERE (Admin_ShowPower.GroupLevel = "& RS1_ProgID &") AND (Admin_ShowPower.RootLevel = 2) AND (Admin_ShowPower.CName = '"& Session("Cname") &"') " &_
                  "ORDER BY Admin_ShowPower.ProgID"
           Set rs2 = GF_GetRS(Str2,ManagCn)%>
           <table cellspacing="2" border="0" cellpadding="3" align="center" class="size02">
                  <caption><font class="size02"><U><%=RS1_ProgName%></U></font>
           <%If Not rs2.EOF Then
                    i = 0
                    response.Write "<tr>"
                    Do While Not rs2.EOF
                             i = i+1
                             RS2_ProgName = Trim(rs2("ProgName"))
                             RS2_URL = Trim(rs2("URL"))
                             RS2_CheckManag = Trim(rs2("CheckManag"))
                             RS2_Check1 = Trim(rs2("Check1"))
                             RS2_Check2 = Trim(rs2("Check2"))%>
                             <td bgcolor=dodgerblue><input type=checkbox name="Ask" value="查詢" onclick="this.checked=false;window.location='<%=RS2_URL%>?Title=<%=RS2_ProgName%>&CheckManag=<%=RS2_CheckManag%>&Check=<%=RS2_Check1%>&CheckC=<%=RS2_Check2%>';"><%=RS2_ProgName%></td>
                             <%If (i Mod 2) = 0 Then
                                  response.Write "</tr><tr>"
                               End If
                       rs2.MoveNext
                    Loop
                    response.Write "</tr>"
             End If
                 rs2.Close
             Set rs2 = Nothing%>     
           </table><br>
   <%rs1.MoveNext
  Loop
      rs1.Close
  Set rs1 = Nothing
      ManagCn.Close
  Set ManagCn = Nothing%>
<table class="size01"><td bgcolor=pink>注 意 事 項</table>
<table class="size01"><tr><td bgcolor=deepskyblue>在使用上如有任何問題，請洽資訊部(分機219、336、338)！謝謝！</table>
</center>
</body>
</html>