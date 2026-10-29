<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->

<%Response.Buffer=true
  Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

   sql = "select * from Assets2DataBase"
   sql = sql & " order by AssetsNumber "
   set rs = server.createobject("adodb.recordset")
   rs.open sql,ManagCn,3
if request("postback") = "1" then
for i = 1 to rs.recordcount
'刪除＝＝＝＝
 if request("medelete" & i) = "on" then  
  sql = "delete * from Assets2DataBase where AssetsNumber = '" & request("AssetsNumber" & i) & "' "
  ManagCn.execute(sql) 
  end if
next
end if

%>
<%'-----------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim arybgcolor		'顏色陣列
      arybgcolor = Array("#ffe4c4","#d1eeee")	'顯示列共有兩種顏色
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title><%=Trim(Request.QueryString("Title"))%></title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 10pt; line-height: 12pt; LETTER-SPACING: 1px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 3px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
-->
</style>
<body background="0153.gif"><center>
<form action= Managdelete_2.asp method=post onSubmit="return confirm('確定要刪除資料嗎?')">
<%Response.Write("<center><table border=3 bgcolor=slateblue class=size03><th>SPORTON(資產管理程式)</table>" &_
                 "<table><td bgcolor=pink class=size02>" & Request.QueryString("Title") & "</table></center>")
%>
<input type=hidden name=postback value=1>
  
        <table border=1 cellspacing=0 cellpadding=3 align=center class=size01>
          <tr valign=middle align=center bgcolor=skyblue> 
            <td>刪除
            <td>資產編號 
            <td>資產名稱 
          </tr>
          <%  


 i = 1
   while not rs.eof %>
          <tr bgcolor="<%=arybgcolor(i Mod 2)%>"> 
            <td valign=middle align=center> 
              <input type=checkbox name=Medelete<%=i%>>               
              <input type=hidden name=AssetsNumber<%=i%> value=<%=rs("AssetsNumber")%>>
            <td> 
              <% =rs("AssetsNumber")%>
            <td> 
              <% =rs("AssetsName")%>
            
          </tr>
          <%
i = i+1
 rs.movenext
wend%>
</table><br>
<input type=submit value="確 定 刪 除" class=size01>
</form>
</center>
</body>
</html>