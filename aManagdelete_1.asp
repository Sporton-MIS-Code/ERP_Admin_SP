<!-- #include file="../Tools/Function.asp" -->
<%
Response.Buffer=true 
  response.Write("ssss")
   response.End
set ManagCn=server.CreateObject("adodb.connection") 
ManagCn.Open GV_strConnAdmin 
'"Admin",,"sportonmanag"  

   sql = "select * from AssetsDataBase"
   sql = sql & " order by AssetsNumber "
 
   set rs = server.createobject("adodb.recordset")
   rs.open sql,ManagCn,3
if request("postback") = "1" then
for i = 1 to rs.recordcount
'刪除＝＝＝＝
 if request("medelete" & i) = "on" then  
  sql = "delete * from AssetsDataBase where AssetsNumber = '" & request("AssetsNumber" & i) & "' "
  ManagCn.execute(sql) 
 end if
next
end if

%>
<body background="0153.gif">
<form action="Managdelete_1.asp" method=post onSubmit="return confirm('確定嗎要刪除資料嗎?')">
<input type=submit value="確 定 刪 除">
<p></p>
<input type=hidden name=postback value=1>


        <table border=1 cellspacing="0" cellpadding="0">
          <tr> 
            <td height="25">　 
            <td height="25">資產編號 
            <td height="25">資產名稱 
          </tr>
          <%  
 i = 1
   while not rs.eof %>
          <tr> 
            <td height="25"> 
              <input type=checkbox name=Medelete<%=i%>>
              刪除 
              <input type=hidden name=AssetsNumber<%=i%> value=<%=rs("AssetsNumber")%>>
            <td height="25"> 
              <% =rs("AssetsNumber")%>
            <td height="25"> 
              <% =rs("AssetsName")%>
            </tr>
          <%
i = i+1
 rs.movenext
wend%>
        </table>

        
</form>
