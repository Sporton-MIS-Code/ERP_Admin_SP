<!-- #include file="../Tools/Function.asp" -->
<%
Response.Buffer=true 
set ManagCn=server.CreateObject("adodb.connection") 
ManagCn.Open GV_strConnAdmin

UpStr="Update AssetsDataBase Set BelongUnit=UseUnit "
Set UP= ManagCn.Execute(UpStr) 
UpStr2="Update Assets2DataBase Set BelongUnit=UseUnit"
Set UP= ManagCn.Execute(UpStr2) 

%>

