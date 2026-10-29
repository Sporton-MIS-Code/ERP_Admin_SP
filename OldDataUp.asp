<!-- #include file="../Tools/Function.asp" -->
<%
Response.Buffer=true 
set newcn=server.CreateObject("adodb.connection") 
set newrs=server.CreateObject("adodb.recordset") 
set oldcn=server.CreateObject("adodb.connection") 
set oldrs=server.CreateObject("adodb.recordset") 
set oldrs2=server.CreateObject("adodb.recordset")
set oldrs3=server.CreateObject("adodb.recordset")
 
newcn.Open GV_strConnAdmin'"Admin",,"sportonmanag"

'oldcn.Open "old"
'oldrs.Open "select * from EMC列帳",oldcn,3,3
'oldrs2.Open "select * from 零件事業部列帳",oldcn,3,3
'oldrs3.Open "select * from 安規列帳",oldcn,3,3


'do until oldrs.EOF 
 
'  Str="Insert into AssetsDataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit) "
'  Str=Str & "values('E-' & '"& left(oldrs.Fields("財產編號"),4) &"' & '"& right(oldrs.Fields("財產編號"),5) &"','"& oldrs.Fields("名稱") &"','"& oldrs.Fields("型號") &"','"& oldrs.Fields("序號") &"','"& oldrs.Fields("數量") &"','"& oldrs.Fields("單位") &"','"& oldrs.Fields("取得日期") &"','"& oldrs.Fields("購入成本") &"','"& oldrs.Fields("耐用年限") &"','"& oldrs.Fields("存放地點") &"')"
 
'  newrs.Open Str,newcn,3,3
 
'  oldrs.MoveNext 
'loop


'do until oldrs2.EOF 
  
'  Str="Insert into AssetsDataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit) "
'  Str=Str & "values('C-' & '"& oldrs2.Fields("財產編號") &"','"& oldrs2.Fields("名稱") &"','"& oldrs2.Fields("型號") &"','"& oldrs2.Fields("序號") &"','"& oldrs2.Fields("數量") &"','"& oldrs2.Fields("單位") &"','"& oldrs2.Fields("取得日期") &"','"& oldrs2.Fields("購入成本") &"','"& oldrs2.Fields("耐用年限") &"','"& oldrs2.Fields("存放地點") &"')"
 
'  newrs.Open Str,newcn,3,3
 

'  oldrs2.MoveNext 
'loop

'do until oldrs3.EOF 
  
'  Str="Insert into AssetsDataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit) "
'  Str=Str & "values('S-' & '"& left(oldrs3.Fields("財產編號"),1) &"' & '"& right(oldrs3.Fields("財產編號"),8) &"','"& oldrs3.Fields("名稱") &"','"& oldrs3.Fields("型號") &"','"& oldrs3.Fields("序號") &"','"& oldrs3.Fields("數量") &"','"& oldrs3.Fields("單位") &"','"& oldrs3.Fields("取得日期") &"','"& oldrs3.Fields("購入成本") &"','"& oldrs3.Fields("耐用年限") &"','"& oldrs3.Fields("存放地點") &"')"
'  on error resume next
'  newrs.Open Str,newcn,3,3


'  oldrs3.MoveNext 
'loop
'newrs.Open "Update AssetsDataBase Set NewState='正常'",newcn,3,3
newrs.Open "Update AssetsDataBase Set BusinessDepartment='Component' where AssetsNumber LIKE 'C-%'",newcn,3,3
newrs.Open "Update AssetsDataBase Set BusinessDepartment='Emc' where AssetsNumber LIKE 'E-%'",newcn,3,3
newrs.Open "Update AssetsDataBase Set BusinessDepartment='Safety' where AssetsNumber LIKE 'S-%'",newcn,3,3



%>

