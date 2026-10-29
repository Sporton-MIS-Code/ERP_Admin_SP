<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->
<%Function Show_Str(Str)
           If ISNULL(Str) OR Str = "" OR Str = "False" Then
              Show_Str = "　"
           Else
               Show_Str = Str
           End If
  End Function%>

<%CheckManag = Trim(Request.QueryString("CheckManag"))
  Check = Trim(Request.QueryString("Check"))
  FileNumber = UCase(Trim(Request.QueryString("FileNumber")))

  '************************************************************************************************
  Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

  '************************************************************************************************
  Select Case CheckManag
         Case "FixAssetsRegistQ"
              If Check <>"M" Then  '列帳
                 If FileNumber <> clear Then
                    Str = "SELECT * FROM AssetsDataBase WHERE (AssetsNumber= '"& FileNumber &"')"
                 Else
                     Str = "SELECT * FROM AssetsDataBase WHERE (AssetsNumber= '1234567890')"
                 End If
              Else  '列管
                 If FileNumber <> clear Then
                    Str = "SELECT * FROM Assets2DataBase WHERE (AssetsNumber= '"& FileNumber &"')"
                 Else
                     Str = "SELECT * FROM Assets2DataBase WHERE (AssetsNumber= '1234567890')"
                 End If  
              End If
  End Select

  Set ManagRs = GF_GetRS(Str,ManagCn)
  
  If ManagRs.EOF Then%>
     <script language=javascript>
             alert('您所輸入的條件，目前並無資料！\n請重新輸入新的查詢條件，謝謝！')
             history.back(-1);
     </script>
   <%response.End
  Else
      AssetsNumber = Trim(ManagRs.fields("AssetsNumber"))
      AssetsName = Trim(ManagRs.fields("AssetsName"))
      AssetsEut = Trim(ManagRs.fields("AssetsEut"))
      AssetsOrdinalNo = Trim(ManagRs.fields("AssetsOrdinalNo"))
      Amount = Trim(ManagRs.fields("Amount"))
      Unit = Trim(ManagRs.fields("Unit"))
      Insure = Trim(ManagRs.fields("Insure"))
      Mortgage = Trim(ManagRs.fields("Mortgage"))
      Iease = Trim(ManagRs.fields("Iease"))
      
      SellCondition = Trim(ManagRs.fields("SellCondition"))
      If SellCondition <> "" Then
         CaseStateBuffer1=Split(SellCondition, "`")
         Count = UBound(CaseStateBuffer1)
         CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
         SellCondition_Str = CaseStateBuffer(3)
      Else
         SellCondition_Str = "正常" 
      End If
      
      ObtainDate = Trim(ManagRs.fields("ObtainDate"))
      Brand = Trim(ManagRs.fields("Brand"))
      BuyCast = Trim(ManagRs.fields("BuyCast"))
      EnduringLimit = Trim(ManagRs.fields("EnduringLimit"))
      KeepDate = Trim(ManagRs.fields("KeepDate"))
      BusinessDepartment = Trim(ManagRs.fields("BusinessDepartment"))
      UseUnit = Trim(ManagRs.fields("UseUnit"))
      BelongUnit = Trim(ManagRs.fields("BelongUnit"))
      AssetsAssortment = Trim(ManagRs.fields("AssetsAssortment"))
      Remark = Trim(ManagRs.fields("Remark"))      
      CustomerBaseData = Trim(ManagRs.fields("CustomerBaseData"))
      ChangesCondition = Trim(ManagRs.fields("ChangesCondition"))
      EstimateCondition = Trim(ManagRs.fields("EstimateCondition"))
      MainTainRecord = Trim(ManagRs.fields("MainTainRecord"))
      RectifyRecord = Trim(ManagRs.fields("RectifyRecord"))
  End If
      ManagRs.Close
  Set ManagRs = Nothing
      ManagCn.Close
  Set ManagCn = Nothing%>

<html>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title><%=Trim(Request.QueryString("Title"))%></title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 9pt; line-height: 11pt; LETTER-SPACING: 0px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 2px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
-->
</style>
</head>
<body background="0153.gif">
<center>
<table border=3 bgcolor=slateblue class="size03"><th>SPORTON(資產管理系統)</table>
<table><td bgcolor=pink class=size02><%=Request.QueryString("Title")%></table>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=skyblue class=size01 width="90%">
       <tr align=center>
           <td colspan=12>資 產 卡</td>
       </tr>
       <tr align=center bgcolor=yellow>
           <td colspan=2>資產編號</td>
           <td colspan=2>資產名稱</td>
           <td>型號</td>
           <td>序號</td>
           <td>數量</td>
           <td>單位</td>
           <td>投抵</td>
           <td>抵押</td>
           <td>租賃</td>
           <td>目前狀況</td>
       </tr>
       <tr align=center>
           <td colspan=2><%=Show_Str(AssetsNumber)%></td>
           <td colspan=2><%=Show_Str(AssetsName)%></td>
           <td><%=Show_Str(AssetsEut)%></td>
           <td><%=Show_Str(AssetsOrdinalNo)%></td>
           <td><%=Show_Str(Amount)%></td>
           <td><%=Show_Str(Unit)%></td>
           <td><%=Show_Str(Insure)%></td>
           <td><%=Show_Str(Mortgage)%></td>
           <td><%=Show_Str(Iease)%></td>
           <td><%=Show_Str(SellCondition_Str)%></td>
       </tr>
       <tr align=center bgcolor=yellow>
           <td>取得日期</td>
           <td>廠牌</td>
     <%If Session("Cname") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then%>
           <td>金額</td>
     <%End If%>
           <td>耐用年限</td>
           <td>保固期</td>
           <td>事業部</td>
           <td>保管單位</td>
           <td>歸屬單位</td>
           <td colspan=3>資產分類</td>
           <td colspan=2>備註</td>
       </tr>
       <tr align=center>
           <td><%=Show_Str(ObtainDate)%></td>           
           <td><%=Show_Str(Brand)%></td>
     <%If Session("Cname") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then%>           
           <td><%=Show_Str(BuyCast)%></td>
     <%End If%>
           <td><%=Show_Str(EnduringLimit)%></td>
           <td><%=Show_Str(KeepDate)%></td>
           <td><%=Show_Str(BusinessDepartment)%></td>
           <td><%=Show_Str(UseUnit)%></td>
           <td><%=Show_Str(BelongUnit)%></td>
           <td colspan=3><%=Show_Str(AssetsAssortment)%></td>
           <td colspan=2><%=Show_Str(Remark)%></td>
       </tr>
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=hotpink class=size01 width="90%">
       <tr align=center>
           <td colspan=5>廠商基本資料</td>
       </tr>
       <tr bgcolor=pink align=center>
           <td>廠商名稱</td>
           <td>聯絡人員</td>
           <td>聯絡電話</td>
           <td>傳真電話</td>
           <td>住址</td>
       </tr>
<%If CustomerBaseData <> "" Then
     CaseStateBuffer1=Split(CustomerBaseData,"`")
     If UBound(CaseStateBuffer1) <> -1 Then
        For Count=0 To UBound(CaseStateBuffer1)
                CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
      <%Next
     Else
         CaseStateBuffer=Split(CustomerBaseData,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=steelblue class=size01 width="90%">
       <tr align=center>
           <td colspan=5>異動情形</td>
       </tr>
       <tr bgcolor=skyblue align=center>
           <td>日期</td>
           <td>異動單號</td>
           <td>移出單位</td>
           <td>移入單位</td>
           <td>說明</td>
       </tr>
<%If ChangesCondition<>"" Then
     CaseStateBuffer1=Split(ChangesCondition,"`")
     If UBound(CaseStateBuffer1)<>-1 Then
        For Count=0 To UBound(CaseStateBuffer1)
            CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
      <%Next
     Else
     response.Write "bbb"
         CaseStateBuffer=Split(ChangesCondition,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
       
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=steelblue class=size01 width="90%">
       <tr align=center>
           <td colspan=5>改良增值及資產重估</td>
       </tr>
       <tr bgcolor=skyblue align=center>
           <td>日期</td>
           <td>改量單號</td>
           <td>規格名稱</td>
           <td>金額</td>
           <td>廠商</td>
       </tr>
<%If EstimateCondition <> "" Then
     CaseStateBuffer1=Split(EstimateCondition,"`")
     If UBound(CaseStateBuffer1)<>-1 Then
        For Count=0 To UBound(CaseStateBuffer1)
            CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
      <%Next
     Else
         CaseStateBuffer=Split(EstimateCondition,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=steelblue class=size01 width="90%">
       <tr align=center>
           <td colspan=6>維修記錄</td>
       </tr>
       <tr bgcolor=skyblue align=center>
           <td>日期</td>
           <td>維修單號</td>
           <td>規格名稱</td>
           <td>金額</td>
           <td>保固期</td>
           <td>廠商</td>
       </tr>
       <tr align=center>
<%If MainTainRecord <> "" Then
     CaseStateBuffer1=Split(MainTainRecord,"`")
     If UBound(CaseStateBuffer1)<>-1 Then
        For Count=0 To UBound(CaseStateBuffer1)
            CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
           <td><%=Show_Str(CaseStateBuffer(5))%></td>
       </tr>
      <%Next
     Else
         CaseStateBuffer=Split(MainTainRecord,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
           <td><%=Show_Str(CaseStateBuffer(5))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=steelblue class=size01 width="90%">
       <tr align=center>
           <td colspan=6>儀器校正紀錄</td>
       </tr>
       <tr bgcolor=skyblue align=center>
           <td>日期</td>
           <td>序號</td>
           <td>型號</td>
           <td>金額</td>
           <td>校驗廠商</td>
           <td>下次需檢驗日期</td>
       </tr>
<%If RectifyRecord <> "" Then
     CaseStateBuffer1=Split(RectifyRecord,"`")
     If UBound(CaseStateBuffer1)<>-1 Then
        For Count=0 To UBound(CaseStateBuffer1)
            CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
           <td><%=Show_Str(CaseStateBuffer(5))%></td>
       </tr>
      <%Next
     Else
         CaseStateBuffer=Split(RectifyRecord,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
           <td><%=Show_Str(CaseStateBuffer(5))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
</table><br>
<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=steelblue class=size01 width="90%">
       <tr align=center>
           <td colspan=5>出售/報廢/閑置</td>
       </tr>
       <tr bgcolor=skyblue align=center>
           <td>日期</td>
           <td>單號</td>
           <td>處份原因</td>
           <td>處份方式</td>
           <td>殘值</td>
       </tr>
<%If SellCondition <> "" Then
     CaseStateBuffer1=Split(SellCondition,"`")
     If UBound(CaseStateBuffer1)<>-1 Then
        For Count=0 To UBound(CaseStateBuffer1)
            CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
      <%Next
     Else
         CaseStateBuffer=Split(SellCondition,"^")
         If UBound(CaseStateBuffer) <> -1 Then%>
       <tr align=center>
           <td><%=Show_Str(CaseStateBuffer(0))%></td>
           <td><%=Show_Str(CaseStateBuffer(1))%></td>
           <td><%=Show_Str(CaseStateBuffer(2))%></td>
           <td><%=Show_Str(CaseStateBuffer(3))%></td>
           <td><%=Show_Str(CaseStateBuffer(4))%></td>
       </tr>
       <%End If
     End If
  Else%>
       <tr align=center>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
           <td>　</td>
       </tr>
<%End If%>
</table><br>
</body>
</html>