<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->
<%Function Show_Str(Str)
           If ISNULL(Str) OR Str = "" Then
              Show_Str = "　"
           Else
               Show_Str = Str
           End If
  End Function%>

<%Function Chk_Selected(Str1, Str2)
           If Str1 = Str2 Then
              Chk_Selected = "Selected"
           Else
               Chk_Selected = ""
           End If
  End Function%>

<%Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

  Select Case Request.QueryString("CheckManag")
         '******************************************************************************************************************************************************
         Case "CheckDetailedList","PrintMark"
              If Request.QueryString("CheckC") = "M" Then
                 Str = "SELECT * FROM Assets2DataBase WHERE "
              Else
                  Str = "SELECT * FROM AssetsDataBase WHERE "
              End If

              If Request.QueryString("Business") = "on" Then Str = Str & "(BusinessDepartment = '"& Request.QueryString("BusinessName") &"') AND "
              If Request.QueryString("Unit") = "on" Then Str = Str & "(UseUnit LIKE '%"& Request.QueryString("UnitName") &"%') AND " 

              Select Case Request.QueryString("AssSel")
                     Case "土地"        AssSel = "A"
                     Case "房屋及建築"  AssSel = "B"
                     Case "機器設備"    AssSel = "C"
                     Case "運輸設備"    AssSel = "D"
                     Case "辦公設備"    AssSel = "E"
                     Case "其它設備"    AssSel = "F"
                     Case "租賃改良"    AssSel = "G"
              End Select

              If Request.QueryString("Check") <> "M" Then
                 If Request.QueryString("AssNumber") = "on" Then Str = Str & "(AssetsNumber LIKE '%-"& AssSel &"%') AND "
              Else
                  If Request.QueryString("AssNumber") = "on" Then Str = Str & "(AssetsNumber LIKE '%-Z"& AssSel &"%') AND " 
              End If

              If Request.QueryString("Obtain") = "on" Then
                 If Request.QueryString("ObtainEnd") = clear Then
                    Str = Str & "(ObtainDate >= #"& Request.QueryString("ObtainBegin") &"#) AND " 
                 Else
                     Str = Str & "(ObtainDate >= #"& Request.QueryString("ObtainBegin") &"#) AND (ObtainDate <= #"& Request.QueryString("ObtainEnd") &"#) AND "
                 End If
              End If

              If Request.QueryString("State") = "on" Then
                 Select Case Request.QueryString("StateSel") 
                        Case "投抵"  Str = Str & "(Insure = true) AND "
                        Case "抵押"  Str = Str & "(Mortgage = true) AND "
                        Case "租賃"  Str = Str & "(Iease = true) AND "
                        Case Else  Str = Str & "(NewState='"& Request.QueryString("StateSel") &"') AND "
                 End Select
              End If

              If Request.QueryString("FileNo") <> Clear Then
                 Str = Str & "(AssetsNumber LIKE '"& Ucase(Request.QueryString("FileNo")) &"%') AND " 
              End If

              Str = Str & "(AssetsNumber IS NOT NULL) AND (NewState = '正常') AND (SellCondition NOT LIKE '%報廢%' OR SellCondition IS NULL) "&_
                          "ORDER BY SUBSTRING(AssetsNumber, 4,LEN(AssetsNumber)),UseUnit"
         '******************************************************************************************************************************************************
         '修改資產登錄(列帳)、修改資產登錄(列管)
         Case "FixAssetsRegist", "PrintAssetsRegist"
              If Request.QueryString("Check") <> "M" Then
                 '列帳
                 If Request.QueryString("FileNumber") <> clear Then
                    Str = "SELECT * FROM AssetsDataBase WHERE (AssetsNumber = '"& UCASE(Request.QueryString("FileNumber")) &"')"
                 End If
              Else
                  
                  If Request.QueryString("FileNumber") <> clear Then
                     Str = "SELECT * FROM Assets2DataBase WHERE (AssetsNumber = '"& UCASE(Request.QueryString("FileNumber")) &"')"
                  End If  
              End If

         '******************************************************************************************************************************************************
         Case "DetailedList"
              if Request.QueryString("CheckC")="M" then
                 Str="select * from Assets2DataBase where "
              else
                  Str="select * from AssetsDataBase where "
              end if

              if Request.QueryString("Business")="on" then Str=Str & "BusinessDepartment='"& Request.QueryString("BusinessName") &"' And "
              if Request.QueryString("Unit")="on" then Str=Str & "UseUnit like '%"& Request.QueryString("UnitName") &"%' And " 

              select case Request.QueryString("AssSel")
                     case "土地"       AssSel="A"
                     case "房屋及建築" AssSel="B"
                     case "機器設備" AssSel="C"
                     case "運輸設備" AssSel="D"
                     case "辦公設備" AssSel="E"
                     case "其它設備" AssSel="F"
                     case "租賃改良" AssSel="G"
              end select

              if Request.QueryString("Check")<>"M" then
                 if Request.QueryString("AssNumber")="on" then Str=Str & "AssetsNumber LIKE '%-"& AssSel &"%' And "         
              else
                  if Request.QueryString("AssNumber")="on" then Str=Str & "AssetsNumber LIKE '%-Z"& AssSel &"%' And " 
              end if

              if Request.QueryString("Obtain")="on" then
                 if Request.QueryString("ObtainEnd")=clear then
                    Str=Str & "ObtainDate>='"& Request.QueryString("ObtainBegin") &"' And " 
                 else
                     Str=Str & "(ObtainDate>='"& Request.QueryString("ObtainBegin") &"' And ObtainDate<='"& Request.QueryString("ObtainEnd") &"') And "
                 end if
              end if 
              if Request.QueryString("State")="on" then
                 select case Request.QueryString("StateSel") 
                        case "投抵": Str=Str & "Insure = true And "
                        case "抵押": Str=Str & "Mortgage = true And "
                        case "租賃": Str=Str & "Iease = true And "
                        case else Str=Str & "NewState='"& Request.QueryString("StateSel") &"' And "                                                
                 end select
              end if                       
              Str=Str & "AssetsNumber IS NOT NULL order by AssetsNumber,UseUnit"      
         '******************************************************************************************************************************************************
  end select
  Set ManagRs=GF_GetRS(Str,ManagCn)
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<META HTTP-EQUIV="Pragma" CONTENT="NO-CACHE">
<title><%=Trim(Request.QueryString("Title"))%></title>
<style>
<!--
.size01 {  font-family: "Arial"; font-size: 7pt;  LETTER-SPACING: 0px; text-decoration: none}
.size02 {  font-family: "Arial"; font-size: 11pt; line-height: 16pt; LETTER-SPACING: 2px; text-decoration: none}
.size03 {  font-family: "Arial"; font-size: 13pt; line-height: 16pt; LETTER-SPACING: 2px;  text-decoration: none}
.size04 {  color:blue; font-family: "細明體", "Arial"; font-size: 13pt; line-height: 16pt; text-decoration: none}
.size05 {  font-family: "Arial"; font-size: 8pt;  LETTER-SPACING: 0px; text-decoration: none}
-->
</style>
<script language="javascript">
        function CheckInput()
        {
                 var Check=true;
                 if(form.AssName.value==''){alert('請填寫資產名稱.謝謝!!');form.AssName.focus();Check=false;}
                 if(form.Amount.value==''&&Check==true){alert('請填寫數量.謝謝!!');form.Amount.focus();Check=false;}
                 //if(form.ObtainDateF.value==''&&Check==true){alert('請填寫取得日期.謝謝!!');form.ObtainDateF.focus();Check=false;}
                 return Check;          
        }

        function CheckAssets(Box)
        {
                 var CheckText=CaseState='';
                 var Year=BusinessDept=Business=AssetsAssortment=Assortment=BUnit=Unit=Customer=Change=Estimate=Maintain=Sell=SellWay=ChangOut=ChangIn=Rectify='';
                 if(form.CusName.value!=''&&form.CusMan.value!=''&&form.CusTel.value!=''&&form.CusFax.value!=''&&form.CusAdd.value!='') Customer=form.CusName.value + '^' + form.CusMan.value + '^' + form.CusTel.value + '^' + form.CusFax.value + '^' + form.CusAdd.value;  

                 for(var e in form.ChangOut){if(e=='selectedIndex') ChangOut=form.ChangOut[form.ChangOut[e]].innerText}
                 for(var f in form.ChangIn){if(f=='selectedIndex') ChangIn=form.ChangIn[form.ChangIn[f]].innerText}
                 for(var BU in form.BelongUnit){if(BU=='selectedIndex') BUnit=form.BelongUnit[form.BelongUnit[BU]].innerText}
                 for(var d in form.UseUnit){if(d=='selectedIndex') Unit=form.UseUnit[form.UseUnit[d]].innerText}
                 if(form.ChangDate.value!=''&&form.ChangNumber.value!=''&&ChangOut!=''&&ChangIn!=''&&form.ChangState.value!='') Change=form.ChangDate.value + '^' + form.ChangNumber.value + '^' + ChangOut + '^' + ChangIn + '^' + form.ChangState.value;

                 if(form.EstDate.value!=''&&form.EstNumber.value!=''&&form.EstName.value!=''&&form.EstMoney.value!=''&&form.EstCust.value!='') Estimate=form.EstDate.value + '^' + form.EstNumber.value + '^' + form.EstName.value + '^' + form.EstMoney.value + '^' + form.EstCust.value;
                 if(form.MainDate.value!=''&&form.MainNumber.value!=''&&form.MainName.value!=''&&form.MainMoney.value!=''&&form.MainKeep.value!=''&&form.MainCust.value!='') Maintain=form.MainDate.value + '^' + form.MainNumber.value + '^' + form.MainName.value + '^' + form.MainMoney.value + '^' + form.MainKeep.value + '^' + form.MainCust.value;
                 if(form.RecDate.value!=''&&form.RecNumber.value!=''&&form.RecModel.value!=''&&form.RecMoney.value!=''&&form.RecApp.value!=''&&form.RecNext.value!='') Rectify=form.RecDate.value + '^' + form.RecNumber.value + '^' + form.RecModel.value + '^' + form.RecMoney.value + '^' + form.RecApp.value + '^' + form.RecNext.value;

                 for(var e in form.SellWay){if(e=='selectedIndex') SellWay=form.SellWay[form.SellWay[e]].innerText}
                 if(SellWay=='')SellWay='正常';
                 if(form.SellDate.value!=''&&form.SellNumber.value!=''&&form.SellReason.value!=''&&form.SellMoney.value!=''&&SellWay!='正常') Sell=form.SellDate.value + '^' + form.SellNumber.value + '^' + form.SellReason.value + '^' + SellWay + '^' + form.SellMoney.value;

                 var i,j,k; // i--Insure j--Mortgage k--Iease
                 if (form.Insure.checked==true)i=1;else i=0 ; 
                 if (form.Mortgage.checked==true)j=1;else j=0 ;
                 if (form.Iease.checked)k=1;else k=0 ;
                 CheckText+= Box.id + '|' + form.AssName.value + '|' + form.AssEut.value + '|' + form.AssNumber.value + '|' + form.Amount.value + '|' + form.Unit.value + '|' + form.BuyCost.value + '|' + form.TimeLimitV.value;
                 CheckText+='|' + Customer + '|' + Change + '|' + Estimate + '|' + Maintain + '|' + Sell + '|' + i + '|' + j + '|' + k + '|' + SellWay + '|' + form.KeepDate.value + '|' + form.Brand.value;
                 CheckText+='|' + BUnit + '|' + Unit + '|'+ Rectify + '|' + form.Remark.value;           
        
                 if(CheckInput()) window.location.replace('ManagUp.asp?CheckText='+CheckText+'&CheckManag=FixAssetsRegist&Check='+'<%=request.querystring("Check")%>');        
        }
</script>
<body background="0153.gif"
 onload="
         if('<%=ManagRs.eof%>'=='True'&'<%=ManagRs.bof%>'=='True'){
          alert('您所輸入的條件,目前並無資料,\n請重新輸入新的查詢時間.謝謝!!');
          window.location.replace('ManagCheck.asp?CheckManag='+'<%=Request.QueryString("CheckManag")%>'+'&Title='+'<%=Request.QueryString("Title")%>'+'&CheckC='+'<%=Request.QueryString("CheckC")%>')
         }
         else{
          <% if Not ManagRs.Eof And Not ManagRs.Bof then 
               select case Request.QueryString("CheckManag")
                  case "FixAssetsRegist":%>      
                      if('<%=ManagRs.fields("Insure")%>'=='True'){with(form){Insure.checked=Mortgage.disabled=Iease.disabled=SellDate.disabled=SellNumber.disabled=SellReason.disabled=SellWay.disabled=SellMoney.disabled=true;}}
                      if('<%=ManagRs.fields("Mortgage")%>'=='True'){with(form){Mortgage.checked=Insure.disabled=Iease.disabled=SellDate.disabled=SellNumber.disabled=SellReason.disabled=SellWay.disabled=SellMoney.disabled=true;}}
                      if('<%=ManagRs.fields("Iease")%>'=='True'){with(form){Iease.checked=Mortgage.disabled=form.Insure.disabled=SellDate.disabled=SellNumber.disabled=SellReason.disabled=SellWay.disabled=SellMoney.disabled=true;}}                     
                      //form.Business.disabled=form.UseUnit.disabled=form.Assortment.disabled=true;
                      //form.Business.disabled=form.UseUnit.disabled=form.Assortment.disabled=false;
          <%   end select
             end if %>
         }
        ">
<%'If Trim(request.QueryString("CheckManag")) <> "PrintMark" Then
  '   Response.Write("<center><table border=3 bgcolor=slateblue class=size03><th>SPORTON(資產管理程式)</table>" &_
  '                  "<table><td bgcolor=pink class=size02>" & Request.QueryString("Title") & "</table></center>")
  'End If%>

<%Select Case Request.QueryString("CheckManag")
         Case "CheckDetailedList": PageCountBuffer=20
         Case "PrintMark": PageCountBuffer=16 'PageCountBuffer = ManagRs.RecordCount
         Case Else PageCountBuffer=1000
  End Select

  If Not ManagRs.EOF And Not ManagRs.BOF Then
         Do Until ManagRs.EOF
                  CountT = CountT + 1
            ManagRs.movenext
         Loop
         ManagRs.MoveFirst
         session("Count") = Request.QueryString("Count")
         If Request.QueryString("Count") = clear Then session("Count") = 0
         Select Case Request.QueryString("Check")
                Case "N":
                     session("Count") = session("Count") + PageCountBuffer
                     if session("Count") > CountT Then session("Count") = session("Count") - PageCountBuffer
                Case "U":
                     session("Count") = Request.QueryString("Count") - PageCountBuffer
                     if session("Count") < PageCountBuffer Then session("Count") = 0
                Case "F":
                     session("Count") = 0
                Case "L":
                     buffer = CountT/PageCountBuffer
                     buffer = Int(buffer)
                     buffer = buffer * PageCountBuffer
                     session("Count") = buffer
         End Select

         ManagRs.Move session("Count")
         If session("Count") = 0 Then
            NumberCount = 1
         Else
             NumberCount = session("Count") + 1
         End If
         line = 1
         PageCount = session("Count")
         
         '*****************************************************************************************************************************************************
         Select Case Request.QueryString("CheckManag")
                '**********************************************************************************************************************************************
                Case "CheckDetailedList":
                     Response.Write("<center><form name=form><table width='100%' cellspacing=0 class=size05><td>列印日期：" & Date & "<td align=right>頁數：")

                     If PageCount = 0 Then
                        Response.write "1"
                     Else   
                         PageCount = PageCount+PageCountBuffer
                         If InStr(PageCount/PageCountBuffer, ".") <> 0 Then
                            Response.Write Int(PageCount/PageCountBuffer) + 1
                         Else
                             Response.write PageCount/PageCountBuffer
                         End If
                     End If
                     Response.Write "/"
                     If instr(CountT/PageCountBuffer, ".") <> 0 Then
                        Response.Write Int(CountT/PageCountBuffer) + 1
                     Else
                         Response.write CountT/PageCountBuffer
                     End If
                     Response.Write("</table><table border=1 cellspacing=0 cellpadding=3 align=center class=size05 bgcolor=skyblue width=100% >")
                '**********************************************************************************************************************************************
                Case "PrintAssetsRegist":
                      Response.Write("<center><form name=form><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >")
                '**********************************************************************************************************************************************
                Case "PrintMark"%>
                     <div id=Function1 style="position:absolute; top:0; right:10; background-color:blue; color:white; width:350; height:30; visibility:visible;">
                     <center>
                     <table width="100%" height="100%" bgcolor=dodgerblue cellspacing=0 border=1 class=size01>
                     <tr>
                         <td align=center onmouseover="this.className='DIVOn';" onmouseout="this.className='DIVOff';"><tt>頁數：
                         <%if PageCount=0 then
                              Response.write "1"
                           else
                               PageCount=PageCount+PageCountBuffer
                               if instr(PageCount/PageCountBuffer,".")<>0 then
                                  Response.Write int(PageCount/PageCountBuffer)+1
                               else
                                   Response.write PageCount/PageCountBuffer
                               end if
                           end if
                           Response.Write "/"
                           if instr(CountT/PageCountBuffer,".")<>0 then
                              Response.Write int(CountT/PageCountBuffer)+1
                           else
                               Response.write CountT/PageCountBuffer
                           end if%>
                           </tt>
                           <td align=center onmouseover="this.className='DIVOn';" onmouseout="this.className='DIVOff';"><tt>
                           <%Response.Write("<input type=button name=first value='F I R S T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                            "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'+'&Title='+" &_
                                            "'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=F&CheckC='+'"& Request.QueryString("CheckC") &"'" &_
                                            "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+" &_
                                            "'&Unit='+'"& Request.QueryString("Unit") &"'" &_
                                            "+'&UnitName='+'"& Request.QueryString("UnitName") &"')>" &_                                
                                            "<input type=button name=up value='U    P' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                            "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                            "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=U'" &_
                                            "+'&Business='+'"& Request.QueryString("Business") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'+" &_
                                            "'&BusinessName='+'"& Request.QueryString("BusinessName") &"'" &_
                                            "+'&Unit='+'"& Request.QueryString("Unit") &"'+'&UnitName='+'"& Request.QueryString("UnitName") &"')>" &_                   
                                            "<input type=button name=next value='N E X T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_                
                                            "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                            "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=N'" &_
                                            "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+" &_
                                            "'&Unit='+'"& Request.QueryString("Unit") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'" &_
                                            "+'&UnitName='+'"& Request.QueryString("UnitName") &"')>" &_    
                                            "<input type=button name=last value='L A S T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                            "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                            "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=L'" &_
                                            "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'+" &_            
                                            "'&Unit='+'"& Request.QueryString("Unit") &"'+'&UnitName='+'"& Request.QueryString("UnitName") &"')>")          
                           %>
                           </tt>
                           <td align=center onclick="Function1.style.visibility='hidden';" onmouseover="this.className='DIVOn';" onmouseout="this.className='DIVOff';"><tt>Hidden</tt>              
                           </table>
                           </center>
                           </div>
                           <%Response.Write("<center><form name=form>")
                '**********************************************************************************************************************************************
                Case Else
                         Response.Write("<center><form name=form>")                        
         End Select
         '*****************************************************************************************************************************************************

         '*****************************************************************************************************************************************************
         Select Case Request.QueryString("CheckManag") 
                Case "DetailedList":Response.Write("<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=skyblue class=size05><tr><td colspan=15><center>" & Request.QueryString("Title") & "</center>") 
                Case "CheckDetailedList":
                     Response.Write "<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=skyblue class=size05><tr><td colspan=10><center>" & Request.QueryString("Title") & "</center>" &_ 
                                    "<tr align=center><td bgcolor=yellow rowspan=2>保管單位<td bgcolor=yellow rowspan=2>資產編號<td bgcolor=yellow colspan=2>規格名稱" &_                        
                                    "<td bgcolor=yellow>取得日期<td bgcolor=yellow>數量"
                     If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                        Response.Write "<td bgcolor=yellow rowspan=2>購入成本"
                     End If
                     Response.Write "<td bgcolor=yellow rowspan=2>初盤" &_
                                    "<td bgcolor=yellow rowspan=2>複盤<td bgcolor=yellow rowspan=2>抽盤<tr align=center><td bgcolor=yellow>型號<td bgcolor=yellow>序號" &_
                                    "<td bgcolor=yellow>目前狀態<td bgcolor=yellow>單位"
         End Select
         '*****************************************************************************************************************************************************

         '*****************************************************************************************************************************************************
         Select Case Request.QueryString("CheckManag") 
                Case "CheckDetailedList" ModValue=20
                Case "DetailedList": ModValue=20:Response.Write "共" & CountT & "筆"    
         End Select

         '*****************************************************************************************************************************************************
         Select Case Request.QueryString("CheckManag")
                '**********************************************************************************************************************************************
                Case "PrintMark":
                     Do Until ManagRs.EOF OR line > PageCountBuffer
                              If ((line Mod 2) <> 0) Then
                                 MarkStr = MarkStr & "<br><table border=2 bgcolor=skyblue frame=void rules=none class=size02 width='100%' cellspacing=0>"
                              End If

                              If ((line Mod 2) = 1) Then
                                 MarkStr = MarkStr & "<tr>"
                              End If

                              If ((line Mod 2) = 0) Then
                                 MarkStr = MarkStr & "<td bgcolor=deepskyblue width='50%' height='100' valign=top>&nbsp;&nbsp;&nbsp;&nbsp;<tt><b>資產編號：</b></tt>" & ManagRs.fields("AssetsNumber") & "<br>" &_
                                                     "&nbsp;&nbsp;&nbsp;&nbsp;<tt><b>資產名稱：</b></tt><tt>" & ManagRs.fields("AssetsName") & "</tt><br>" &_
                                                     "&nbsp;&nbsp;&nbsp;&nbsp;<tt><b>保管單位：</b></tt><tt>" & ManagRs.fields("UseUnit") & "</tt>" &_
                                                     "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"& Trim(request.QueryString("PrintYear")) &"<br>"
                              Else
                                  MarkStr = MarkStr & "<td bgcolor=deepskyblue width='50%' height='100' valign=top><tt><b>資產編號：</b></tt>" & ManagRs.fields("AssetsNumber") & "<br>" &_
                                                      "<tt><b>資產名稱：</b></tt><tt>" & ManagRs.fields("AssetsName") & "</tt><br>" &_
                                                      "<tt><b>保管單位：</b></tt><tt>" & ManagRs.fields("UseUnit") & "</tt>" &_
                                                      "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"& Trim(request.QueryString("PrintYear")) &"<br>"
                              End If

                              If ((line Mod 2)=0) Then
                                 MarkStr = MarkStr & "</table><br>"
                              End If
                              Response.Write MarkStr
                              MarkStr = Clear
                        ManagRs.MoveNext
                        line = line + 1
                        NumberCount = NumberCount + 1 
                     Loop 
                '**********************************************************************************************************************************************
                Case "PrintAssetsRegist":        
                     Response.Write("<tr><td colspan=4 align=left class=size02><b>資產基本資料</b>" &_
                                    "<font class=size02><tt><b><u>(" & ManagRs.fields("AssetsNumber") & ")</font></table><br>" &_
                                    "<table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td align=center bgcolor=yellow width=10% >資產名稱<td align=left>" & Show_Str(ManagRs.fields("AssetsName")) &_
                                    "<td align=center bgcolor=yellow width=10% >型號<td align=left>" & Show_Str(ManagRs.fields("AssetsEut")) &_
                                    "<tr><td align=center bgcolor=yellow>序號<td align=left>" & Show_Str(ManagRs.fields("AssetsOrdinalNo")) &_
                                    "<td align=center bgcolor=yellow>數量<td align=left>" & Show_Str(ManagRs.fields("Amount")) &_
                                    "<tr><td align=center bgcolor=yellow>單位<td align=left>" & Show_Str(ManagRs.fields("Unit")) &_
                                    "<td align=center bgcolor=yellow>投抵/抵押/租賃<td align=left>")
                                
                     If ManagRs.fields("Insure") = True Then Response.Write("投抵") Else response.Write "　" End If
                     If ManagRs.fields("Mortgage") = True Then Response.Write("抵押") Else response.Write "　" End If
                     If ManagRs.fields("Iease") = True Then Response.Write("租賃") Else response.Write "　" End If

                     Response.Write("<tr><td align=center bgcolor=yellow>目前狀況")
                        
                     ManagRsBuffer = ManagRs.fields("SellCondition")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer,"`")
                        Count = Ubound(CaseStateBuffer1)
                        CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                        Response.Write("<td align=left>" & Show_Str(CaseStateBuffer(3))) 
                     Else
                         Response.Write "<td align=left>正常" 
                     End If

                     Response.Write("<td align=center bgcolor=yellow>取得日期<td align=left>" & Show_Str(ManagRs.fields("ObtainDate")) &_
                                    "<tr><td align=center bgcolor=yellow>廠牌<td align=left>" & Show_Str(ManagRs.fields("Brand")) &_
                                    "<td align=center bgcolor=yellow>購入成本<td align=left>" & Show_Str(ManagRs.fields("BuyCast")) &_
                                    "<tr><td align=center bgcolor=yellow>耐用年限<td align=left>" & Show_Str(ManagRs.fields("EnduringLimit")) &_
                                    "<td align=center bgcolor=yellow>保固期<td align=left>" & Show_Str(ManagRs.fields("KeepDate")) &_
                                    "<tr><td align=center bgcolor=yellow>事業部<td align=left>" & Show_Str(ManagRs.fields("BusinessDepartment")) &_
                                    "<td align=center bgcolor=yellow>保管單位<td align=left>" & Show_Str(ManagRs.fields("UseUnit")) &_ 
                                    "<tr><td align=center bgcolor=yellow>歸屬單位<td align=left>" & Show_Str(ManagRs.fields("BelongUnit")) &_
                                    "<td align=center bgcolor=yellow>資產分類<td align=left>" & Show_Str(ManagRs.fields("AssetsAssortment")) &_
                                    "<tr><td align=center bgcolor=yellow>備註<td align=left colspan=3>" & Show_Str(ManagRs.fields("Remark")) & "</table>" &_
                                    "<br><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left ><b>廠 商 基 本 資 料</b>")
                        
                     ManagRsBuffer = ManagRs.fields("CustomerBaseData")
                     If ManagRsBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=pink width=30 align=center>編號" &_
                                       "<td bgcolor=pink width=180 align=center>廠商名稱" &_
                                       "<td bgcolor=pink width=90 align=center>聯絡人員" &_
                                       "<td bgcolor=pink width=90 align=center>聯絡電話" &_
                                       "<td bgcolor=pink width=90 align=center>傳真電話" &_
                                       "<td bgcolor=pink align=center>住址")
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           NumberCount = 1
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td align=center>" & NumberCount &_
                                              "<td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4))
                               NumberCount = NumberCount + 1
                           Next
                        Else
                            CaseStateBuffer = Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4))                     
                        End If
                     Else 
                         Response.Write("-----尚無紀錄!!")
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left><b>異 動 情 形</b>")
                     ManagRsBuffer = ManagRs.fields("ChangesCondition")
                     If ManagRsBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=skyblue width=90 align=center>日期" &_
                                       "<td bgcolor=skyblue width=90 align=center>異動單號" &_
                                       "<td bgcolor=skyblue width=90 align=center>移出單位" &_
                                       "<td bgcolor=skyblue width=90 align=center>移入單位" &_
                                       "<td bgcolor=skyblue align=center>說明")
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer, "^")
                            If UBound(CaseStateBuffer) <> -1 Then 
                               Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4))
                            End If 
                        End If
                     Else
                          Response.Write("-----尚無紀錄!!")
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left><b>改 良 增 值 及 資 產 重 估</b>")
                     ManagRsBuffer = ManagRs.fields("EstimateCondition")
                     If ManagRsBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=skyblue width=90 align=center>日期" &_
                                       "<td bgcolor=skyblue width=90 align=center>改量單號" &_
                                       "<td bgcolor=skyblue align=center>規格名稱" &_
                                       "<td bgcolor=skyblue width=90 align=center>金額" &_
                                       "<td bgcolor=skyblue width=90 align=center>廠商")
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then
                               Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4))
                            End If
                        End If
                     Else
                         Response.Write("-----尚無紀錄!!")
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left><b>維 修 記 錄</b>")
                     ManagRsBuffer = ManagRs.fields("MainTainRecord")
                     If ManagRsBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=skyblue width=90 align=center>日期" &_
                                       "<td bgcolor=skyblue width=90 align=center>維修單號" &_
                                       "<td bgcolor=skyblue align=center>規格名稱" &_
                                       "<td bgcolor=skyblue width=90 align=center>金額" &_
                                       "<td bgcolor=skyblue width=90 align=center>保固期" &_
                                       "<td bgcolor=skyblue width=90 align=center>廠商")
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4) &_
                                              "<td align=center>" & CaseStateBuffer(5))
                           Next
                        Else
                            CaseStateBuffer = Split(ManagRsBuffer, "^")
                            If UBound(CaseStateBuffer) <> -1 Then 
                               Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4) &_ 
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(5))
                            End If
                        End If
                     Else
                         Response.Write("-----尚無紀錄!!")
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left><b>儀 器 校 正 紀 錄</b>")
                     RecBuffer=ManagRs.fields("RectifyRecord")
                     If RecBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=skyblue width=90 align=center>日期" &_
                                       "<td bgcolor=skyblue width=90 align=center>序號" &_
                                       "<td bgcolor=skyblue align=center>型號" &_
                                       "<td bgcolor=skyblue width=90 align=center>金額" &_
                                       "<td bgcolor=skyblue width=90 align=center>校驗廠商" &_
                                       "<td bgcolor=skyblue width=90 align=center>下次需檢驗日期")
                        CaseStateBuffer1 = Split(RecBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4) &_
                                              "<td align=center>" & CaseStateBuffer(5))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then 
                            Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) &_
                                       "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) &_
                                       "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) &_
                                       "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) &_
                                       "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4) &_
                                       "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(5))
                            End If
                        End If
                     Else
                         Response.Write("-----尚無紀錄!!")  
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table><table border=0 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                    "<tr><td colspan=6 align=left><b>出 售 / 報 廢 / 閑 置</b>")
                     ManagRs.MoveFirst
                     ManagRsBuffer = ManagRs.fields("SellCondition")
                     If ManagRsBuffer <> "" Then
                        Response.Write("<hr></table><table border=1 cellspacing=0 cellpadding=3 align=center class=size01 width=100% >" &_
                                       "<tr><td bgcolor=skyblue width=90 align=center>日期" &_
                                       "<td bgcolor=skyblue width=90 align=center>單號" &_
                                       "<td bgcolor=skyblue align=center>處份原因" &_
                                       "<td bgcolor=skyblue width=90 align=center>處份方式" &_
                                       "<td bgcolor=skyblue width=90 align=center>殘值")

                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td align=center>" & CaseStateBuffer(0) &_
                                              "<td align=center>" & CaseStateBuffer(1) &_
                                              "<td align=center>" & CaseStateBuffer(2) &_
                                              "<td align=center>" & CaseStateBuffer(3) &_
                                              "<td align=center>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then 
                               Response.Write("<tr><td bgcolor=pink width=180 align=center>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=pink width=180 align=center>" & CaseStateBuffer(4))
                            End If
                        End If
                     Else
                         Response.Write("-----尚無紀錄!!")  
                     End If
                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("</table></form></center></body></html>")
                '**********************************************************************************************************************************************
                Case "FixAssetsRegist":
                     Response.Write("<table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=skyblue frame=box class=size05><tr><td colspan=12><center>資 產 卡</center>" &_          
                                    "<tr><td align=center bgcolor=yellow>資產編號<td align=center bgcolor=yellow colspan=2>資產名稱<td align=center bgcolor=yellow colspan=2>型號<td align=center bgcolor=yellow>序號<td align=center bgcolor=yellow>數量<td align=center bgcolor=yellow>單位<td align=center bgcolor=yellow>投抵<td align=center bgcolor=yellow>抵押<td align=center bgcolor=yellow>租賃<td align=center bgcolor=yellow>目前狀況" &_
                                    "<tr><td align=center>" & ManagRs.fields("AssetsNumber") & "<td align=center colspan=2><input type=text name=AssName size=15 value='"& ManagRs.fields("AssetsName") &"' class=size05><td align=center colspan=2><input type=text size=12 name=AssEut value='"& ManagRs.fields("AssetsEut") &"' class=size05><td align=center><input type=text size=12 name=AssNumber value='"& ManagRs.fields("AssetsOrdinalNo") &"' class=size05><td align=center><input type=text size=5 name=Amount value='"& ManagRs.fields("Amount") &"' class=size05 disabled><td align=center><input type=text size=5 name=Unit value='"& ManagRs.fields("Unit") &"' class=size05 disabled><td align=center><input type=checkbox name=Insure class=size05 onclick=if(this.checked){form.Iease.disabled=form.Mortgage.disabled=true;}else{form.Iease.disabled=form.Mortgage.disabled=false;}><td align=center><input type=checkbox name=Mortgage class=size05 onclick=if(this.checked){form.Insure.disabled=form.Iease.disabled=true;}else{form.Insure.disabled=form.Iease.disabled=false;}><td align=center><input type=checkbox name=Iease onclick=if(this.checked){form.Insure.disabled=form.Mortgage.disabled=true;}else{form.Insure.disabled=form.Mortgage.disabled=false;}>")
                     ManagRsBuffer = ManagRs.fields("SellCondition")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        Count = UBound(CaseStateBuffer1)
                        CaseStateBuffer = Split(CaseStateBuffer1(Count),"^")
                        Response.Write("<td align=center>" & CaseStateBuffer(3)) 
                     Else
                         Response.Write "<td align=center>正常"
                     End If
                     Response.Write "<tr><td align=center bgcolor=yellow>取得日期<td align=center bgcolor=yellow>廠牌" &_
                                    "<td align=center bgcolor=yellow>購入成本<td align=center bgcolor=yellow>耐用年限" &_
                                    "<td align=center bgcolor=yellow>保固期<td align=center bgcolor=yellow>事業部" &_
                                    "<td align=center bgcolor=yellow>保管單位<td align=center bgcolor=yellow>歸屬單位" &_
                                    "<td align=center bgcolor=yellow colspan=3>資產分類<td align=center bgcolor=yellow>備註" &_
                                    "<tr><td align=center>"& ManagRs.fields("ObtainDate") &_
                                    "<td align=center><input type=text size=8 name=Brand value='"& ManagRs.fields("Brand") &"' class=size05>" &_
                                    "<td align=center><input type=text size=8 name=BuyCost value='"& ManagRs.fields("BuyCast") &"' class=size05>" &_
                                    "<td align=center><input type=text size=6 name=TimeLimitV value='"& ManagRs.fields("EnduringLimit") &"' class=size05>" &_
                                    "<td align=center><input type=text size=10 name=KeepDate value='"& ManagRs.fields("KeepDate") &"' class=size05>" &_
                                    "<td align=center> " & ManagRs.fields("BusinessDepartment") &_
                                    "<td align=center>"%>
                                    <select name=UseUnit class=size05>
                                    <option <%=Chk_Selected("認證一部華亞實驗室", ManagRs.fields("UseUnit"))%>>認證一部華亞實驗室
                                    <option <%=Chk_Selected("認證一部技術支援部", ManagRs.fields("UseUnit"))%>>認證一部技術支援部
                                    <option <%=Chk_Selected("認證一部研發部", ManagRs.fields("UseUnit"))%>>認證一部研發部
                                    <option <%=Chk_Selected("認證一部品管部", ManagRs.fields("UseUnit"))%>>認證一部品管部
                                    <option <%=Chk_Selected("認證一部文件部", ManagRs.fields("UseUnit"))%>>認證一部文件部
                                    <option <%=Chk_Selected("認證一部案件管理部", ManagRs.fields("UseUnit"))%>>認證一部案件管理部
                                    <option <%=Chk_Selected("認證一部安規工程部", ManagRs.fields("UseUnit"))%>>認證一部安規工程部
                                    <option <%=Chk_Selected("認證一部汐止工程部", ManagRs.fields("UseUnit"))%>>認證一部汐止工程部
                                    <option <%=Chk_Selected("認證一部東湖工程部", ManagRs.fields("UseUnit"))%>>認證一部東湖工程部
                                    <option <%=Chk_Selected("認證一部林口工程部", ManagRs.fields("UseUnit"))%>>認證一部林口工程部
                                    <option <%=Chk_Selected("認證一部EMC業務部", ManagRs.fields("UseUnit"))%>>認證一部EMC業務部
                                    <option <%=Chk_Selected("認證一部Safety業務部", ManagRs.fields("UseUnit"))%>>認證一部Safety業務部
                                    <option <%=Chk_Selected("認證一部新竹營運處業務部", ManagRs.fields("UseUnit"))%>>認證一部新竹營運處業務部
                                    <option <%=Chk_Selected("認證一部新竹營運處工程部", ManagRs.fields("UseUnit"))%>>認證一部新竹營運處工程部
                                    <option <%=Chk_Selected("認證二部業務部", ManagRs.fields("UseUnit"))%>>認證二部業務部
                                    <option <%=Chk_Selected("認證二部安規工程部", ManagRs.fields("UseUnit"))%>>認證二部安規工程部
                                    <option <%=Chk_Selected("認證二部EMC工程部", ManagRs.fields("UseUnit"))%>>認證二部EMC工程部
                                    <option <%=Chk_Selected("認證三部業務部", ManagRs.fields("UseUnit"))%>>認證三部業務部
                                    <option <%=Chk_Selected("認證三部工程組", ManagRs.fields("UseUnit"))%>>認證三部工程組
                                    <option <%=Chk_Selected("認證三部修改組", ManagRs.fields("UseUnit"))%>>認證三部修改組
                                    <option <%=Chk_Selected("認證三部文件組", ManagRs.fields("UseUnit"))%>>認證三部文件組
                                    <option <%=Chk_Selected("認證五部天線設計部", ManagRs.fields("UseUnit"))%>>認證五部天線設計部
                                    <option <%=Chk_Selected("認證五部RFI工程部", ManagRs.fields("UseUnit"))%>>認證五部RFI工程部
                                    <option <%=Chk_Selected("認證五部RF工程部", ManagRs.fields("UseUnit"))%>>認證五部RF工程部
                                    <option <%=Chk_Selected("零件事業部業務部", ManagRs.fields("UseUnit"))%>>零件事業部業務部
                                    <option <%=Chk_Selected("零件事業部資材部", ManagRs.fields("UseUnit"))%>>零件事業部資材部
                                    <option <%=Chk_Selected("總經理室-全球", ManagRs.fields("UseUnit"))%>>總經理室-全球
                                    <option <%=Chk_Selected("總經理室", ManagRs.fields("UseUnit"))%>>總經理室
                                    <option <%=Chk_Selected("財務部", ManagRs.fields("UseUnit"))%>>財務部
                                    <option <%=Chk_Selected("資訊部", ManagRs.fields("UseUnit"))%>>資訊部
                                    <option <%=Chk_Selected("客服部", ManagRs.fields("UseUnit"))%>>客服部
                                    <option <%=Chk_Selected("稽核室", ManagRs.fields("UseUnit"))%>>稽核室
                                    <option <%=Chk_Selected("公共行政部", ManagRs.fields("UseUnit"))%>>公共行政部
                                    </select>
                                    <td align=center>
                                    <select name=BelongUnit class=size05>
                                    <option <%=Chk_Selected("認證一部華亞實驗室", ManagRs.fields("BelongUnit"))%>>認證一部華亞實驗室
                                    <option <%=Chk_Selected("認證一部技術支援部", ManagRs.fields("BelongUnit"))%>>認證一部技術支援部
                                    <option <%=Chk_Selected("認證一部研發部", ManagRs.fields("BelongUnit"))%>>認證一部研發部
                                    <option <%=Chk_Selected("認證一部品管部", ManagRs.fields("BelongUnit"))%>>認證一部品管部
                                    <option <%=Chk_Selected("認證一部文件部", ManagRs.fields("BelongUnit"))%>>認證一部文件部
                                    <option <%=Chk_Selected("認證一部案件管理部", ManagRs.fields("BelongUnit"))%>>認證一部案件管理部
                                    <option <%=Chk_Selected("認證一部安規工程部", ManagRs.fields("BelongUnit"))%>>認證一部安規工程部
                                    <option <%=Chk_Selected("認證一部汐止工程部", ManagRs.fields("BelongUnit"))%>>認證一部汐止工程部
                                    <option <%=Chk_Selected("認證一部東湖工程部", ManagRs.fields("BelongUnit"))%>>認證一部東湖工程部
                                    <option <%=Chk_Selected("認證一部林口工程部", ManagRs.fields("BelongUnit"))%>>認證一部林口工程部
                                    <option <%=Chk_Selected("認證一部EMC業務部", ManagRs.fields("BelongUnit"))%>>認證一部EMC業務部
                                    <option <%=Chk_Selected("認證一部Safety業務部", ManagRs.fields("BelongUnit"))%>>認證一部Safety業務部
                                    <option <%=Chk_Selected("認證一部新竹營運處業務部", ManagRs.fields("BelongUnit"))%>>認證一部新竹營運處業務部
                                    <option <%=Chk_Selected("認證一部新竹營運處工程部", ManagRs.fields("BelongUnit"))%>>認證一部新竹營運處工程部
                                    <option <%=Chk_Selected("認證二部業務部", ManagRs.fields("BelongUnit"))%>>認證二部業務部
                                    <option <%=Chk_Selected("認證二部安規工程部", ManagRs.fields("BelongUnit"))%>>認證二部安規工程部
                                    <option <%=Chk_Selected("認證二部EMC工程部", ManagRs.fields("BelongUnit"))%>>認證二部EMC工程部
                                    <option <%=Chk_Selected("認證三部業務部", ManagRs.fields("BelongUnit"))%>>認證三部業務部
                                    <option <%=Chk_Selected("認證三部工程組", ManagRs.fields("BelongUnit"))%>>認證三部工程組
                                    <option <%=Chk_Selected("認證三部修改組", ManagRs.fields("BelongUnit"))%>>認證三部修改組
                                    <option <%=Chk_Selected("認證三部文件組", ManagRs.fields("BelongUnit"))%>>認證三部文件組
                                    <option <%=Chk_Selected("認證五部天線設計部", ManagRs.fields("BelongUnit"))%>>認證五部天線設計部
                                    <option <%=Chk_Selected("認證五部RFI工程部", ManagRs.fields("UseUnit"))%>>認證五部RFI工程部
                                    <option <%=Chk_Selected("認證五部RF工程部", ManagRs.fields("UseUnit"))%>>認證五部RF工程部
                                    <option <%=Chk_Selected("零件事業部業務部", ManagRs.fields("BelongUnit"))%>>零件事業部業務部
                                    <option <%=Chk_Selected("零件事業部資材部", ManagRs.fields("BelongUnit"))%>>零件事業部資材部
                                    <option <%=Chk_Selected("總經理室-全球", ManagRs.fields("BelongUnit"))%>>總經理室-全球
                                    <option <%=Chk_Selected("總經理室", ManagRs.fields("BelongUnit"))%>>總經理室
                                    <option <%=Chk_Selected("財務部", ManagRs.fields("BelongUnit"))%>>財務部
                                    <option <%=Chk_Selected("資訊部", ManagRs.fields("BelongUnit"))%>>資訊部
                                    <option <%=Chk_Selected("客服部", ManagRs.fields("BelongUnit"))%>>客服部
                                    <option <%=Chk_Selected("稽核室", ManagRs.fields("BelongUnit"))%>>稽核室
                                    <option <%=Chk_Selected("公共行政部", ManagRs.fields("BelongUnit"))%>>公共行政部
                                    </select>
                   <%response.Write "<td align=center colspan=3>" & ManagRs.fields("AssetsAssortment") &_
                                    "<td align=center>" &_
                                    "<input type=text size=8 name=Remark value='"& ManagRs.fields("Remark") &"' class=size05>" &_
                                    "</table><br><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size05 bgcolor=hotpink><tr>" &_
                                    "<td colspan=6 bgcolor=hotpink align=center style='color:snow'>廠商基本資料" &_
                                    "<tr><td bgcolor=pink align=center>廠商名稱<td bgcolor=pink align=center>聯絡人員" &_
                                    "<td bgcolor=pink align=center>聯絡電話<td bgcolor=pink align=center>傳真電話" &_
                                    "<td bgcolor=pink align=center>住址"

                     ManagRsBuffer=ManagRs.fields("CustomerBaseData")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1=Split(ManagRsBuffer,"`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("<tr><td><input type=text name=CusName size=25 class=size01><td><input type=text size=10 name=CusMan class=size05><td><input type=text size=12 name=CusTel class=size05><td><input type=text size=12 name=CusFax class=size05><td><input type=text size=55 name=CusAdd class=size05>" &_
                                    "</table><BR><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size05 bgcolor=steelblue><tr><td colspan=6 bgcolor=steelblue align=center style='color:snow'>異動情形" &_
                                    "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>異動單號<td bgcolor=skyblue align=center>移出單位<td bgcolor=skyblue align=center>移入單位<td bgcolor=skyblue align=center>說明")
                     ManagRsBuffer = ManagRs.fields("ChangesCondition")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer=Split(CaseStateBuffer1(Count),"^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write "<tr><td><input type=text name=ChangDate size=8 class=size05><td><input type=text size=8 name=ChangNumber class=size05>" &_
                                    "<td><select name=ChangOut class=size05>" &_
                                    "<option>" &_
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
                                    "<option>零件事業部業務部" &_
                                    "<option>零件事業部資材部" &_
                                    "<option>總經理室-全球" &_
                                    "<option>總經理室" &_
                                    "<option>財務部" &_
                                    "<option>資訊部" &_
                                    "<option>客服部" &_
                                    "<option>稽核室" &_
                                    "<option>公共行政部" &_
                                    "</select>" &_
                                    "<td><select name=ChangIn class=size05>" &_
                                    "<option>" &_
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
                                    "<option>零件事業部業務部" &_
                                    "<option>零件事業部資材部" &_
                                    "<option>總經理室-全球" &_
                                    "<option>總經理室" &_
                                    "<option>財務部" &_
                                    "<option>資訊部" &_
                                    "<option>客服部" &_
                                    "<option>稽核室" &_
                                    "<option>公共行政部" &_
                                    "</select><td><input type=text size=60 name=ChangState>" &_
                                    "</table><BR><table border=1 style=font-size:12 frame=box bgcolor=steelblue cellspacing=0 class=size05><tr><td colspan=6 bgcolor=steelblue align=center style='color:snow'>改良增值及資產重估" &_
                                    "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>改量單號<td bgcolor=skyblue align=center>規格名稱<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>廠商"
                     ManagRsBuffer = ManagRs.fields("EstimateCondition")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("<tr><td><input type=text name=EstDate size=8 class=size05><td><input type=text size=8 name=EstNumber class=size05><td><input type=text size=80 name=EstName class=size05><td><input type=text size=8 name=EstMoney class=size05><td><input type=text size=8 name=EstCust class=size05>" &_
                                    "</table><BR><table border=1 bgcolor=steelblue cellspacing=0 class=size05><tr><td colspan=6 bgcolor=steelblue align=center style='color:snow'>維修記錄" &_
                                    "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>維修單號<td bgcolor=skyblue align=center>規格名稱<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>保固期<td bgcolor=skyblue align=center>廠商")
                     ManagRsBuffer = ManagRs.fields("MainTainRecord")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(5))
                           Next
                        Else
                            CaseStateBuffer=Split(ManagRsBuffer,"^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4) & "<td bgcolor=pink align=center>" & CaseStateBuffer(5))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("<tr><td><input type=text name=MainDate size=8 class=size05><td><input type=text size=8 name=MainNumber class=size05><td><input type=text size=72 name=MainName class=size05><td><input type=text size=8 name=MainMoney class=size05><td><input type=text size=8 name=MainKeep class=size05><td><input type=text size=8 name=MainCust class=size05>" &_
                                    "</table><BR><table border=1 bgcolor=steelblue cellspacing=0 class=size05><tr><td colspan=6 bgcolor=steelblue align=center style='color:snow'>儀器校正紀錄" &_
                                    "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>序號<td bgcolor=skyblue align=center>型號<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>校驗廠商<td bgcolor=skyblue align=center>下次需檢驗日期")
                     RecBuffer = ManagRs.fields("RectifyRecord")
                     If RecBuffer <> "" Then
                        CaseStateBuffer1 = Split(RecBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(5))
                           Next
                        Else
                            CaseStateBuffer = Split(ManagRsBuffer, "^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4) & "<td bgcolor=pink align=center>" & CaseStateBuffer(5))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("<tr><td><input type=text name=RecDate size=8 class=size05><td><input type=text size=8 name=RecNumber class=size05><td><input type=text size=72 name=RecModel class=size05><td><input type=text size=8 name=RecMoney class=size05><td><input type=text size=8 name=RecApp class=size05><td><input type=text size=8 name=RecNext class=size05>" &_
                                    "</table><BR><table border=1 bgcolor=steelblue cellspacing=0 class=size05><tr><td colspan=6 bgcolor=steelblue align=center style='color:snow'>出售/報廢/閑置" &_
                                    "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>單號<td bgcolor=skyblue align=center>處份原因<td bgcolor=skyblue align=center>處份方式<td bgcolor=skyblue align=center>殘值")                  
                     ManagRs.MoveFirst
                     ManagRsBuffer = ManagRs.fields("SellCondition")
                     If ManagRsBuffer <> "" Then
                        CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                        If UBound(CaseStateBuffer1) <> -1 Then
                           For Count = 0 To UBound(CaseStateBuffer1)
                               CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                               Response.Write("<tr><td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(0) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(1) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(2) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(3) &_
                                              "<td bgcolor=blue align=center style='color:snow'>" & CaseStateBuffer(4))
                           Next
                        Else
                            CaseStateBuffer = Split(ManagRsBuffer, "^")
                            If UBound(CaseStateBuffer) <> -1 Then Response.Write("<tr><td bgcolor=pink align=center>" & CaseStateBuffer(0) & "<td bgcolor=pink align=center>" & CaseStateBuffer(1) & "<td bgcolor=pink align=center>" & CaseStateBuffer(2) & "<td bgcolor=pink align=center>" & CaseStateBuffer(3) & "<td bgcolor=pink align=center>" & CaseStateBuffer(4))
                        End If
                     End If

                     '------------------------------------------------------------------------------------------------------------------------------------------
                     Response.Write("<tr><td><input type=text name=SellDate size=8 class=size05><td><input type=text size=8 name=SellNumber class=size05><td><input type=text size=80 name=SellReason class=size05><td><select name=SellWay class=size05><option>出售<option>報廢<option>閑置<option>盤盈<option>盤虧<option>其它<option selected></select><td><input type=text size=8 name=SellMoney class=size05>" &_
                                    "</table><BR><input type=button name=AssetsEnter value=儲存上列輸入資料 id='"& ManagRs.fields("AssetsNumber") &"' onclick=CheckAssets(this); class=size05>" &_
                                    "</form></center></body></html>")

                '**********************************************************************************************************************************************
                Case "DetailedList","CheckDetailedList":
                     Do Until (ManagRs.EOF OR line > PageCountBuffer)
                              'set engrs=adocn.Execute("select * from Eng_Data where Test_FileNo = '"& adors.fields("File_No") &"'") 
                              If ((NumberCount-1) Mod ModValue) = 0 Then
                                 Select Case Request.QueryString("CheckManag")
                                        Case "DetailedList"
                                             Response.Write "<tr align=center><td bgcolor=yellow>編號<td bgcolor=yellow>保管單位<td bgcolor=yellow>資產編號<td bgcolor=yellow>資產名稱<td bgcolor=yellow>型號<td bgcolor=yellow>序號<td bgcolor=yellow>廠牌<td bgcolor=yellow>數量<td bgcolor=yellow>單位<td bgcolor=yellow>取得日期"
                                             If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                                                Response.Write "<td bgcolor=yellow>購入成本"
                                             End If
                                             Response.Write "<td bgcolor=yellow>投抵<td bgcolor=yellow>抵押<td bgcolor=yellow>租賃<td bgcolor=yellow>目前狀態"
                                 End Select
                              End If

                              '********************************************************************************************************************************
                              Select Case Request.QueryString("CheckManag")
                                     '*************************************************************************************************************************
                                     Case "CheckDetailedList"
                                          If i = 0 Then
                                             ManagRsBuffer = ManagRs.fields("SellCondition")
                                             If ManagRsBuffer <> "" Then
                                                CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                                                Count = Ubound(CaseStateBuffer1)
                                                CaseStateBuffer = Split(CaseStateBuffer1(Count), "^")
                                                StateBuffer = CaseStateBuffer(3)
                                             Else
                                                 StateBuffer = "正常"
                                             End If
                                             Response.Write "<tr><td bgcolor=deepskyblue rowspan=2>" & Show_Str(ManagRs.fields("UseUnit")) & "<td bgcolor=deepskyblue rowspan=2>" & Show_Str(ManagRs.fields("AssetsNumber")) &_
                                                            "<td bgcolor=deepskyblue colspan=2>" & Show_Str(ManagRs.fields("AssetsName")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("ObtainDate")) &_
                                                            "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("Amount"))
                                             If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                                                Response.Write "<td bgcolor=deepskyblue rowspan=2 align=right>" & Show_Str(ManagRs.fields("BuyCast"))
                                             End If
                                             Response.Write "<td bgcolor=deepskyblue rowspan=2>　" &_
                                                            "<td bgcolor=deepskyblue rowspan=2>　<td bgcolor=deepskyblue rowspan=2>　<tr><td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("AssetsEut")) & "<td bgcolor=deepskyblue>" &_
                                             Show_Str(ManagRs.fields("AssetsOrdinalNo")) & "<td bgcolor=deepskyblue>" &_ 
                                             Show_Str(StateBuffer) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("Unit"))

                                             i=1  
                                          Else 
                                              ManagRsBuffer = ManagRs.fields("SellCondition")
                                              If ManagRsBuffer <> "" Then
                                                 CaseStateBuffer1 = Split(ManagRsBuffer, "`")
                                                 Count = Ubound(CaseStateBuffer1)
                                                 CaseStateBuffer=Split(CaseStateBuffer1(Count), "^")
                                                 StateBuffer = CaseStateBuffer(3)
                                              Else
                                                  StateBuffer = "正常"
                                              End If
                                              Response.Write "<tr><td rowspan=2>" & Show_Str(ManagRs.fields("UseUnit")) & "<td rowspan=2>" & Show_Str(ManagRs.fields("AssetsNumber")) &_
                                                             "<td colspan=2>" & Show_Str(ManagRs.fields("AssetsName")) & "<td>" & Show_Str(ManagRs.fields("ObtainDate")) &_
                                                             "<td>" & Show_Str(ManagRs.fields("Amount"))
                                              If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                                                 Response.Write "<td rowspan=2 align=right>" & Show_Str(ManagRs.fields("BuyCast"))
                                              End If
                                              Response.Write "<td rowspan=2>　" &_
                                                             "<td rowspan=2>　<td rowspan=2>　<tr><td>" & Show_Str(ManagRs.fields("AssetsEut")) & "<td>" &_                                    
                                              Show_Str(ManagRs.fields("AssetsOrdinalNo")) & "<td>" &_                                    
                                              Show_Str(StateBuffer) & "<td>" & Show_Str(ManagRs.fields("Unit"))

                                              i=0  
                                          End If
                                     '*************************************************************************************************************************
                                     Case "DetailedList":
                                          If i = 0 Then
                                             Response.Write "<tr><td bgcolor=deepskyblue>" & NumberCount & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("UseUnit")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("AssetsNumber")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("AssetsName")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("AssetsEut")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("AssetsOrdinalNo")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("Brand")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("Amount")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("Unit")) & "<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("ObtainDate"))
                                             If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                                                response.Write "<td bgcolor=deepskyblue align=right>" & Show_Str(ManagRs.fields("BuyCast"))
                                             End If
                                             If ManagRs.fields("Insure") = True Then
                                                Response.Write("<td bgcolor=deepskyblue>Y")
                                             Else
                                                 Response.Write("<td bgcolor=deepskyblue>N")
                                             End If
                                             If ManagRs.fields("Mortgage") = True Then
                                                Response.Write("<td bgcolor=deepskyblue>Y")
                                             Else
                                                 Response.Write("<td bgcolor=deepskyblue>N")
                                             End If
                                             If ManagRs.fields("Iease") = True Then
                                                Response.Write("<td bgcolor=deepskyblue>Y")
                                             Else
                                                 Response.Write("<td bgcolor=deepskyblue>N")
                                             End If
                                             Response.Write("<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("NewState")))

                                             i=1  
                                          Else 
                                              Response.Write "<tr><td>" & NumberCount & "<td>" & Show_Str(ManagRs.fields("UseUnit")) & "<td>" & Show_Str(ManagRs.fields("AssetsNumber")) & "<td>" & Show_Str(ManagRs.fields("AssetsName")) & "<td>" & Show_Str(ManagRs.fields("AssetsEut")) & "<td>" & Show_Str(ManagRs.fields("AssetsOrdinalNo")) & "<td>" & Show_Str(ManagRs.fields("Brand")) & "<td>" & Show_Str(ManagRs.fields("Amount")) & "<td>" & Show_Str(ManagRs.fields("Unit")) & "<td>" & Show_Str(ManagRs.fields("ObtainDate"))
                                              If Session("CName") = "徐瓊芬" OR Session("CName") = "黃尹苗" Then
                                                 response.Write "<td align=right>" & Show_Str(Trim(ManagRs.fields("BuyCast")))
                                              End If
                                              If ManagRs.fields("Insure") = True Then
                                                 Response.Write("<td>Y")
                                              Else
                                                  Response.Write("<td>N")
                                              End If
                                              If ManagRs.fields("Mortgage") = True Then
                                                 Response.Write("<td>Y")
                                              Else
                                                  Response.Write("<td>N")
                                              End If
                                              If ManagRs.fields("Iease") = True Then
                                                 Response.Write("<td>Y")
                                              Else
                                                  Response.Write("<td>N")
                                              End If
                                              Response.Write("<td bgcolor=deepskyblue>" & Show_Str(ManagRs.fields("NewState")))
                                              i=0  
                                          End If
                              End Select
                              '********************************************************************************************************************************
                              ManagRs.MoveNext   
                              line = line+1   
                              NumberCount = NumberCount+1              
                              'set ManagRs=Nothing  
                     Loop                
                End Select
                '**********************************************************************************************************************************************
                Select Case Request.QueryString("CheckManag")    
                       Case "CheckDetailedList"
                            Response.Write("</table><br><table width=100% cellspacing=0 class=size05><td>抽盤人員:<td>複盤人員:<td>初盤人員:</table></form>" &_
                                           "<input type=button name=first value='F I R S T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                           "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'+'&Title='+" &_
                                           "'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=F&CheckC='+'"& Request.QueryString("CheckC") &"'" &_
                                           "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+" &_
                                           "'&Unit='+'"& Request.QueryString("Unit") &"'" &_
                                           "+'&UnitName='+'"& Request.QueryString("UnitName") &"')>　" &_                                
                                           "<input type=button name=up value='U    P' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                           "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                           "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=U'" &_
                                           "+'&Business='+'"& Request.QueryString("Business") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'+" &_
                                           "'&BusinessName='+'"& Request.QueryString("BusinessName") &"'" &_
                                           "+'&Unit='+'"& Request.QueryString("Unit") &"'+'&UnitName='+'"& Request.QueryString("UnitName") &"')>　" &_                   
                                           "<input type=button name=next value='N E X T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_                
                                           "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                           "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=N'" &_
                                           "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+" &_
                                           "'&Unit='+'"& Request.QueryString("Unit") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'" &_
                                           "+'&UnitName='+'"& Request.QueryString("UnitName") &"')>　" &_    
                                           "<input type=button name=last value='L A S T' onclick=window.location.replace('ManagShow.asp?Begin='+'"& request.querystring("Begin") &"'" &_
                                           "+'&End='+'"& request.querystring("End") &"'+'&Count='+'"& session("Count") &"'+'&CheckK='+'"& Request.QueryString("CheckK") &"'" &_
                                           "+'&Title='+'"& Request.QueryString("Title") &"'+'&CheckManag='+'"& Request.QueryString("CheckManag") &"'+'&Check=L'" &_
                                           "+'&Business='+'"& Request.QueryString("Business") &"'+'&BusinessName='+'"& Request.QueryString("BusinessName") &"'+'&CheckC='+'"& Request.QueryString("CheckC") &"'+" &_            
                                           "'&Unit='+'"& Request.QueryString("Unit") &"'+'&UnitName='+'"& Request.QueryString("UnitName") &"')></center></body></html>")
                End Select
                '**********************************************************************************************************************************************
  End If%>