<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->
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
</head>  
<script language="javascript">
        function CheckDate(Body)
        {
                 var NewDate=new Date();
                 var DateBuffer=Body.value.split('/')
                 if(DateBuffer.length<3){alert('日期尚未輸入或格式有問題\n正確的格式為--->(2000/12/1)');Body.value='';Body.focus();}
                 else{
                      if(DateBuffer[0]<1911){alert('年份不可小於西元1911年(民國元年).謝謝!!');Body.value='';Body.focus();} 
                      if((DateBuffer[1]<1||DateBuffer[1]>12)){alert('目前中華民國並無--'+DateBuffer[1]+'這個月份,請重新輸入.謝謝!!');Body.value='';Body.focus();}
                      if(DateBuffer[2]<1||DateBuffer[2]>31){alert('目前中華民國並無--'+DateBuffer[2]+'這個日期,請重新輸入.謝謝!!');Body.value='';Body.focus();}
                 }
        }

        function CheckInput()
        {
                 var Check=true;
                 if(form.AssName.value==''){alert('請填寫資產名稱.謝謝!!');form.AssName.focus();Check=false;}
                 if(form.Amount.value==''&&Check==true){alert('請填寫數量.謝謝!!');form.Amount.focus();Check=false;}
                 if(form.Unit.value==''&&Check==true){alert('請填寫單位.謝謝!!');form.Unit.focus();Check=false;}
                 if(form.ObtainDateF.value==''&&Check==true){alert('請填寫取得日期.謝謝!!');form.ObtainDateF.focus();Check=false;}
                 if(form.KeepDate.value==''&&Check==true){alert('請填寫保固日期.謝謝!!');form.KeepDate.focus();Check=false;}
                 if(form.BuyCost.value==''&&Check==true){alert('請填寫購入日期.謝謝!!');form.BuyCost.focus();Check=false;}
                 if(form.TimeLimit.value==''&&Check==true){alert('請填寫耐用年限.謝謝!!');form.TimeLimit.focus();Check=false;}        
                 return Check;          
        }


        function CheckAssets(Box)
        {
                 var CheckText=CaseState='';
                 var Year=BusinessDept=Business=AssetsAssortment=Assortment=BUnit=Unit=Customer=Change=Estimate=Maintain=Sell=SellWay=Rectify='';
                 
                 Year=form.ObtainDateF.value.substring(0,4)-1911;

                 for(var b in form.Business)
                 {
                     if(b=='selectedIndex') 
                     switch(BusinessDept=form.Business[form.Business[b]].innerText)
                     {
                            case "Component":Business='C';break;
                            case "EMC":Business='E';break;
                            case "Safety":Business='S';break;
                            case "BTL":Business='B';break;
                     }
                 }
        
                 for(var c in form.Assortment)
                 {
                     if(c=='selectedIndex') 
                     switch(AssetsAssortment=form.Assortment[form.Assortment[c]].innerText)
                     {
                            case "土地":Assortment='A';break;
                            case "房屋及建築":Assortment='B';break;
                            case "機器設備":Assortment='C';break;
                            case "運輸設備":Assortment='D';break;
                            case "辦公設備":Assortment='E';break;
                            case "其它設備":Assortment='F';break;
                            case "租賃改良":Assortment='G';break;
                     }
                 }

                 for(var d in form.UseUnit){if(d=='selectedIndex') Unit=form.UseUnit[form.UseUnit[d]].innerText}
                 for(var BU in form.BelongUnit){if(BU=='selectedIndex') BUnit=form.BelongUnit[form.BelongUnit[BU]].innerText}
                 for(var Out in form.ChangOut){if(Out=='selectedIndex') ChangOut=form.ChangOut[form.ChangOut[Out]].innerText}
                 for(var In in form.ChangIn){if(In=='selectedIndex') ChangIn=form.ChangIn[form.ChangIn[In]].innerText}

                 if(form.CusName.value!=''&&form.CusMan.value!=''&&form.CusTel.value!=''&&form.CusFax.value!=''&&form.CusAdd.value!='') Customer=form.CusName.value + '^' + form.CusMan.value + '^' + form.CusTel.value + '^' + form.CusFax.value + '^' + form.CusAdd.value;  
                 //if(form.ChangDate.value!=''&&form.ChangNumber.value!=''&&form.ChangOut.value!=''&&form.ChangIn.value!=''&&form.ChangState.value!='') Change=form.ChangDate.value + '^' + form.ChangNumber.value + '^' + form.ChangOut.value + '^' + form.ChangIn.value + '^' + form.ChangState.value;
                 if(form.ChangDate.value!=''&&form.ChangNumber.value!=''&& ChangOut !='' && ChangIn !='' && form.ChangState.value!='') Change=form.ChangDate.value + '^' + form.ChangNumber.value + '^' + ChangOut + '^' + ChangIn + '^' + form.ChangState.value;
                 
                 if(form.EstDate.value!=''&&form.EstNumber.value!=''&&form.EstName.value!=''&&form.EstMoney.value!=''&&form.EstCust.value!='') Estimate=form.EstDate.value + '^' + form.EstNumber.value + '^' + form.EstName.value + '^' + form.EstMoney.value + '^' + form.EstCust.value;
                 if(form.RecDate.value!=''&&form.RecNumber.value!=''&&form.RecModel.value!=''&&form.RecMoney.value!=''&&form.RecApp.value!=''&&form.RecNext.value!='') Rectify=form.RecDate.value + '^' + form.RecNumber.value + '^' + form.RecModel.value + '^' + form.RecMoney.value + '^' + form.RecApp.value + '^' + form.RecNext.value;
                 if(form.MainDate.value!=''&&form.MainNumber.value!=''&&form.MainName.value!=''&&form.MainMoney.value!=''&&form.MainKeep.value!=''&&form.MainCust.value!='') Maintain=form.MainDate.value + '^' + form.MainNumber.value + '^' + form.MainName.value + '^' + form.MainMoney.value + '^' + form.MainKeep.value + '^' + form.MainCust.value;
        
                 for(var e in form.SellWay){if(e=='selectedIndex') SellWay=form.SellWay[form.SellWay[e]].innerText}
                 if(SellWay=='')SellWay='正常'
        
                 if(form.SellDate.value!=''&&form.SellNumber.value!=''&&form.SellReason.value!=''&&form.SellMoney.value!=''&&SellWay=='正常') Sell=form.SellDate.value + '^' + form.SellNumber.value + '^' + form.SellReason.value + '^' + SellWay + '^' + form.SellMoney.value;
        
                 if(form.Insure.checked) blnInsure=1; else blnInsure=0 ;
                 if(form.Mortgage.checked) blnMortgage=1; else blnMortgage=0 ;
                 if(form.Iease.checked) blnIease=1; else blnIease=0 ;

                 CheckText+= Business + '|' + Assortment + '|' + Year + '|' + form.AssName.value + '|' + form.AssEut.value + '|' + form.AssNumber.value + '|' + form.Amount.value + '|' + form.Unit.value + '|' + form.ObtainDateF.value + '|' + form.BuyCost.value + '|' + form.TimeLimit.value;
                 CheckText+='|' + Unit + '|' + Customer + '|' + Change + '|' + Estimate + '|' + Maintain + '|' + Sell + '|' + blnInsure + '|' + blnMortgage + '|' + blnIease  + '|' + AssetsAssortment + '|' + BusinessDept + '|' + form.KeepDate.value + '|' + form.Remark.value + '|' + SellWay + '|' + form.Brand.value;
                 CheckText+='|' + BUnit + '|' + Rectify;

                 if(CheckInput()) window.location.replace('ManagUp.asp?CheckText='+CheckText+'&CheckManag=AssetsRegist&Check='+'<%=request.querystring("Check")%>');       
        }

</script>
<%Response.Write("</head><center><table border=3 bgcolor=slateblue class=size03><th><center>SPORTON(資產管理系統)</center></table>" &_
                 "</center><body background=0153.gif onload=form.ObtainDateF.focus();><center>" &_
                 "<form name=form><table border=1 cellspacing=0 cellpadding=3 align=center bgcolor=skyblue frame=box class=size01>" &_ 
                 "<tr><td colspan=10><center>" & Request.QueryString("Title") & "</center>" &_          
                 "<tr><td align=center bgcolor=yellow>取得日期<td align=center bgcolor=yellow>資產名稱" &_
                 "<td align=center bgcolor=yellow colspan=2>型號<td align=center bgcolor=yellow>序號" &_
                 "<td align=center bgcolor=yellow>數量<td align=center bgcolor=yellow>單位<td align=center bgcolor=yellow>投抵" &_
                 "<td align=center bgcolor=yellow>抵押<td align=center bgcolor=yellow>租賃" &_
                 "<tr><td align=center><input type=text size=8 name=ObtainDateF onblur=CheckDate(this); class=size01>" &_
                 "<td align=center><input type=text name=AssName size=15><td align=center colspan=2 class=size01>" &_
                 "<input type=text size=12 name=AssEut class=size01><td align=center><input type=text size=12 name=AssNumber class=size01>" &_
                 "<td align=center><input type=text size=1 maxlength=1 value=1 name=Amount class=size01 onkeypress='if ((event.keyCode < 48) || (event.keyCode > 57)) event.returnValue = false;' readonly><td align=center><input type=text size=5 name=Unit class=size01>" &_
                 "<td align=center><input type=checkbox name=Insure onclick=" &_
                 "if(this.checked){form.Iease.disabled=form.Mortgage.disabled=true;}" &_
                 "else{form.Iease.disabled=form.Mortgage.disabled=false;}>" &_
                 "<td align=center><input type=checkbox name=Mortgage onclick=" &_
                 "if(this.checked){form.Insure.disabled=form.Iease.disabled=true;}" &_
                 "else{form.Insure.disabled=form.Iease.disabled=false;}>" &_
                 "<td align=center><input type=checkbox name=Iease onclick=" &_
                 "if(this.checked){form.Insure.disabled=form.Mortgage.disabled=true;}" &_
                 "else{form.Insure.disabled=form.Mortgage.disabled=false;}>" &_
                 "<tr><td align=center bgcolor=yellow>保固期限<td align=center bgcolor=yellow>購入成本" &_
                 "<td align=center bgcolor=yellow>廠牌<td align=center bgcolor=yellow>耐用年限<td align=center bgcolor=yellow>事業部" &_
                 "<td align=center bgcolor=yellow>保管單位<td align=center bgcolor=yellow>歸屬單位" &_
                 "<td align=center bgcolor=yellow colspan=3>資產分類" &_   
                 "<tr><td align=center><input type=text size=8 name=KeepDate onblur=CheckDate(this); class=size01><td align=center>" &_
                 "<input type=text size=8 name=BuyCost class=size01><td align=center><input type=text name=Brand size=8><td align=center>" &_
                 "<input type=text size=3 name=TimeLimit class=size01><td align=center><select name=Business class=size01><option>Component" &_
                 "<option>EMC<option>Safety<option>WBU</select><td align=center>" &_
                 "<select name=UseUnit class=size01>" &_
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
                 "<option>總經理室-全球" &_
                 "<option>總經理室" &_
                 "<option>財務部" &_
                 "<option>資訊部" &_
                 "<option>客服部" &_
                 "<option>稽核室" &_
                 "<option>公共行政部" &_
                 "</select>" &_
                 "<td align=center>" &_
                 "<select name=BelongUnit class=size01>" &_
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
                 "<option>總經理室-全球" &_
                 "<option>總經理室" &_
                 "<option>財務部" &_
                 "<option>資訊部" &_
                 "<option>客服部" &_
                 "<option>稽核室" &_
                 "<option>公共行政部" &_
                 "</select>" &_
                 "<td align=center colspan=3><select name=Assortment class=size01>" &_
                 "<option>土地<option>房屋及建築<option>機器設備<option>運輸設備<option>辦公設備<option>其它設備<option>租賃改良" &_
                 "</select><tr><td colspan=10 bgcolor=yellow><center>備註欄</center>" &_
                 "<tr><td colspan=10 align=center><input type=text name=Remark size=80 class=size01>" &_
                 "</table><br><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=hotpink><tr><td colspan=6 bgcolor=hotpink align=center>廠商基本資料" &_
                 "<tr><td bgcolor=pink width=180 align=center>廠商名稱<td bgcolor=pink width=90 align=center>聯絡人員<td bgcolor=pink width=90 align=center>聯絡電話<td bgcolor=pink width=90 align=center>傳真電話<td bgcolor=pink align=center>住址" &_                  
                 "<tr><td><input type=text name=CusName size=25 class=size01><td><input type=text size=10 name=CusMan class=size01><td><input type=text size=12 name=CusTel class=size01><td><input type=text size=12 name=CusFax class=size01><td><input type=text size=55 name=CusAdd class=size01>" &_
                 "</table><br><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=deepskyblue><tr><td colspan=6 bgcolor=deepskyblue align=center>異動情形" &_
                 "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>異動單號<td bgcolor=skyblue align=center>移出單位<td bgcolor=skyblue align=center>移入單位<td bgcolor=skyblue align=center>說明" &_   
                 "<tr><td><input type=text name=ChangDate size=8 class=size01><td><input type=text size=8 name=ChangNumber class=size01>" &_
                 "<td>" &_
                 "<select name=ChangOut class=size01>" &_
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
                 "<option>總經理室" &_
                 "<option>財務部" &_
                 "<option>資訊部" &_
                 "<option>客服部" &_
                 "<option>稽核室" &_
                 "<option>公共行政部" &_
                 "</select>" &_
                 "<td>" &_
                 "<select name=ChangIn class=size01>" &_
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
                 "<option>總經理室" &_
                 "<option>財務部" &_
                 "<option>資訊部" &_
                 "<option>客服部" &_
                 "<option>稽核室" &_
                 "<option>公共行政部" &_
                 "</select>" &_
                 "<td><input type=text size=55 name=ChangState class=size01>" &_
                 "</table><br><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=deepskyblue><tr><td colspan=6 bgcolor=deepskyblue align=center>改良增值及資產重估" &_
                 "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>改量單號<td bgcolor=skyblue align=center>規格名稱<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>廠商" &_                  
                 "<tr><td><input type=text name=EstDate size=8 class=size01><td><input type=text size=8 name=EstNumber class=size01><td><input type=text size=80 name=EstName class=size01><td><input type=text size=8 name=EstMoney class=size01><td><input type=text size=8 name=EstCust class=size01>" &_
                 "</table><br><table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=deepskyblue><tr><td colspan=6 bgcolor=deepskyblue align=center>維修記錄" &_
                 "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>維修單號<td bgcolor=skyblue align=center>規格名稱<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>保固期<td bgcolor=skyblue align=center>廠商" &_                  
                 "<tr><td><input type=text name=MainDate size=8 class=size01><td><input type=text size=8 name=MainNumber class=size01><td><input type=text size=72 name=MainName class=size01><td><input type=text size=8 name=MainMoney class=size01><td><input type=text size=8 name=MainKeep class=size01><td><input type=text size=8 name=MainCust class=size01>" &_
                 "</table><br>" &_
                 "<table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=deepskyblue><tr><td colspan=6 bgcolor=deepskyblue align=center>儀器校正紀錄" &_
                 "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>序號<td bgcolor=skyblue align=center>型號<td bgcolor=skyblue align=center>金額<td bgcolor=skyblue align=center>校驗廠商<td bgcolor=skyblue align=center>下次需檢驗日期" &_                  
                 "<tr><td><input type=text name=RecDate size=8 class=size01><td><input type=text size=8 name=RecNumber class=size01><td><input type=text size=72 name=RecModel class=size01><td><input type=text size=8 name=RecMoney class=size01><td><input type=text size=8 name=RecApp class=size01><td><input type=text size=8 name=RecNext class=size01>" &_
                 "</table><br>" &_
                 "<table border=1 cellspacing=0 cellpadding=3 align=center frame=box class=size01 bgcolor=deepskyblue><tr><td colspan=6 bgcolor=deepskyblue align=center>出售/報廢/閑置" &_
                 "<tr><td bgcolor=skyblue align=center>日期<td bgcolor=skyblue align=center>單號<td bgcolor=skyblue align=center>處份原因<td bgcolor=skyblue align=center>處份方式<td bgcolor=skyblue align=center>殘值" &_                  
                 "<tr><td><input type=text name=SellDate size=8 class=size01><td><input type=text size=8 name=SellNumber class=size01><td><input type=text size=80 name=SellReason class=size01><td><select name=SellWay class=size01><option>出售<option>報廢<option>閑置<option>盤盈<option>盤虧<option selected></select><td><input type=text size=8 name=SellMoney class=size01>" &_                 
                 "</table></form><input type=button name=AssetsEnter value=儲存上列輸入資料 onclick=CheckAssets(this); class=size01>" &_
                 "</center></body></html>")            
%>