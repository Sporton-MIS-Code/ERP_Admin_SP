<!-- #INCLUDE FILE="../Tools/Function.asp" -->
<!-- #INCLUDE FILE="Check_Login.inc" -->

<%Response.Buffer = True
  Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

  '******************************************************************************************************************************************************
  Select Case Trim(Request.QueryString("CheckManag"))
         '***********************************************************************************************************************************************
         '新增資產登錄(列帳)、新增資產登錄(列管)
         Case "AssetsRegist"
              CheckText = Trim(Request.QueryString("CheckText"))
              AssetsFile = Split(CheckText, "|")
              '***********************************************************************************************************************************************
              If Request.QueryString("Check") <> "M" Then
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 '列帳
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 'AssetsFile(1)為資產分類(土地：A，房屋及建築：B，機器設備：C，運輸設備：D，辦公設備：E，其它設備：F，租賃改良：G)
                 'AssetsFile(2)為年度(如:94年)
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 Str1 = "SELECT ISNULL(MAX(SUBSTRING(AssetsNumber, 7, 5)), '') AS Number " &_
                        "FROM AssetsDataBase " &_
                        "WHERE (SUBSTRING(AssetsNumber, 3, 4) = '"& AssetsFile(1) & AssetsFile(2) &"')"
                 Set rs1 = GF_GetRS(Str1, ManagCn)
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 Number = Trim(rs1("Number"))
                 If Number = "" Then
                    AssetsNumberBuffer = AssetsFile(0) & "-" & AssetsFile(1) & AssetsFile(2) & "00001"
                 Else
                     Number1 = CStr(CLng(Number) +1)
                     AssetsNumberBuffer = AssetsFile(0) & "-" & AssetsFile(1) & AssetsFile(2) & Right(("00000" & Number1), 5)
                 End If
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 Str2 = "INSERT INTO AssetsDataBase(" &_
                        "AssetsNumber, AssetsName, AssetsEut, AssetsOrdinalNo, Amount, Unit, ObtainDate, BuyCast, EnduringLimit, UseUnit, " &_
                        "CustomerBaseData, ChangesCondition, EstimateCondition, MaintainRecord, SellCondition, Insure, Mortgage, Iease, " &_
                        "AssetsAssortment, BusinessDepartment, KeepDate, Remark, NewState, Brand, BelongUnit, RectifyRecord) " &_
                        "VALUES('"& AssetsNumberBuffer &"', '"& AssetsFile(3) &"', '"& AssetsFile(4) &"', '"& AssetsFile(5) &"', "& AssetsFile(6) &", " &_
                        "'"& AssetsFile(7) &"', '"& AssetsFile(8) &"', '"& AssetsFile(9) &"', '"& AssetsFile(10) &"', '"& AssetsFile(11) &"', " &_
                        "'"& AssetsFile(12) &"', '"& AssetsFile(13) &"', '"& AssetsFile(14) &"', '"& AssetsFile(15) &"', '"& AssetsFile(16) &"', " &_
                        AssetsFile(17) & ", "& AssetsFile(18) & ", " & AssetsFile(19) & ", '"& AssetsFile(20) &"', '"& AssetsFile(21) &"', " &_
                        "'"& AssetsFile(22) &"','"& AssetsFile(23) &"','"& AssetsFile(24) &"','"& AssetsFile(25) &"','"& AssetsFile(26) &"','"& AssetsFile(27) &"')"
                 ManagCn.Execute(Str2)
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 Str3 = "SELECT * FROM AssetsDataBase WHERE (AssetsNumber = '"& AssetsNumberBuffer &"')"
                 Set rs3 = GF_GetRS(Str3,ManagCn)
                 '---------------------------------------------------------------------------------------------------------------------------------------
                 'Set TestRs = GF_GetRS("select Max(AssetsNumber) from AssetsDataBase where AssetsNumber LIKE '"& AssetsFile(0) & "-" & AssetsFile(1) & AssetsFile(2) &"%'",ManagCn)    
                 'if isnull(TestRs.Fields(0))=true then  
                 '   if Request.QueryString("Check")<>"M" then  
                 '      AssetsNumberBuffer=AssetsFile(0) & "-" & AssetsFile(1) & AssetsFile(2) & "000001"        
                 '   else
                 '       AssetsNumberBuffer=AssetsFile(0) & "-Z" & AssetsFile(1) & AssetsFile(2) & "000001"
                 '   end if
                 'else 
                 '    select case len(mid(TestRs.Fields(0),6,6)+1)
                 '           case "1":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & "00000" & mid(TestRs.Fields(0),6,6)+1 
                 '           case "2":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & "0000" & mid(TestRs.Fields(0),6,6)+1
                 '           case "3":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & "000" & mid(TestRs.Fields(0),6,6)+1
                 '           case "4":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & "00" & mid(TestRs.Fields(0),6,6)+1
                 '           case "5":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & "0" & mid(TestRs.Fields(0),6,6)+1
                 '           case "6":AssetsNumberStr=Mid(TestRs.Fields(0),1,5) & mid(TestRs.Fields(0),6,6)+1
                 '    end select                  
                 '    AssetsNumberBuffer=AssetsNumberStr                 
                 'end if                                           
     
                 'set ControlRs=Server.CreateObject("adodb.recordset")
     
                 'if AssetsFile(6)>1 then         
                 '   for i=1 to AssetsFile(6)
                 '       AssetsNumberBuffer1=AssetsNumberBuffer & "-" & i
                        
                 '       AssetsString="Insert into AssetsDataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit,CustomerBaseData,ChangesCondition,EstimateCondition,MaintainRecord,SellCondition,Insure,Mortgage,Iease,AssetsAssortment,BusinessDepartment,KeepDate,Remark,NewState,Brand,BelongUnit,RectifyRecord) "
                 '       AssetsString=AssetsString & "values ('"& AssetsNumberBuffer1 &"','"& AssetsFile(3) &"','"& AssetsFile(4) &"','"& AssetsFile(5) &"',"& AssetsFile(6) &",'"& AssetsFile(7) &"','"& AssetsFile(8) &"','"& AssetsFile(9) &"','"& AssetsFile(10) &"','"& AssetsFile(11) &"' "
                 '       AssetsString=AssetsString & ",'"& AssetsFile(12) &"','"& AssetsFile(13) &"','"& AssetsFile(14) &"','"& AssetsFile(15) &"','"& AssetsFile(16) &"',"& AssetsFile(17) &","& AssetsFile(18) &","& AssetsFile(19) &",'"& AssetsFile(20) &"','"& AssetsFile(21) &"','"& AssetsFile(22) &"','"& AssetsFile(23) &"','"& AssetsFile(24) &"','"& AssetsFile(25) &"','"& AssetsFile(26) &"','"& AssetsFile(27) &"')"
                 '       ControlRs.Open AssetsString,ManagCn,3,3 
                 '       response.write(err.Description)
                 '   next
                 'else
                 '    AssetsString="Insert into AssetsDataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit,CustomerBaseData,ChangesCondition,EstimateCondition,MaintainRecord,SellCondition,Insure,Mortgage,Iease,AssetsAssortment,BusinessDepartment,KeepDate,Remark,NewState,Brand,BelongUnit,RectifyRecord) "
                 '    AssetsString=AssetsString & "values ('"& AssetsNumberBuffer &"','"& AssetsFile(3) &"','"& AssetsFile(4) &"','"& AssetsFile(5) &"',"& AssetsFile(6) &",'"& AssetsFile(7) &"','"& AssetsFile(8) &"','"& AssetsFile(9) &"','"& AssetsFile(10) &"','"& AssetsFile(11) &"' "
                 '    AssetsString=AssetsString & ",'"& AssetsFile(12) &"','"& AssetsFile(13) &"','"& AssetsFile(14) &"','"& AssetsFile(15) &"','"& AssetsFile(16) &"',"& AssetsFile(17) &","& AssetsFile(18) &","& AssetsFile(19) &",'"& AssetsFile(20) &"','"& AssetsFile(21) &"','"& AssetsFile(22) &"','"& AssetsFile(23) &"','"& AssetsFile(24) &"','"& AssetsFile(25) &"','"& AssetsFile(26) &"','"& AssetsFile(27) &"')"
                 '    ControlRs.Open AssetsString,ManagCn,3,3 
                 'end if        
                 'Set TestRs=GF_GetRS("select * from AssetsDataBase where AssetsNumber LIKE '%"& AssetsNumberBuffer &"%'",ManagCn)
              '***********************************************************************************************************************************************
              Else
                  '列管
                  '--------------------------------------------------------------------------------------------------------------------------------------
                  'AssetsFile(1)為資產分類(土地：A，房屋及建築：B，機器設備：C，運輸設備：D，辦公設備：E，其它設備：F，租賃改良：G)
                  'AssetsFile(2)為年度(如:94年)
                  '--------------------------------------------------------------------------------------------------------------------------------------
                  Str1 = "SELECT ISNULL(MAX(SUBSTRING(AssetsNumber, 8, 5)), '') AS Number " &_
                         "FROM Assets2DataBase " &_
                         "WHERE (SUBSTRING(AssetsNumber, 4, 4) = '"& AssetsFile(1) & AssetsFile(2) &"')"
                  Set rs1 = GF_GetRS(Str1, ManagCn)
                 '---------------------------------------------------------------------------------------------------------------------------------------
                  Number = Trim(rs1("Number"))
                  If Number = "" Then
                     AssetsNumberBuffer = AssetsFile(0) & "-Z" & AssetsFile(1) & AssetsFile(2) & "00001"
                  Else
                      Number1 = CStr(CLng(Number) +1)
                      AssetsNumberBuffer = AssetsFile(0) & "-Z" & AssetsFile(1) & AssetsFile(2) & Right(("00000" & Number1), 5)
                  End If
                  '---------------------------------------------------------------------------------------------------------------------------------------
                  Str2 = "INSERT INTO Assets2DataBase( " &_
                         "AssetsNumber, AssetsName, AssetsEut, AssetsOrdinalNo, Amount, Unit, ObtainDate, BuyCast, EnduringLimit, UseUnit, CustomerBaseData, " &_
                         "ChangesCondition, EstimateCondition, MaintainRecord, SellCondition, Insure, Mortgage, Iease, AssetsAssortment, BusinessDepartment, " &_
                         "KeepDate,Remark,NewState,Brand,BelongUnit,RectifyRecord) " &_
                         "VALUES ('"& AssetsNumberBuffer &"', '"& AssetsFile(3) &"', '"& AssetsFile(4) &"', '"& AssetsFile(5) &"', "& AssetsFile(6) &", " &_
                         "'"& AssetsFile(7) &"', '"& AssetsFile(8) &"', '"& AssetsFile(9) &"', '"& AssetsFile(10) &"', '"& AssetsFile(11) &"', " &_
                         "'"& AssetsFile(12) &"', '"& AssetsFile(13) &"', '"& AssetsFile(14) &"', '"& AssetsFile(15) &"', '"& AssetsFile(16) &"', " &_
                         AssetsFile(17) &", "& AssetsFile(18) &", "& AssetsFile(19) &", '"& AssetsFile(20) &"', '"& AssetsFile(21) &"', '"& AssetsFile(22) &"', " &_
                         "'"& AssetsFile(23) &"', '"& AssetsFile(24) &"', '"& AssetsFile(25) &"', '"& AssetsFile(26) &"', '"& AssetsFile(27) &"')"
                  ManagCn.Execute(Str2)
                  '---------------------------------------------------------------------------------------------------------------------------------------
                  Str3 = "SELECT * FROM Assets2DataBase WHERE (AssetsNumber = '"& AssetsNumberBuffer &"')"
                  Set rs3 = GF_GetRS(Str3,ManagCn)
                  '---------------------------------------------------------------------------------------------------------------------------------------
                  'Set TestRs=GF_GetRS("select Max(AssetsNumber) from Assets2DataBase where AssetsNumber LIKE '"& AssetsFile(0) & "-Z" & AssetsFile(1) & AssetsFile(2) &"%'",ManagCn)
                  'if isnull(TestRs.Fields(0))=true then                 
                  '   AssetsNumberBuffer=AssetsFile(0) & "-Z" & AssetsFile(1) & AssetsFile(2) & "000001"
                  'else 
                  '    select case len(mid(TestRs.Fields(0),7,6)+1)
                  '           case "1":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & "00000" & mid(TestRs.Fields(0),7,6)+1 
                  '           case "2":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & "0000" & mid(TestRs.Fields(0),7,6)+1
                  '           case "3":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & "000" & mid(TestRs.Fields(0),7,6)+1
                  '           case "4":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & "00" & mid(TestRs.Fields(0),7,6)+1
                  '           case "5":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & "0" & mid(TestRs.Fields(0),7,6)+1
                  '           case "6":AssetsNumberStr=Mid(TestRs.Fields(0),1,6) & mid(TestRs.Fields(0),7,6)+1
                  '    end select                  

                  '    AssetsNumberBuffer=AssetsNumberStr                 
                  'end if                                           
                  'set ControlRs=Server.CreateObject("adodb.recordset")       
                  'if AssetsFile(6)>1 then         
                  '   for i=1 to AssetsFile(6)          
                  '       AssetsNumberBuffer1=AssetsNumberBuffer & "-" & i

                  '       AssetsString="Insert into Assets2DataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit,CustomerBaseData,ChangesCondition,EstimateCondition,MaintainRecord,SellCondition,Insure,Mortgage,Iease,AssetsAssortment,BusinessDepartment,KeepDate,Remark,NewState,Brand,BelongUnit,RectifyRecord) "
                  '       AssetsString=AssetsString & "values ('"& AssetsNumberBuffer1 &"','"& AssetsFile(3) &"','"& AssetsFile(4) &"','"& AssetsFile(5) &"',"& AssetsFile(6) &",'"& AssetsFile(7) &"','"& AssetsFile(8) &"','"& AssetsFile(9) &"','"& AssetsFile(10) &"','"& AssetsFile(11) &"' "
                  '       AssetsString=AssetsString & ",'"& AssetsFile(12) &"','"& AssetsFile(13) &"','"& AssetsFile(14) &"','"& AssetsFile(15) &"','"& AssetsFile(16) &"',"& AssetsFile(17) &","& AssetsFile(18) &","& AssetsFile(19) &",'"& AssetsFile(20) &"','"& AssetsFile(21) &"','"& AssetsFile(22) &"','"& AssetsFile(23) &"','"& AssetsFile(24) &"','"& AssetsFile(25) &"','"& AssetsFile(26) &"','"& AssetsFile(27) &"')"
                  '       ControlRs.Open AssetsString,ManagCn,3,3 
                  '   next
                  'else
                  '    AssetsString="Insert into Assets2DataBase(AssetsNumber,AssetsName,AssetsEut,AssetsOrdinalNo,Amount,Unit,ObtainDate,BuyCast,EnduringLimit,UseUnit,CustomerBaseData,ChangesCondition,EstimateCondition,MaintainRecord,SellCondition,Insure,Mortgage,Iease,AssetsAssortment,BusinessDepartment,KeepDate,Remark,NewState,Brand,BelongUnit,RectifyRecord) "
                  '    AssetsString=AssetsString & "values ('"& AssetsNumberBuffer &"','"& AssetsFile(3) &"','"& AssetsFile(4) &"','"& AssetsFile(5) &"',"& AssetsFile(6) &",'"& AssetsFile(7) &"','"& AssetsFile(8) &"','"& AssetsFile(9) &"','"& AssetsFile(10) &"','"& AssetsFile(11) &"' "
                  '    AssetsString=AssetsString & ",'"& AssetsFile(12) &"','"& AssetsFile(13) &"','"& AssetsFile(14) &"','"& AssetsFile(15) &"','"& AssetsFile(16) &"',"& AssetsFile(17) &","& AssetsFile(18) &","& AssetsFile(19) &",'"& AssetsFile(20) &"','"& AssetsFile(21) &"','"& AssetsFile(22) &"','"& AssetsFile(23) &"','"& AssetsFile(24) &"','"& AssetsFile(25) &"','"& AssetsFile(26) &"','"& AssetsFile(27) &"')"
                  '    ControlRs.Open AssetsString,ManagCn,3,3 
                  'end if                      
                  'Set TestRs=GF_GetRS("select * from Assets2DataBase where AssetsNumber LIKE '%"& AssetsNumberBuffer &"%'",ManagCn)    
              End If
              '***********************************************************************************************************************************************
              If Not rs3.EOF Then
                 RS3_AssetsNumber = Trim(rs3("AssetsNumber"))
                 Response.Write("<html><head></head><body bgcolor=skyblue onload=alert('資產編號：'+'"& RS3_AssetsNumber &"');window.location.replace('ManagMainMenu.asp');></body></html>")
              Else
                  Response.Write("<html><head></head><body bgcolor=skyblue onload=alert('新增失敗！');window.location.replace('ManagMainMenu.asp');></body></html>")
              End If

              'do until TestRs.EOF                     
              '         Text=Text & TestRs.Fields(0) & ".\n" 
              '         TestRs.MoveNext
              'loop
              'Response.Write("<html><head></head><body bgcolor=skyblue onload=alert('資產編號:\n'+'"& Text &"');window.location.replace('ManagMainMenu.asp');></body></html>")
       
              'Set ControlRs=Nothing                                                                                          

         '******************************************************************************************************************************************************
         '修改資產登錄(列帳)、修改資產登錄(列管)
         Case "FixAssetsRegist"
              FixAssetsFile = Split(Request.QueryString("CheckText"),"|")

              Set FixTestRs = Server.CreateObject("ADODB.RecordSet")

              If Request.QueryString("Check") <> "M" Then
                 '列帳
                 FixTestRs.Open "SELECT * FROM AssetsDataBase WHERE (AssetsNumber = '"& Ucase(FixAssetsFile(0)) &"')", ManagCn, 3, 3
              Else
                  '列管
                  FixTestRs.Open "SELECT * FROM Assets2DataBase WHERE (AssetsNumber = '"& Ucase(FixAssetsFile(0)) &"')", ManagCn, 3, 3
              End If

              Set FixAssetsRs = Server.CreateObject("ADODB.RecordSet")

              If Request.QueryString("Check") <> "M" Then
                 FixAssetsString = "UPDATE AssetsDataBase "
              Else
                  FixAssetsString = "UPDATE Assets2DataBase "
              End If

              FixAssetsString = FixAssetsString & "SET AssetsName = '"& FixAssetsFile(1) & "', " &_
                                                  "AssetsEut = '"& FixAssetsFile(2) & "', " &_
                                                  "AssetsOrdinalNo = '"& FixAssetsFile(3) & "', " &_
                                                  "Amount = '"& FixAssetsFile(4) & "', " &_
                                                  "Unit = '"& FixAssetsFile(5) &"', " &_
                                                  "BuyCast = '"& FixAssetsFile(6) & "', " &_
                                                  "EnduringLimit = '"& FixAssetsFile(7) & "', " &_
                                                  "Insure = "& FixAssetsFile(13) & ", " &_
                                                  "Mortgage = "& FixAssetsFile(14) & ", " &_
                                                  "Iease = "& FixAssetsFile(15) & ", " &_
                                                  "NewState = '"& FixAssetsFile(16) & "', " &_
                                                  "Brand = '"& FixAssetsFile(18) &"', " &_
                                                  "BelongUnit = '"& FixAssetsFile(19) & "', " &_
                                                  "UseUnit = '"& FixAssetsFile(20) & "', " &_
                                                  "Remark = '"& FixAssetsFile(22) &"' "

              If FixAssetsFile(17) = Clear Then
                 FixAssetsString = FixAssetsString & ", KeepDate = NULL "
              Else
                  FixAssetsString = FixAssetsString & ", KeepDate = '"& FixAssetsFile(17) &"' "
              End If

              If FixAssetsFile(8) <> Clear Then
                 If ISNULL(FixTestRs.Fields("CustomerBaseData")) = True OR FixTestRs.Fields("CustomerBaseData") = Clear Then
                    FixAssetsString = FixAssetsString & ", CustomerBaseData = '"& FixAssetsFile(8) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", CustomerBaseData = '"& FixTestRs.Fields("CustomerBaseData") & "`" & FixAssetsFile(8) &"' "
                 End If
              End If

              If FixAssetsFile(9) <> Clear Then
                 Buffer = Split(FixAssetsFile(9), "^")
                 'FixAssetsString = FixAssetsString & ", UseUnit = '"& Buffer(3) &"' "
                 If ISNULL(FixTestRs.Fields("ChangesCondition")) = True OR FixTestRs.Fields("ChangesCondition") = Clear Then
                    FixAssetsString = FixAssetsString & ", ChangesCondition = '"& FixAssetsFile(9) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", ChangesCondition = '"& FixTestRs.Fields("ChangesCondition") & "`" & FixAssetsFile(9) &"' "
                 End If
              End If

              If FixAssetsFile(10) <> Clear Then
                 If ISNULL(FixTestRs.Fields("EstimateCondition")) = True OR FixTestRs.Fields("EstimateCondition") = Clear Then
                    FixAssetsString = FixAssetsString & ", EstimateCondition = '"& FixAssetsFile(10) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", EstimateCondition = '"& FixTestRs.Fields("EstimateCondition") & "`" & FixAssetsFile(10) &"' "
                 End If
              End If

              If FixAssetsFile(11) <> Clear Then
                 If ISNULL(FixTestRs.Fields("MainTainRecord")) = True OR FixTestRs.Fields("MainTainRecord") = Clear Then
                    FixAssetsString = FixAssetsString & ", MainTainRecord = '"& FixAssetsFile(11) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", MainTainRecord = '"& FixTestRs.Fields("MainTainRecord") & "`" & FixAssetsFile(11) &"' "
                 End If
              End If 

              If FixAssetsFile(21) <> Clear Then
                 If ISNULL(FixTestRs.Fields("RectifyRecord")) = True OR FixTestRs.Fields("RectifyRecord") = Clear Then
                    FixAssetsString = FixAssetsString & ", RectifyRecord = '"& FixAssetsFile(21) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", RectifyRecord = '"& FixTestRs.Fields("RectifyRecord") & "`" & FixAssetsFile(21) &"' "
                 End If
              End If 

              If FixAssetsFile(12) <> Clear Then
                 If ISNULL(FixTestRs.Fields("SellCondition")) = true Or FixTestRs.Fields("SellCondition") = Clear Then
                    FixAssetsString = FixAssetsString & ", SellCondition = '"& FixAssetsFile(12) &"' "
                 Else
                     FixAssetsString = FixAssetsString & ", SellCondition = '"& FixTestRs.Fields("SellCondition") & "`" & FixAssetsFile(12) &"' "
                 End If
              End If
              
              FixAssetsString = FixAssetsString & " WHERE (AssetsNumber = '"& UCase(FixAssetsFile(0)) &"')"
              ManagCn.Execute(FixAssetsString)

              'FixAssetsRs.Open FixAssetsString, ManagCn, 3, 3

              Response.Write("<html><head></head><body bgcolor=skyblue onload=alert('資產編號:'+'"& FixTestRs.Fields(0) &"'+'.已修正完畢!!');window.location.replace('ManagMainMenu.asp');></body></html>")                     

              Set FixTestRs = Nothing
              Set FixAssetsRs = Nothing

         '******************************************************************************************************************************************************
  end select
      ManagCn.Close 
  Set ManagCn=Nothing%>