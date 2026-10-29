<%@ Language=VBScript %>
<!-- #INCLUDE FILE="../Tools/Function.asp" -->

<%EmpID = Trim(request.Form("EmpID"))
  Pwd = Trim(request.Form("Pwd"))
  
  If EmpID = "" OR Pwd = "" Then
     response.Redirect "Default.asp"
     response.End
  End If

  '-----------------------------------------------------------------------------------------------------------------------------------------------------------
  Set ManagCn = Server.CreateObject("ADODB.Connection")
      ManagCn.Open GV_strConnAdmin

  '-----------------------------------------------------------------------------------------------------------------------------------------------------------  
  Str1 = "SELECT EmpID, Pwd, CName FROM Employee WHERE (EmpID = '"& EmpID &"')"
  Set rs1 = GF_GetRS(Str1,ManagCn)
  If rs1.EOF Then%>
     <script language=javascript>
             alert('µL¦¹±b¸¹¡I');
             history.back(-1);
     </script>
   <%response.End
  Else
      RS1_EmpID = Trim(rs1("EmpID"))
      RS1_Pwd = Trim(rs1("Pwd"))
      RS1_CName = Trim(rs1("CName"))
      If Pwd <> RS1_Pwd Then%>
         <script language=javascript>
                 alert('±K½X¿ù»~¡I');
                 history.back(-1);
         </script>
       <%response.End
      Else
          Session("EmpID") = RS1_EmpID
          Session("Cname") = RS1_CName
          response.Redirect "ManagMainMenu.asp"
      End If
  End If
      rs1.Close
  Set rs1 = Nothing
      adocn.Close
  Set adocn = Nothing%>