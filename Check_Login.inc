<%If Session("EmpID") = "" OR Session("Cname") = "" Then
     host_addr = request.ServerVariables("HTTP_HOST")%>
     <script language=javascript>
             alert('停滯過久！請重新登入！');
             window.location.href = 'http://<%=host_addr%>/Admin/';
     </script>
   <%response.End
  End If%>