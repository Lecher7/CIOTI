<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/connDUpaypalAdmin.asp" -->
<% Response.Buffer = True %>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("id"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="products.asp"
  MM_redirectLoginFailed="default.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_connDUpaypal_STRING
  MM_rsUser.Source = "SELECT U_ID, U_PASSWORD"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM USERS WHERE U_ID='" & Replace(MM_valUsername,"'","''") &"' AND U_PASSWORD='" & Replace(Request.Form("password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<html>
<head>
<title>DUpaypal</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="assets/DUpaypal.css" rel="stylesheet" type="text/css">
<link href="../assets/DUpaypal.css" rel="stylesheet" type="text/css">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bg>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" background="../assets/bground.gif">
  <tr>
    <td align="center" valign="middle"><table width="100%" height="300" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="2" align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td align="center" valign="middle" bgcolor="#006699"><form name="form1" method="POST" action="<%=MM_LoginAction%>">
              <table border="0" cellspacing="4" cellpadding="4">
                <tr> 
                  <td colspan="2" class="textBold"><font color="#FFFFFF" size="4"><strong>DUpaypal 
                    ADMIN PANEL</strong></font></td>
                </tr>
                <tr> 
                  <td align="right" valign="middle" class="textBold"><font color="#FFFFFF">Admin 
                    ID: </font></td>
                  <td><font color="#FFFFFF"> 
                    <input name="id" type="text" class="form" id="id" value="admin" size="20">
                    </font></td>
                </tr>
                <tr> 
                  <td align="right" valign="middle" class="textBold"><font color="#FFFFFF">Password: 
                    </font></td>
                  <td><font color="#FFFFFF"> 
                    <input name="password" type="password" class="form" id="password" size="20">
                    </font></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td> <input name="Submit" type="submit" class="form"  value="Login"></td>
                </tr>
              </table>
            </form></td>
        </tr>
        <tr> 
          <td height="2" align="left" valign="top" bgcolor="#000000"><img src="../assets/_spacer.gif" width="1" height="1"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>

