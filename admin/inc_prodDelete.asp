<!--#include file="../Connections/connDUpaypalAdmin.asp" -->
<!--#include file="../ScriptLibrary/incPUAddOn.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice                                                               
'**  Copyright 2003 DUware All Rights Reserved.                                
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to DUware and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.duware.com/support
'****************************************************************************************
%>

<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="default.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>


<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUpaypal_STRING
  MM_editTable = "PRODUCTS"
  MM_editColumn = "PRO_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "products.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<% 
' *** Delete File Before Delete Record 1.6.0
If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then
  Dim DF_filesStr, DF_path, DF_suffix
  DF_filesStr = "PRO_IMAGE"
  DF_path = "../images"
  DF_suffix = "_small"
  DeleteFileBeforeRecord DF_filesStr,DF_path,MM_editConnection,MM_editTable,MM_editColumn,MM_recordId,DF_suffix
end if
%>

<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<%
Dim rsProducts__MMColParam
rsProducts__MMColParam = "1"
if (Request.QueryString("iPro") <> "") then rsProducts__MMColParam = Request.QueryString("iPro")
%>
<%
set rsProducts = Server.CreateObject("ADODB.Recordset")
rsProducts.ActiveConnection = MM_connDUpaypal_STRING
rsProducts.Source = "SELECT * FROM PRODUCTS WHERE PRO_ID = " + Replace(rsProducts__MMColParam, "'", "''") + ""
rsProducts.CursorType = 0
rsProducts.CursorLocation = 2
rsProducts.LockType = 3
rsProducts.Open()
rsProducts_numRows = 0
%>
<%
set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.ActiveConnection = MM_connDUpaypal_STRING
rsTypes.Source = "SELECT * FROM TYPES ORDER BY TYPE_NAME ASC"
rsTypes.CursorType = 0
rsTypes.CursorLocation = 2
rsTypes.LockType = 3
rsTypes.Open()
rsTypes_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 20
Dim Repeat1__index
Repeat1__index = 0
rsProducts_numRows = rsProducts_numRows + Repeat1__numRows
%>
<script language="JavaScript">
<!--
function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>

<link href="../assets/DUpaypal.css" rel="stylesheet" type="text/css">

  
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}


//-->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
            <td align="left" valign="top"> 
        <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" onSubmit="checkFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',false,'','','','','','','');return document.MM_returnValue">
            <td align="left" valign="top"> <table align="center" cellpadding="2" cellspacing="2" class="textBold">
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">CATEGORY:</td>
                  <td valign="middle"> <select name="PRO_TYPE" class="form">
                      <%
While (NOT rsTypes.EOF)
%>
                      <option value="<%=(rsTypes.Fields.Item("TYPE_ID").Value)%>" <%if (CStr(rsTypes.Fields.Item("TYPE_ID").Value) = CStr(rsProducts.Fields.Item("PRO_TYPE").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(rsTypes.Fields.Item("TYPE_NAME").Value)%></option>
                      <%
  rsTypes.MoveNext()
Wend
If (rsTypes.CursorType > 0) Then
  rsTypes.MoveFirst
Else
  rsTypes.Requery
End If
%>
                    </select> </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">NAME:</td>
                  <td valign="middle"> <input name="PRO_NAME" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_NAME").Value)%>" size="60"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">DATED:</td>
                  <td valign="middle"> <input name="PRO_DATED" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_DATED").Value)%>" size="20"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">PRICE:</td>
                  <td valign="middle"> <input name="PRO_PRICE" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_PRICE").Value)%>" size="12"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">SHIPPING/HANDLING:</td>
                  <td valign="middle"> <input name="PRO_SHIP" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_SHIP").Value)%>" size="20"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">SHIPPING/STOCK INFO:</td>
                  <td valign="middle"> <input name="PRO_INFO" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_INFO").Value)%>" size="40"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">BRAND:</td>
                  <td valign="middle"> <input name="PRO_BRAND" type="text" class="form" value="<%=(rsProducts.Fields.Item("PRO_BRAND").Value)%>" size="45"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td colspan="2" align="left" valign="middle" nowrap>Leave blank 
                    if not replacing the image:</td>
                </tr>
                <tr> 
                  <td nowrap align="right" valign="top">DESCRIPTION:</td>
                  <td valign="middle"> <textarea name="PRO_DESCRIPTION" cols="60" rows="15" class="form"><%=(rsProducts.Fields.Item("PRO_DESCRIPTION").Value)%></textarea> 
                  </td>
                </tr>
                <tr valign="baseline" align="center"> 
                  <td nowrap colspan="2"> <input name="delete" type="submit" class="form" id="delete" onClick="MM_validateForm('PRO_NAME','','R','PRO_DATED','','R','PRO_PRICE','','RisNum','PRO_SHIP','','RisNum','PRO_INFO','','R','PRO_DESCRIPTION','','R');return document.MM_returnValue" value="Delete This Product"> 
                  </td>
                </tr>
              </table></td>
            <input type="hidden" name="MM_delete" value="form1">
            <input type="hidden" name="MM_recordId" value="<%= rsProducts.Fields.Item("PRO_ID").Value %>">
          </form>
        </tr>
        
      </table>
      </td>
    </tr>
  </table>
<%
rsProducts.Close()
%>
<%
rsTypes.Close()
%>

