<!--#include file="../Connections/connDUpaypalAdmin.asp" -->
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

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_connDUpaypal_STRING
  MM_editTable = "TYPES"
  MM_editRedirectUrl = "type.asp"
  MM_fieldsStr  = "TYPE_NAME|value|TYPE_DESCRIPTION|value"
  MM_columnsStr = "TYPE_NAME|',none,''|TYPE_DESCRIPTION|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connDUpaypal_STRING
  MM_editTable = "TYPES"
  MM_editColumn = "TYPE_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "type.asp"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
  
  	' delete products of this type
	set cmdDelete = Server.CreateObject("ADODB.Command")
	cmdDelete.ActiveConnection = MM_connDUpaypal_STRING
	cmdDelete.CommandText = "DELETE * FROM PRODUCTS WHERE PRO_TYPE = " & MM_recordId
	cmdDelete.CommandType = 1
	cmdDelete.CommandTimeout = 0
	cmdDelete.Prepared = true
	cmdDelete.Execute()


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
set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.ActiveConnection = MM_connDUpaypal_STRING
rsTypes.Source = "SELECT *, (SELECT COUNT(*) FROM PRODUCTS WHERE PRO_TYPE = TYPE_ID) AS PRO_COUNT  FROM TYPES  ORDER BY TYPE_NAME ASC"
rsTypes.CursorType = 0
rsTypes.CursorLocation = 2
rsTypes.LockType = 3
rsTypes.Open()
rsTypes_numRows = 0
%>

<%
Dim HLooper1__numRows
HLooper1__numRows = -2
Dim HLooper1__index
HLooper1__index = 0
rsTypes_numRows = rsTypes_numRows + HLooper1__numRows
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
          <td align="left" valign="middle"> <form method="post" action="<%=MM_editAction%>" name="ADD">
              <table align="center" class="textBold">
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">CATEGORY:</td>
                  <td> <input name="TYPE_NAME" type="text" class="form" value="" size="50"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">DESCRIPTION:</td>
                  <td> <input name="TYPE_DESCRIPTION" type="text" class="form" size="60"> </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right">&nbsp;</td>
                  <td> <input name="Add" type="submit" class="form" id="Add" onClick="MM_validateForm('TYPE_NAME','','R','TYPE_DESCRIPTION','','R');return document.MM_returnValue" value="Add"> 
                  </td>
                </tr>
              </table>
              <input type="hidden" name="MM_insert" value="true">
            </form></td>
        </tr>
        <tr> 
          <td align="left" valign="middle"> <hr> </td>
        </tr>
        <tr> 
          <td align="left" valign="top" class="textBold"><font color="#FF0000">If 
            you would like to delete a product category, be sure to check if you 
            really want to do this. Deleting a product category will also delete 
            all of its products.</font></td>
        </tr>
        <tr> 
          <td align="center" valign="top"> <table>
              <%
startrw = 0
endrw = HLooper1__index
numberColumns = 2
numrows = -1
while((numrows <> 0) AND (Not rsTypes.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
              <tr align="center" valign="top"> 
                <%
While ((startrw <= endrw) AND (Not rsTypes.EOF))
%>
                <form ACTION="<%=MM_editAction%>" METHOD="POST" name="DELETE">
                  <td> <table width="100%" border="0" cellpadding="2" cellspacing="2">
                      <tr align="left" valign="middle"> 
                        <td width="22"> <input name="Submit" type="submit" class="form" value="DEL"> 
                        </td>
                        <td class="textBold"><a href="type.asp?iType=<%=(rsTypes.Fields.Item("TYPE_ID").Value)%>"><%=(rsTypes.Fields.Item("TYPE_NAME").Value)%></a> (<%=(rsTypes.Fields.Item("PRO_COUNT").Value)%>) </td>
                      </tr>
                      <tr align="left" valign="middle"> 
                        <td>&nbsp;</td>
                        <td class="textGray"><i><%=(rsTypes.Fields.Item("TYPE_DESCRIPTION").Value)%></i></td>
                        <input type="hidden" name="MM_delete" value="DELETE">
                        <input type="hidden" name="MM_recordId" value="<%= rsTypes.Fields.Item("TYPE_ID").Value %>">
                      </tr>
                    </table></td>
                </form>
                <%
	startrw = startrw + 1
	rsTypes.MoveNext()
	Wend
	%>
              </tr>
              <%
 numrows=numrows-1
 Wend
 %>
            </table></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
<%
rsTypes.Close()
%>

