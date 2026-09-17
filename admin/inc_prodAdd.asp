<!--#include file="../Connections/connDUpaypalAdmin.asp" -->
<!--#include file="../ScriptLibrary/incPureUpload.asp" -->
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
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright (c) 2001-2002 George Petrov, www.UDzone.com
' Process the upload
' Version: 2.0.9
'------------------------------------------------------------------------------
'*** File Upload to: """../images""", Extensions: "GIF,JPG,JPEG,BMP,PNG", Form: form1, Redirect: "", "file", "", "uniq", "false", "", "" , "", "", "", "", "600", "", "", ""

Dim GP_redirectPage, RequestBin, UploadQueryString, GP_uploadAction, UploadRequest
PureUploadSetup

If (CStr(Request.QueryString("GP_upload")) <> "") Then
  on error resume next
  Dim reqPureUploadVersion, foundPureUploadVersion
  reqPureUploadVersion = 2.09
  foundPureUploadVersion = getPureUploadVersion()
  if err or reqPureUploadVersion > foundPureUploadVersion then
    Response.Write "<b>You don't have latest version of ScriptLibrary/incPureUpload.asp uploaded on the server.</b><br>"
    Response.Write "This library is required for the current page. It is fully backwards compatible so old pages will work as well.<br>"
    Response.End    
  end if
  on error goto 0
  GP_redirectPage = ""
  Server.ScriptTimeout = 600
  
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  BuildUploadRequest RequestBin, """../images""", "file", "", "uniq"
  
  If (GP_redirectPage <> "" and not (CStr(UploadFormRequest("MM_insert")) <> "" or CStr(UploadFormRequest("MM_update")) <> "")) Then
    If (InStr(1, GP_redirectPage, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      GP_redirectPage = GP_redirectPage & "?" & UploadQueryString
    End If
    Response.Redirect(GP_redirectPage)  
  end if  
else
  if UploadQueryString <> "" then
    UploadQueryString = UploadQueryString & "&GP_upload=true"
  else  
    UploadQueryString = "GP_upload=true"
  end if  
end if
' End Pure Upload
'------------------------------------------------------------------------------
%>
<%
' *** Edit Operations: (Modified for File Upload) declare variables

MM_editAction = CStr(Request.ServerVariables("URL")) 'MM_editAction = CStr(Request("URL"))
If (UploadQueryString <> "") Then
  MM_editAction = MM_editAction & "?" & UploadQueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: (Modified for File Upload) set variables

If (CStr(UploadFormRequest("MM_insert")) <> "") Then

  MM_editConnection = MM_connDUpaypal_STRING
  MM_editTable = "PRODUCTS"
  MM_editRedirectUrl = "products.asp"
  MM_fieldsStr  = "PRO_TYPE|value|PRO_NAME|value|PRO_DATED|value|PRO_PRICE|value|PRO_SHIP|value|PRO_INFO|value|PRO_BRAND|value|PRO_IMAGE|value|PRO_DESCRIPTION|value"
  MM_columnsStr = "PRO_TYPE|none,none,NULL|PRO_NAME|',none,''|PRO_DATED|',none,''|PRO_PRICE|',none,''|PRO_SHIP|',none,''|PRO_INFO|',none,''|PRO_BRAND|',none,''|PRO_IMAGE|',none,''|PRO_DESCRIPTION|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(UploadFormRequest(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And UploadQueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & UploadQueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & UploadQueryString
    End If
  End If

End If
%>

<%
' *** Insert Record: (Modified for File Upload) construct a sql insert statement and execute it

If (CStr(UploadFormRequest("MM_insert")) <> "") Then

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
set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.ActiveConnection = MM_connDUpaypal_STRING
rsTypes.Source = "SELECT * FROM TYPES ORDER BY TYPE_NAME ASC"
rsTypes.CursorType = 0
rsTypes.CursorLocation = 2
rsTypes.LockType = 3
rsTypes.Open()
rsTypes_numRows = 0
%>
<link rel="stylesheet" href="../assets/default.css" type="text/css">
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

function checkFileUpload(form,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.09
  document.MM_returnValue = true;
  for (var i = 0; i<form.elements.length; i++) {
    field = form.elements[i];
    if (field.type.toUpperCase() != 'FILE') continue;
    checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
} }

function checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.09
  document.MM_returnValue = true;
  if (extensions != '') var re = new RegExp("\.(" + extensions.replace(/,/gi,"|").replace(/\s/gi,"") + ")$","i");
    if (field.value == '') {
      if (requireUpload) {alert('File is required!');document.MM_returnValue = false;field.focus();return;}
    } else {
      if(extensions != '' && !re.test(field.value)) {
        alert('This file type is not allowed for uploading.\nOnly the following file extensions are allowed: ' + extensions + '.\nPlease select another file and try again.');
        document.MM_returnValue = false;field.focus();return;
      }
    document.PU_uploadForm = field.form;
    re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
    if(re.test(field.value) && (sizeLimit != '' || minWidth != '' || minHeight != '' || maxWidth != '' || maxHeight != '' || saveWidth != '' || saveHeight != '')) {
      checkImageDimensions(field,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
    } }
}

function showImageDimensions(fieldImg) { //v2.09
  var isNS6 = (!document.all && document.getElementById ? true : false);
  var img = (fieldImg && !isNS6 ? fieldImg : this);
  if (img.width > 0 && img.height > 0) {
  if ((img.minWidth != '' && img.minWidth > img.width) || (img.minHeight != '' && img.minHeight > img.height)) {
    alert('Uploaded Image is too small!\nShould be at least ' + img.minWidth + ' x ' + img.minHeight); return;}
  if ((img.maxWidth != '' && img.width > img.maxWidth) || (img.maxHeight != '' && img.height > img.maxHeight)) {
    alert('Uploaded Image is too big!\nShould be max ' + img.maxWidth + ' x ' + img.maxHeight); return;}
  if (img.sizeLimit != '' && img.fileSize > img.sizeLimit) {
    alert('Uploaded Image File Size is too big!\nShould be max ' + (img.sizeLimit/1024) + ' KBytes'); return;}
  if (img.saveWidth != '') document.PU_uploadForm[img.saveWidth].value = img.width;
  if (img.saveHeight != '') document.PU_uploadForm[img.saveHeight].value = img.height;
  document.MM_returnValue = true;
} }

function checkImageDimensions(field,sizeL,minW,minH,maxW,maxH,saveW,saveH) { //v2.09
  if (!document.layers) {
    var isNS6 = (!document.all && document.getElementById ? true : false);
    document.MM_returnValue = false; var imgURL = 'file:///' + field.value.replace(/\\/gi,'/').replace(/:/gi,'|').replace(/"/gi,'').replace(/^\//,'');
    if (!field.gp_img || (field.gp_img && field.gp_img.src != imgURL) || isNS6) {field.gp_img = new Image();
		   with (field) {gp_img.sizeLimit = sizeL*1024; gp_img.minWidth = minW; gp_img.minHeight = minH; gp_img.maxWidth = maxW; gp_img.maxHeight = maxH;
  	   gp_img.saveWidth = saveW; gp_img.saveHeight = saveH; gp_img.onload = showImageDimensions; gp_img.src = imgURL; }
	 } else showImageDimensions(field.gp_img);}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script>
<link href="../assets/DUpaypal.css" rel="stylesheet" type="text/css">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
    <tr> 
     
      <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
         
          <tr> 
          <form ACTION="<%=MM_editAction%>" METHOD="POST" enctype="multipart/form-data" name="form1" onSubmit="checkFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',false,'','','','','','','');return document.MM_returnValue">
            <td align="left" valign="top"> <table align="center" cellpadding="2" cellspacing="2" class="textBold">
                <tr valign="top" align="left"> 
                  <td nowrap colspan="2">&nbsp;</td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">CATEGORY:</b></td>
                  <td valign="middle"> <select name="PRO_TYPE" class="form">
                      <%
While (NOT rsTypes.EOF)
%>
                      <option value="<%=(rsTypes.Fields.Item("TYPE_ID").Value)%>" ><%=(rsTypes.Fields.Item("TYPE_NAME").Value)%></option>
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
                  <td nowrap align="right" valign="middle">NAME:</b></td>
                  <td valign="middle"> <input name="PRO_NAME" type="text" class="form" size="60"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">DATED:</b></td>
                  <td valign="middle"> <input name="PRO_DATED" type="text" class="form" value="<%= date() %>" size="20"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">PRICE:</b></td>
                  <td valign="middle"> <input name="PRO_PRICE" type="text" class="form" size="12">
                    (Numbers only) </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">SHIPPING/HANDLING:</b></td>
                  <td valign="middle"> <input name="PRO_SHIP" type="text" class="form" size="20">
                    (Numbers only)</td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">SHIPPING/STOCK INFO:</b></td>
                  <td valign="middle"> <input name="PRO_INFO" type="text" class="form" value="Usually ship within 24 hours" size="40"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">BRAND:</td>
                  <td valign="middle"> <input name="PRO_BRAND" type="text" class="form" size="45"> 
                  </td>
                </tr>
                <tr valign="baseline"> 
                  <td nowrap align="right" valign="middle">PRODUCT IMAGE:</b></td>
                  <td valign="middle"> <input name="PRO_IMAGE" type="file" class="form" id="PRO_IMAGE" onChange="checkOneFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',false,'','','','','','','')" size="45"> 
                  </td>
                </tr>
                <tr> 
                  <td nowrap align="right" valign="top">DESCRIPTION:</td>
                  <td valign="middle"> <textarea name="PRO_DESCRIPTION" cols="60" rows="15" class="form"></textarea> 
                  </td>
                </tr>
                <tr valign="baseline" align="center"> 
                  <td nowrap colspan="2"> <input type="submit" class="form" onClick="MM_validateForm('PRO_NAME','','R','PRO_DATED','','R','PRO_PRICE','','RisNum','PRO_SHIP','','RisNum','PRO_INFO','','R','PRO_BRAND','','R','PRO_DESCRIPTION','','R');return document.MM_returnValue" value="Add This Product"> 
                  </td>
                </tr>
              </table></td>
            <input type="hidden" name="MM_insert" value="form1">
          </form>
          </tr>
        </table></td>
    </tr>
  </table>
<%
rsTypes.Close()
%>

