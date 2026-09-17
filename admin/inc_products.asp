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
Dim tfm_orderby, tfm_order
tfm_orderby = "PRO_NAME"
tfm_order = "ASC"
If(CStr(Request.QueryString("tfm_orderby")) <> "") Then
	tfm_orderby = Cstr(Request.QueryString("tfm_orderby"))
End If
If(Cstr(Request.QueryString("tfm_order")) <> "") Then
	tfm_order = Cstr(Request.QueryString("tfm_order"))
End If

Dim sql_orderby
sql_orderby = " " & tfm_orderby & " " & tfm_order
%>
<%
Dim rsProducts__sql_orderby
rsProducts__sql_orderby = "PRO_NAME"
if (sql_orderby <> "") then rsProducts__sql_orderby = sql_orderby
%>
<%
set rsProducts = Server.CreateObject("ADODB.Recordset")
rsProducts.ActiveConnection = MM_connDUpaypal_STRING
rsProducts.Source = "SELECT *  FROM PRODUCTS, TYPES WHERE PRO_TYPE = TYPE_ID  ORDER BY " + Replace(rsProducts__sql_orderby, "'", "''") + ""
rsProducts.CursorType = 0
rsProducts.CursorLocation = 2
rsProducts.LockType = 3
rsProducts.Open()
rsProducts_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 20
Dim Repeat1__index
Repeat1__index = 0
rsProducts_numRows = rsProducts_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsProducts_total = rsProducts.RecordCount

' set the number of rows displayed on this page
If (rsProducts_numRows < 0) Then
  rsProducts_numRows = rsProducts_total
Elseif (rsProducts_numRows = 0) Then
  rsProducts_numRows = 1
End If

' set the first and last displayed record
rsProducts_first = 1
rsProducts_last  = rsProducts_first + rsProducts_numRows - 1

' if we have the correct record count, check the other stats
If (rsProducts_total <> -1) Then
  If (rsProducts_first > rsProducts_total) Then rsProducts_first = rsProducts_total
  If (rsProducts_last > rsProducts_total) Then rsProducts_last = rsProducts_total
  If (rsProducts_numRows > rsProducts_total) Then rsProducts_numRows = rsProducts_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsProducts_total = -1) Then

  ' count the total records by iterating through the recordset
  rsProducts_total=0
  While (Not rsProducts.EOF)
    rsProducts_total = rsProducts_total + 1
    rsProducts.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsProducts.CursorType > 0) Then
    rsProducts.MoveFirst
  Else
    rsProducts.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsProducts_numRows < 0 Or rsProducts_numRows > rsProducts_total) Then
    rsProducts_numRows = rsProducts_total
  End If

  ' set the first and last displayed record
  rsProducts_first = 1
  rsProducts_last = rsProducts_first + rsProducts_numRows - 1
  If (rsProducts_first > rsProducts_total) Then rsProducts_first = rsProducts_total
  If (rsProducts_last > rsProducts_total) Then rsProducts_last = rsProducts_total

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsProducts
MM_rsCount   = rsProducts_total
MM_size      = rsProducts_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsProducts_first = MM_offset + 1
rsProducts_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsProducts_first > MM_rsCount) Then rsProducts_first = MM_rsCount
  If (rsProducts_last > MM_rsCount) Then rsProducts_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<%
'sort column headers for rsProducts
Dim tfm_saveParams, tfm_keepParams, tfm_orderbyURL
tfm_saveParams = ""
tfm_keepParams = ""
If tfm_order = "ASC" Then
	tfm_order = "DESC"
Else
	tfm_order = "ASC"
End If
		
If tfm_saveParams <> "" Then
	tfm_params = Split(tfm_saveParams,",")
	For i = 0 to UBound(tfm_params)
		If Cstr(Request(tfm_params(i))) <> "" Then
			tfm_keepParams = tfm_keepParams & LCase(tfm_params(i)) & "=" & Server.URLEncode(Request(tfm_params(i))) & "&"
		End If
	Next
End If
tfm_orderbyURL = Request.ServerVariables("URL") & "?" & tfm_keepParams & "tfm_order=" & tfm_order & "&tfm_orderby="
%>
<link href="../assets/DUpaypal.css" rel="stylesheet" type="text/css">
<div class = "links"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    
    
    <tr> 
      <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="left" valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td align="right" valign="middle" class="textBold"> 
                    <%
TM_counter = 0
For i = 1 to rsProducts_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > rsProducts_total Then TM_PageEndCount = rsProducts_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>Page " & TM_counter & "</b>")
End if
if(TM_PageEndCount <> rsProducts_total) then Response.Write(" | ")
next
 %> 
                  </td>
                </tr>
                <tr> 
                  <td align="left" valign="top"> <table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#333333">
                      <tr align="center" valign="middle" bgcolor="#CCCCCC" class="textBold"> 
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_NAME">NAME</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>TYPE_NAME">TYPE</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_DATED">DATE</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_PRICE">PRICE</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_SHIP">SHIP</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_INFO">INFO</a></td>
                        <td height="18"><a href="<%=tfm_orderbyURL%>PRO_BRAND">BRAND</a></td>
                        <td height="18">EDIT</td>
						<td height="18">DELETE</td>
                      </tr>
                      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsProducts.EOF)) 
%>
                      <tr align="center" valign="middle" class="text"> 
                        <td align="left" bgcolor="#FFFFFF" class="textBold"><a href="../detail.asp?iPro=<%=(rsProducts.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsProducts.Fields.Item("PRO_TYPE").Value)%>" target="_blank"><%=(rsProducts.Fields.Item("PRO_NAME").Value)%></a></td>
                        <td align="left" bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("TYPE_NAME").Value)%></td>
                        <td align="center" bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("PRO_DATED").Value)%></td>
                        <td bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("PRO_PRICE").Value)%></td>
                        <td align="center" bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("PRO_SHIP").Value)%></td>
                        <td align="left" bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("PRO_INFO").Value)%></td>
                        <td align="left" bgcolor="#FFFFFF"><%=(rsProducts.Fields.Item("PRO_BRAND").Value)%></td>
                        <td bgcolor="#FFFFFF"><a href="prodEdit.asp?iPro=<%=(rsProducts.Fields.Item("PRO_ID").Value)%>"><img src="../assets/folderIcon.gif" width="22" height="18" align="absmiddle" border="0"></a></td>
						<td bgcolor="#FFFFFF"><a href="prodDelete.asp?iPro=<%=(rsProducts.Fields.Item("PRO_ID").Value)%>"><img src="../assets/folderIcon.gif" width="22" height="18" align="absmiddle" border="0"></a></td>
                      </tr>
                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsProducts.MoveNext()
Wend
%>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
<%
rsProducts.Close()
%>
