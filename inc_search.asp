<!--#include file="Connections/connDUpaypal.asp" -->
<!--#include file="inc_config.asp" -->
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
Dim rsSearch__MMColParam
rsSearch__MMColParam = "0"
if (Request.QueryString("keyword") <> "") then rsSearch__MMColParam = Request.QueryString("keyword")
%>
<%
set rsSearch = Server.CreateObject("ADODB.Recordset")
rsSearch.ActiveConnection = MM_connDUpaypal_STRING
rsSearch.Source = "SELECT *  FROM qryInventory,  TYPES  WHERE (PRO_TITLE LIKE '%" + Replace(rsSearch__MMColParam, "'", "''") + "%' ) AND PRO_TYPE = TYPE_ID  ORDER BY PRO_NAME ASC"
rsSearch.CursorType = 0
rsSearch.CursorLocation = 2
rsSearch.LockType = 3
rsSearch.Open()
rsSearch_numRows = 0
%>
<%
Dim rsSearch__numRows
rsSearch__numRows = 25
Dim rsSearch__index
rsSearch__index = 0
rsSearch_numRows = rsSearch_numRows + rsSearch__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsSearch_total = rsSearch.RecordCount

' set the number of rows displayed on this page
If (rsSearch_numRows < 0) Then
  rsSearch_numRows = rsSearch_total
Elseif (rsSearch_numRows = 0) Then
  rsSearch_numRows = 1
End If

' set the first and last displayed record
rsSearch_first = 1
rsSearch_last  = rsSearch_first + rsSearch_numRows - 1

' if we have the correct record count, check the other stats
If (rsSearch_total <> -1) Then
  If (rsSearch_first > rsSearch_total) Then rsSearch_first = rsSearch_total
  If (rsSearch_last > rsSearch_total) Then rsSearch_last = rsSearch_total
  If (rsSearch_numRows > rsSearch_total) Then rsSearch_numRows = rsSearch_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsSearch_total = -1) Then

  ' count the total records by iterating through the recordset
  rsSearch_total=0
  While (Not rsSearch.EOF)
    rsSearch_total = rsSearch_total + 1
    rsSearch.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsSearch.CursorType > 0) Then
    rsSearch.MoveFirst
  Else
    rsSearch.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsSearch_numRows < 0 Or rsSearch_numRows > rsSearch_total) Then
    rsSearch_numRows = rsSearch_total
  End If

  ' set the first and last displayed record
  rsSearch_first = 1
  rsSearch_last = rsSearch_first + rsSearch_numRows - 1
  If (rsSearch_first > rsSearch_total) Then rsSearch_first = rsSearch_total
  If (rsSearch_last > rsSearch_total) Then rsSearch_last = rsSearch_total

End If
%>
<%
Dim MM_paramName
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsSearch
MM_rsCount   = rsSearch_total
MM_size      = rsSearch_numRows
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
rsSearch_first = MM_offset + 1
rsSearch_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsSearch_first > MM_rsCount) Then rsSearch_first = MM_rsCount
  If (rsSearch_last > MM_rsCount) Then rsSearch_last = MM_rsCount
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
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)
    If (properly = 1) Then
      Dim TempArray
      TempArray = split(strRet, " ")
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<link href="assets/DUpaypal.css" rel="stylesheet" type="text/css">
<div class = "links">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <% If Not rsSearch.EOF Or Not rsSearch.BOF Then %>
    <tr>
      <td align="left" valign="middle" class="textBold"><a href="default.asp">SHOP</a>
        &raquo; <a href="type.asp?iType=<%=(rsSearch.Fields.Item("PRO_TYPE").Value)%>"><%=UCASE(rsSearch.Fields.Item("TYPE_NAME").Value)%></a> &raquo; LISTING</td>
      <td align="right" valign="middle" class="textBold">
        <%
TM_counter = 0
For i = 1 to rsSearch_total Step MM_size
TM_counter = TM_counter + 1
TM_PageEndCount = i + MM_size - 1
if TM_PageEndCount > rsSearch_total Then TM_PageEndCount = rsSearch_total
if i <> MM_offset + 1 then
Response.Write("<a href=""" & Request.ServerVariables("URL") & "?" & MM_keepMove & "offset=" & i-1 & """>")
Response.Write(TM_counter & "</a>")
else
Response.Write("<b>Page " & TM_counter & "</b>")
End if
if(TM_PageEndCount <> rsSearch_total) then Response.Write(" | ")
next
 %> </td>
    </tr>
    <tr>
      <td colspan="2">
        <%
While ((rsSearch__numRows <> 0) AND (NOT rsSearch.EOF))
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="left" valign="top" bgcolor="#000000"><img src="assets/_spacer.gif" width="1" height="1"></td>
          </tr>
          <tr>
          <form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post"><td width="30" align="left">
		                                    <input type="hidden" name="cmd" value="_cart">
		                                    <input type="hidden" name="business" value="<%= myPaypalID %>">
		  								  <input type="hidden" name="currency_code" value="<%= myPaypalCurrency %>">
		  								  <input type="hidden" name="return" value="<%= myReturnURL %>">
		  								  <input type="hidden" name="cancel_return" value="<%= myCancelURL %>">
		                                    <input type="hidden" name="item_name" value="<%=(rsSearch.Fields.Item("PRO_NAME").Value)%>">
		                                    <input type="hidden" name="item_number" value="<%=(rsSearch.Fields.Item("PRO_ID").Value)%>">
		                                    <input type="hidden" name="amount" value="<%=(rsSearch.Fields.Item("PRO_PRICE").Value)%>">
		                                    <input type="hidden" name="shipping" value="<%=(rsSearch.Fields.Item("PRO_SHIP").Value)%>">
		                                    <input type="hidden" name="add" value="1">
		                                    <input name="addtocart" type="submit" class="form" value="Add to Cart">
                               </td> </form>
            <td align="left" valign="top"> <table width="100%" border="0" cellpadding="2" cellspacing="2">
                <tr align="left" valign="middle">
                  <td align="center" valign="top"><a href="detail.asp?iPro=<%=(rsSearch.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsSearch.Fields.Item("PRO_TYPE").Value)%>"></a></td>
                  <td valign="top">
                  	<table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr valign="top">
                        <td width="65%" align="left" class="textBold"><b></b>
                          <a href="detail.asp?iPro=<%=(rsSearch.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsSearch.Fields.Item("PRO_TYPE").Value)%>"><%=(rsSearch.Fields.Item("PRO_TITLE").Value)%></a>&nbsp;</td>
                        <td width="15%" align="left" class="text"><strong>#:</strong>
                          <%=(rsSearch.Fields.Item("PRO_CARDNUMBER").Value)%></td>
                        <td width="15%" align="left" class="text"><strong>Price:</strong>
                          <%= myPaypalCurrencySign %><%=(rsSearch.Fields.Item("PRO_PRICE").Value)%></td>
                        <td width="5%" align="left" class="text"><a href="detail.asp?iPro=<%=(rsSearch.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsSearch.Fields.Item("PRO_TYPE").Value)%>">more</a></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <%
  rsSearch__index=rsSearch__index+1
  rsSearch__numRows=rsSearch__numRows-1
  rsSearch.MoveNext()
Wend
%> </td>
    </tr>
    <% End If ' end Not rsSearch.EOF Or NOT rsSearch.BOF %>

      <% If rsSearch.EOF And rsSearch.BOF Then %>
	  <tr>
      <td colspan="2"><div class = "links"><font color="#FF0000" class="textBold">Currently
          there are no products listed. Please check back later. Thank You.</font>
        </div></td>
       </tr>
	   <% End If ' end rsSearch.EOF And rsSearch.BOF %>

    <tr align="right" valign="middle">
      <td colspan="2" class="text"><a href="http://www.eyesoft.net">Powered by EyeCart</a></td>
    </tr>

  </table>

</div>

<%
rsSearch.Close()
Set rsSearch = Nothing
%>
