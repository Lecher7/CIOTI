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
Dim rsDetail__MMColParam
rsDetail__MMColParam = "0"
if (Request.QueryString("iPro") <> "") then rsDetail__MMColParam = Request.QueryString("iPro")
%>
<%
set rsDetail = Server.CreateObject("ADODB.Recordset")
rsDetail.ActiveConnection = MM_connDUpaypal_STRING
rsDetail.Source = "SELECT *  FROM qryInventory,  TYPES  WHERE PRO_ID = " + Replace(rsDetail__MMColParam, "'", "''") + " AND PRO_TYPE = TYPE_ID"
rsDetail.CursorType = 0
rsDetail.CursorLocation = 2
rsDetail.LockType = 3
rsDetail.Open()
rsDetail_numRows = 0
%>
<%
Dim rsRelated__MMColParam
rsRelated__MMColParam = "0"
if (Request.QueryString("iType") <> "") then rsRelated__MMColParam = Request.QueryString("iType")
%>
<%
Dim rsRelated__var_id
rsRelated__var_id = "0"
if (Request.QueryString("iPro") <> "") then rsRelated__var_id = Request.QueryString("iPro")
%>
<%
set rsRelated = Server.CreateObject("ADODB.Recordset")
rsRelated.ActiveConnection = MM_connDUpaypal_STRING
rsRelated.Source = "SELECT *  FROM qryInventory  WHERE PRO_TYPE = " + Replace(rsRelated__MMColParam, "'", "''") + " AND PRO_ID <> " + Replace(rsRelated__var_id, "'", "''") + "  ORDER BY PRO_TITLE DESC"
rsRelated.CursorType = 0
rsRelated.CursorLocation = 2
rsRelated.LockType = 3
rsRelated.Open()
rsRelated_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 3
Dim Repeat1__index
Repeat1__index = 0
rsRelated_numRows = rsRelated_numRows + Repeat1__numRows
%>
<link href="assets/DUpaypal.css" rel="stylesheet" type="text/css">
<div class = "links">
<table width="100%" border="0" cellspacing="0" cellpadding="3">
  <tr>
      <td align="left" valign="middle" class="textBold"><a href="default.asp">SHOP</a>
        &raquo; <a href="type.asp?iType=<%=(rsDetail.Fields.Item("PRO_TYPE").Value)%>"><%=UCASE(rsDetail.Fields.Item("TYPE_NAME").Value)%></a> &raquo; DETAIL</td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
          <td align="left" valign="top" bgcolor="#000000"><img src="assets/_spacer.gif" width="1" height="1"></td>
        </tr>
        <tr>
            <td align="left" valign="top"> <table width="100%" border="0" cellpadding="2" cellspacing="2">
                <tr align="left" valign="middle">
                  <td align="center" valign="top"><a href="detail.asp?iPro=<%=(rsDetail.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsDetail.Fields.Item("PRO_TYPE").Value)%>"><img src="images/<%=(rsDetail.Fields.Item("PRO_IMAGE").Value)%>" border="0" vspace="0" hspace="0" alt="<%=(rsDetail.Fields.Item("PRO_NAME").Value)%>"></a></td>
                  <td valign="top"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                      <tr valign="top">
                        <td width="70%" align="left" class="textBold"><b>Name:</b>
                          <%=(rsDetail.Fields.Item("PRO_TITLE").Value)%></td>
                        <td width="30%" align="left" class="text"><strong>By:</strong>
                          <%=(rsDetail.Fields.Item("PRO_BRAND").Value)%></td>
                      </tr>
                      <tr valign="top">
                        <td width="50%" align="left" class="text"><strong>Stock/Shipping
                          Info:</strong> <%=(rsDetail.Fields.Item("PRO_INFO").Value)%></td>
                        <td width="50%" align="left" class="text"></td>
                      </tr>
                      <tr valign="top">
                        <td align="left" class="text">

						<table border="0" cellspacing="1" cellpadding="1">
  <tr align="left" valign="middle">

	<form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post"><td>
                                  <input type="hidden" name="cmd" value="_cart">
                                  <input type="hidden" name="business" value="<%= myPaypalID %>">
								  <input type="hidden" name="currency_code" value="<%= myPaypalCurrency %>">
								  <input type="hidden" name="return" value="<%= myReturnURL %>">
								  <input type="hidden" name="cancel_return" value="<%= myCancelURL %>">
                                  <input type="hidden" name="item_name" value="<%=(rsDetail.Fields.Item("PRO_NAME").Value)%>">
                                  <input type="hidden" name="item_number" value="<%=(rsDetail.Fields.Item("PRO_ID").Value)%>">
                                  <input type="hidden" name="amount" value="<%=(rsDetail.Fields.Item("PRO_PRICE").Value)%>">
                                  <input type="hidden" name="shipping" value="<%=(rsDetail.Fields.Item("PRO_SHIP").Value)%>">
                                  <input type="hidden" name="add" value="1">
                                  <input name="addtocart" type="submit" class="form" value="Add to Cart">
                               </td> </form>

  </tr>
</table>
						</td>
                        <td align="left" class="text">
						<table border="0" cellspacing="1" cellpadding="1">
  <tr align="left" valign="middle">


	<form target="paypal" action="https://www.paypal.com/cgi-bin/webscr" method="post"> <td>
                                  <input type="hidden" name="cmd" value="_cart">
                                  <input type="hidden" name="business" value="<%= myPaypalID %>">
                                  <input type="hidden" name="display" value="1">
                                  <input name="viewcart" type="submit" class="form" value="View Cart">
                              </td>  </form>

  </tr>
</table>

						</td>
                      </tr>
                      <tr>
                        <td colspan="2" align="left" valign="middle" class="text"><b>Description:</b>
                          <% =DoSpace(rsDetail.Fields.Item("PRO_DESCRIPTION").Value) %> </td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#000000"><img src="assets/_spacer.gif" width="1" height="1"></td>
        </tr>
        <tr>
            <td height="30" align="left" valign="middle" class="textBold"><font color="#FF0000">MAY
              WE SUGGEST:</font></td>
        </tr>
        <tr>
          <td align="left" valign="top">
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr>
                <%
While ((Repeat1__numRows <> 0) AND (NOT rsRelated.EOF))
%>
                <td align="center" valign="top">
                  <table border="0" cellspacing="2" cellpadding="2">
                    <tr>
                      <td align="center" valign="middle"><a href="detail.asp?iPro=<%=(rsRelated.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsRelated.Fields.Item("PRO_TYPE").Value)%>"><img src="images/<%=(rsRelated.Fields.Item("PRO_IMAGE").Value)%>" border="0" alt="Buy  <%=(rsRelated.Fields.Item("PRO_NAME").Value)%>"></a></td>
                    </tr>
                    <tr>
                        <td align="center" valign="middle" class="textBold"><a href="detail.asp?iPro=<%=(rsRelated.Fields.Item("PRO_ID").Value)%>&iType=<%=(rsRelated.Fields.Item("PRO_TYPE").Value)%>"><%=(rsRelated.Fields.Item("PRO_NAME").Value)%></a> <br>
                          <font size="1">(</font><%=(rsRelated.Fields.Item("PRO_PRICE").Value)%><font size="1"><b>)</b></font></td>
                    </tr>
                  </table>
                </td>
                <%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsRelated.MoveNext()
Wend
%>
              </tr>
            </table>
          </td>
        </tr>
		 <tr>
          <td align="left" valign="top" bgcolor="#000000"><img src="assets/_spacer.gif" width="1" height="1"></td>
        </tr>
		 <tr>
      <td align="right" valign="middle" class="text"><a href="http://www.eyesoft.net">Powered by EyeCart</a></td>
    </tr>
      </table>
    </td>
  </tr>


</table>
</div>
<%
rsDetail.Close()
%>
<%
rsRelated.Close()
%>
