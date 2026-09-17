<!--#include file="Connections/connDUpaypal.asp" -->
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
set rsCount = Server.CreateObject("ADODB.Recordset")
rsCount.ActiveConnection = MM_connDUpaypal_STRING
rsCount.Source = "SELECT (SELECT COUNT(*) FROM qryInventory) AS PRO_COUNTs, COUNT(*)   AS TYPE_COUNT  FROM TYPES"
rsCount.CursorType = 0
rsCount.CursorLocation = 2
rsCount.LockType = 3
rsCount.Open()
rsCount_numRows = 0
%>

 <link href="assets/DUpaypal.css" rel="stylesheet" type="text/css">
 <table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr align="left" valign="middle">
     <td align="left" valign="middle" class="text">There are <b><%=(rsCount.Fields.Item("PRO_COUNTs").Value)%></b> products in <b><%=(rsCount.Fields.Item("TYPE_COUNT").Value)%></b> categories</td>
    <form action="search.asp" method="get" name="search" id="search">
      <td align="right">
        <input name="keyword" type="text" class="form" id="keyword" size="25">
        <input name="Search" type="submit" class="form" id="Search" value="Search">
      </td></form>
  </tr>
</table>
<%
rsCount.Close()
%>
