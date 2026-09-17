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
set rsTypes = Server.CreateObject("ADODB.Recordset")
rsTypes.ActiveConnection = MM_connDUpaypal_STRING
rsTypes.Source = "SELECT *, (SELECT COUNT(*) FROM qryInventory WHERE PRO_TYPE = TYPE_ID) AS PRO_COUNT  FROM TYPES  ORDER BY TYPE_NAME ASC"
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
<link href="assets/DUpaypal.css" rel="stylesheet" type="text/css">
<div class = "links">
  <table border="0" cellspacing="2" cellpadding="2" width="100%">

    <tr>
      <td align="center" valign="middle"> <table width="98%" cellpadding="2" cellspacing="2">
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
            <td> <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr align="left" valign="middle">
                  <td width="22"><img src="assets/folderIcon.gif" width="22" height="18" align="absmiddle"></td>
                  <td class="textBold"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><a href="type.asp?iType=<%=(rsTypes.Fields.Item("TYPE_ID").Value)%>"><%=(rsTypes.Fields.Item("TYPE_NAME").Value)%></a></font> (<%=(rsTypes.Fields.Item("PRO_COUNT").Value)%>)</td>
                </tr>
                <tr align="left" valign="middle">
                  <td>&nbsp;</td>
                  <td class="textGray"><i><%=(rsTypes.Fields.Item("TYPE_DESCRIPTION").Value)%></i></td>
                </tr>
              </table></td>
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
    <tr>
      <td align="right" valign="middle" class="text"><a href="http://www.eyesoft.net">Powered by EyeCart</a></td>
    </tr>
  </table>
</div>
<%
rsTypes.Close()
%>
