<%

'#########################################
'## SET SIGNATURE COUNTER DATABASE PATH ##
'#########################################

  SQL="SELECT Count([ID]) AS OFFENDERS FROM tblCIOTI"
  DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../CIOTI/database/petition.mdb")

   dim conntemp, rstemp
   set conntemp=server.createobject("adodb.connection")
   conntemp.open DSN
   set rstemp=conntemp.execute(SQL)

   if not(rstemp.eof) THEN

   OFFENDERS=rstemp(0)

   end if


%>
<html>

<head>
<title>www.CIOTI.com</title>
</head>

<body bgcolor="#AFBFCB">
<div align="center"><center>

<font color=white><b>C</b></font>opyright <font color=white><b>I</b></font>nfringement <font color=white><b>O</b></font>n <font color=white><b>T</b></font>he <font color=white><b>I</b></font>nternet
<table width="750"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top" background="img/lt_shad.gif"><img src="img/lt_shad.gif" width="9" height="3"></td>
    <td width="100%" bgcolor="#FFFFFF">
      <table width="100%"  border="0" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF" class="shell">
        <tr>
          <td valign="top">
            <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#C1CDD6">
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>
                        <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="hdcella">
                          <tr>
                            <td bgcolor="#87A0B2"><font color="white"><div class="head" style="padding: 6px;margin: 0px auto";>Welcome</div></font></td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td valign="top" background="img/bg4.gif">
                        <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="contenta">
                          <tr>
                            <td valign="top" bgcolor="#AFBFCB"><div class="subtx" style="padding-left: 8px;padding-right: 8px;padding-top: 1px;margin: 0px auto";>
                              <p align="justify" style="margin-bottom: 0">CIOTI is an online repository of people who have committed Copyright Infringement on the Internet.  This list was culled from a vast number of resources by a team of professionals who know the in's and out's of Internet copyright infringers, and was collected from personal contact with the infringers via websites, Usenet, auction sites and other contacts.<br>
                              <br>
                              This database represents the largest collection of bootleggers and pirates on the Internet, and the people herein account for well over <b>$5,000 worth of sales PER DAY</b> of illegal contraband.  This list grows larger every day, as does the amount of money these individuals take from the holders of these copyrights.<br>
                              <br>
                      		  If you represent musicians and their interests, the question you need to ask yourself is, "How much longer can you afford NOT to use our services?" <br>
                      		  <br>
                      		  CIOTI is here to <a href="testimonials.html">help</a>.  <a href="mailto:info@cioti.com">Contact us</a> if you would like to find out how.<br>
                      		  <br>
                              <center>
                      	   	    <font color="#FFFFFF"><strong><br>
							    Total offenders in our database:</strong></font><font color="#FFFF00"> <strong><%=OFFENDERS%></strong></font></small></small></font>
					          </center></p></div>
					        </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
<!--
              <tr>
                    <td bgcolor="#C1CDD6"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><table width="390"  border="0" cellpadding="0" cellspacing="0" class="hdcella">
                            <tr>
                              <td bgcolor="#87A0B2"><font color="white"><div class="head" style="padding: 6px;margin: 0px auto";>Current Activities</div></font></td>
                            </tr>
                          </table></td>
                          <td><img src="img/invis.gif" width="5" height="5"></td>
                          <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="hdcellb">
                              <tr>
                                <td bgcolor="#87A0B2"><font color="white"><div class="head" style="padding: 6px;margin: 0px auto";>Spotlight Feature </div></font></td>
                              </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td rowspan="3" valign="top" background="img/bg4.gif"><table width="390"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center"><img src="img/fimg/61505.jpg" width="300" height="200"></div></td>
                              </tr>
                              <tr>
                                <td><div class="subtx" style="padding-left: 8px;padding-right: 8px;padding-top: 8px;margin: 0px auto";><span class="subHead">New York -- June 15, 2005</span> <br>
                                        <br>
                    Green Street Holding (owner), RCDolner (contruction manager) and H. Thomas O'Hara (architect) celebrate the ground breaking....</div></td>
                              </tr>
                              <tr>
                                <td><div class="subtx" style="padding-left: 8px;padding-right: 8px;padding-top: 8px;margin: 0px auto";>
                                    <div align="right"><a href="current.htm" class="dblink">More &gt;&gt;</a></div>
                                </div></td>
                              </tr>
                          </table></td>
                          <td rowspan="3"><img src="img/invis.gif" width="5" height="5"></td>
                        </tr>
                        <tr>
                          <td valign="top"><table width="100%" height="350"  border="0" cellpadding="0" cellspacing="0" class="contentb">
                              <tr>
                                <td valign="top" bgcolor="#AFBFCB"><div class="subtx" style="padding-left: 6px;padding-right: 6px;padding-top: 8px;margin: 0px auto";>
                                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td><div align="center"></div></td>
                                      </tr>
                                      <tr>
                                        <td><div class="subtx" style="padding-left: 8px;padding-right: 8px;padding-top: 8px;margin: 0px auto";><span class="subHead">Team Effort Conquers a Non-Negotiable Deadline...</span></div></td>
                                      </tr>
                                      <tr>
                                        <td><div class="subtx" style="padding-left: 8px;padding-right: 8px;padding-top: 8px;margin: 0px auto";>
                                            <div align="right"><a href="61505_cutc.htm" class="dblink">More &gt;&gt;</a></div>
                                        </div></td>
                                      </tr>
                                    </table>
                                </div></td>
                              </tr>
                          </table>
                            </td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
-->
			</tr>
            </table>
            </td>
			                        </tr>
			                    </table></td>
			                  </tr>
                </table>
<p>
<p>

<!--
<table border="0" width="468" cellspacing="0" cellpadding="0"
style="border: 1px dashed rgb(0,0,0)">
  <tr>
    <td width="100%"><img src="images/petition.gif" width="468" height="60"></td>
  </tr>
  <tr>
    <td width="100%"><p align="center"><font face="Verdana"><small><small><small><font
    color="#FFFFFF">.</font></small></small></small><font color="#000000"><strong><br>
    <small><small>Total offenders in our database:</strong></font><font color="#FF0000"> <strong><%=OFFENDERS%></strong></font></small></small></font></td>
  </tr>
  <tr>
    <td width="100%"><p align="center">&nbsp;</p>
    <div align="center"><center><table border="0" width="90%" cellspacing="0" cellpadding="5">
      <tr>
        <td width="100%"><font face="Verdana" color="#000000"><small><small>Dear Visitor!</small></small></font><p><font
        face="Verdana" color="#000000">
        <small><small>Gas prices in the U.S. have risen an average of $0.85 over the past 2 years!  Just before the Gulf War, prices of $1.50/gallon of regular unleaded were high.  As of June 22nd, most stations are selling the same gas for $2.29/gallon.  Meanwhile, gas companies are seeing record-setting profits at our expense (see <a href="http://www.consumersunion.org/pub/core_other_issues/001086.html">this article</a> for more info).  It's time to do something about it! </small></small></font></p>
        <p>
        <bold>UPDATE:</bold> As of August 14th, the average price is now $2.65/gallon.  Isn't it time to do something??<p>
        <p><font face="Verdana" color="#000000"><small><small>If you belive that gas prices are too high, show it by signing this petition. Click the &quot;Sign the
        petition!&quot; link here below.</small></small></font></p>
        <p><font face="Verdana" color="#000000"><small><small>Please provide a correct e-mail
        address in order to be able to get the confirmation e-mail, otherwise your entry will be deleted.</small></small></font></td>
      </tr>
    </table>
    </center></div><p align="center"><a href="firma.asp" style="color: rgb(0,0,255)"
    target="_top"><strong><font face="Verdana"><small><small>Sign the petition!</small></small></font></strong></a></p>
    <p align="center"><font face="Verdana"><strong>
    <a href="webmaster.asp" style="color: rgb(0,0,255)" target="_blank"><small><small>Support us!</small></small></a><br>
    </strong><font color="#FFFFFF"><small><small>.</small></small></font></font></td>
  </tr>
  <tr>
    <td width="100%"><p align="center">&nbsp;</p>
    <p align="center"><font face="Verdana" color="#FF0000"></font><br>
    <font face="Verdana"><br>
    <font color="#FFFFFF"></font></font></td>
  </tr>
</table>
-->
<hr>
<%

'#########################################
'## SET SIGNATURE COUNTER DATABASE PATH ##
'#########################################

  SQL="SELECT * FROM tblCIOTI where OffenderID > 0 order by OffenderID"
  DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../CIOTI/database/petition.mdb")

   dim conntemp2, rstemp2
   set conntemp2=server.createobject("adodb.connection")
   conntemp2.open DSN
   set rstemp2=conntemp2.execute(SQL)

   if not(rstemp2.eof) THEN

   OFFENDERS=rstemp2(0)
   FName=rstemp2(1)
   LName=rstemp2(2)
   Addr1=rstemp2(3)
   City=rstemp2(5)
   State=rstemp2(6)
   Zip=rstemp2(7)
   Country=rstemp2(8)
   Offense=rstemp2(10)
   OffenderID=rstemp2(13)

   end if


%>

<table>
<tr>
<td>
<!--
<b><font size=4>Our Current Top 5 Offenders:</font></b>
</td>
</tr>
<%
if not rstemp2.eof then
	cnt = 0
	do while not rstemp2.eof

response.write "  <tr>"
response.write "    <td>"
response.write "      <font>" & rstemp2("OffenderID") & ") "& rstemp2("FName") & " " & rstemp2("LName") & ":   " & rstemp2("City") & ", " & rstemp2("State") & "  " & rstemp2("Zip") & " - " & rstemp2("Country") & "</font>"
response.write "    </td>"
response.write "  </tr>"
response.write "  <tr>"
response.write "    <td>"
response.write "      <font color=white>" & rstemp2("Offense") & "</font>"
response.write "    </td>"
response.write "  </tr>"
response.write "  <tr>"
response.write "    <td><p><p></td>"
response.write "  </tr>"
	rstemp2.movenext
	cnt = cnt + 1
	loop
end if
%>
-->
</table>

</center></div>
<%
rstemp2.Close()
Set rstemp2 = Nothing
%>
</body>
</html>
