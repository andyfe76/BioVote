<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Administrare voturi</TITLE>
</HEAD>
<BODY>

<%
mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
action=Request.QueryString("action")

if action="" Then
 ttop1
 %>
       <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Lege.</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Articol</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1><B>DA</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1><B>NU</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1><B>ABTIN</B></FONT></TD>
                </TR>
 <%
 Set record1=CreateObject("adodb.recordset")
 Set record2=CreateObject("adodb.recordset")
 record.ActiveConnection=mdb
 record.Source="SELECT * from legi"
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 If record.eof=False Then
 Do
 record1.ActiveConnection=mdb
 record1.Source="SELECT * from articole WHERE lege="&record.fields.item("id")
 record1.CursorType=0
 record1.CursorLocation=2
 record1.LockType=1
 record1.Open()
  %>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%=record.fields.item("pl")%>
                  </FONT></TD>
                <TD align=middle><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%="Toate"%>
                  </FONT></TD>
                <TD>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='DA' AND lege="&record.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                <td>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='NU' AND lege="&record.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                <td>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='Abtin' AND lege="&record.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                </TR>
  <%
 If record1.eof=False Then
 Do
 %>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%=record.fields.item("pl")%>
                  </FONT></TD>
                <TD align=middle><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%="Art. "&record1.fields.item("number")%>
                  </FONT></TD>
                <TD>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='DA' AND lege="&record.fields.item("id")&" AND articol="&record1.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                <td>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='NU' AND lege="&record.fields.item("id")&" AND articol="&record1.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                <td>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%
                record2.ActiveConnection=mdb
 				record2.Source="SELECT * from vot WHERE vot='Abtin' AND lege="&record.fields.item("id")&" AND articol="&record1.fields.item("id")
 				record2.CursorType=0
 				record2.CursorLocation=2
 				record2.LockType=1
 				record2.Open()
 				nr=0
 				If record2.eof=False Then
 				Do
 				 nr=nr+1
 				 record2.movenext
 				Loop Until record2.eof
 				End If
 				record2.close
				Response.Write(nr)
                %>
                </FONT></TD>
                </TR>
  <%
   record1.movenext
  Loop Until record1.eof
  End If
  record1.close
  record.movenext
  Loop Until record.eof
 End If
 record.close
 %>
			</table>
		</td>
	  </tr>
	 </table>
<%
 
End If
%>

</BODY>
</HTML>

<%
Function ttop1
%>
<TABLE cellSpacing=0 cellPadding=0 width=600 align=center border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR vAlign=bottom>
          <TD><IMG src="images/bannt.jpg" border=0></A></TD>
          <TD></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD><IMG src="images/bann1.jpg" border=0></TD></TR>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR vAlign=top>
          <TD><IMG src="images/bannd.jpg" border=0></A></TD>
          <TD></TD></TR></TBODY></TABLE></TD></TR>
        </TBODY>
      </TABLE>

<TABLE cellSpacing=0 cellPadding=0 width=600 align=center border=0>
  <TBODY>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face=Verdana,Arial size=1><B>
      <P>
      <FONT color=#999999>Sunteti în sectiunea: </FONT>
      <A href="index.asp">Acasa</a>
       >
      Administrare voturi
      </FONT></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function
%>