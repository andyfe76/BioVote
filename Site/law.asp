<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Legi</TITLE>
</HEAD>
<BODY>
<%

mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
law=Request.QueryString("id")

record.ActiveConnection=mdb
record.Source="SELECT * from legi where id="&law
record.CursorType=0
record.CursorLocation=2
record.LockType=1
record.Open()
ttop1(record.fields.item("pl"))


%>
<BR>
<center><A href="vot.asp?law=<%=law%>">Voteaza legea <%=record.fields.item("pl")%></a></center>
<BR>
<%
record.close
record.ActiveConnection=mdb
record.Source="SELECT * from articole where lege="&law
record.CursorType=0
record.CursorLocation=2
record.LockType=1
record.Open()
If record.eof=False Then
%>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Art.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Descriere.</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Voteaza
                  </B></FONT></TD></TR>
 <%
 Do
 %>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  Art. <%=record.fields.item("id")%>
                  </FONT></TD>
                <TD align=middle colSpan=3><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%=record.fields.item("desc")%>
                  </FONT></TD>
                <TD>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <A href="vot.asp?law=<%=law%>&art=<%=record.fields.item("id")%>">Voteaza</a>
                </FONT></TD>
                </TR>
  <%
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



%>

</BODY>
</HTML>




<%
Function ttop1(txt)
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
      <A href="index.asp">Acasa</a> > Legea <%=txt%>

      </P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function
%>