<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Ordinea de zi</TITLE>
</HEAD>
<BODY>

<%
mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
action=Request.QueryString("action")

If action="" Then
 ttop1
 record.ActiveConnection=mdb
 record.Source="SELECT * from legi"
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
                  size=1><B>Nr.</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>P.L.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descrierea punctului de pe ordinea de zi
                  </B></FONT></TD></TR>
 <%
 Do
 %>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%=record.fields.item("id")%>
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <A href="law.asp?id=<%=record.fields.item("id")%>"><%=record.fields.item("pl")%></a>
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%=record.fields.item("desc")%>
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
	 <BR>
	 <A href="admin_law.asp">Administrare legi</a>
	 <BR>
	 <A href="admin_users.asp">Administrare utilizatori</a>
	 <BR>
	 <A href="admin_vote.asp">Administrare voturi</a>
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
      <A href="index.asp">Acasa</a></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function
%>