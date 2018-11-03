<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Administrare legi</TITLE>
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
                  <A href="admin_law.asp?action=edit&law=<%=record.fields.item("id")%>"><%=record.fields.item("pl")%></a>
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
	 <A href="admin_law.asp?action=add">Adauga lege</a>
<%
End If

If action="add" Then
 ttop2
 %>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_law.asp?action=add_confirm">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan="2"><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>P.L.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descriere
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle colSpan="2"><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="pl" size="20">
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <textarea rows="2" name="desc" cols="53"></textarea>
                </FONT></TD></TR>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="submit" value="Adauga" name="B1">
                  </FONT></TD>
                <TD colSpan=4>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                
                </FONT></TD></TR>

			</table>
</form>
		</td>
	  </tr>
	 </table>

<%
End If

If action="add_confirm" Then
 pl=Request.Form("pl")
 desc=Request.Form("desc")
 conn.open mdb
 conn.Execute("INSERT INTO legi (`pl`,`desc`) VALUES ('"&pl&"','"&desc&"')")
 conn.close
 Response.Redirect("admin_law.asp")
End If

If action="edit" Then
 law=Request.QueryString("law")
 record.ActiveConnection=mdb
 record.Source="SELECT * from legi WHERE id="&law
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 ttop3(record.fields.item("pl"))
  %>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_law.asp?action=edit_confirm&law=<%=law%>">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan="2"><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>P.L.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descriere
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle colSpan="2"><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="pl" size="20" value="<%=record.fields.item("pl")%>">
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <textarea rows="2" name="desc" cols="53" align="left"><%=record.fields.item("desc")%></textarea>
                </FONT></TD></TR>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="submit" value="Modifica" name="B1">
                  </FONT></TD>
                <TD colSpan=4>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                
                </FONT></TD></TR>
<%
 record.close
 record.ActiveConnection=mdb
 record.Source="SELECT * from articole WHERE lege="&law&" ORDER BY number"
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 If record.eof=False Then
%>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan="2"><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Nr.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descriere
                  </B></FONT></TD></TR>
<%
Do
%>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle colSpan="2"><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%=record.fields.item("number")%>
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <A href="admin_art.asp?action=edit&art=<%=record.fields.item("id")%>"><%=record.fields.item("desc")%></a>
                </FONT></TD></TR>
<%
 record.movenext
 Loop Until record.eof
 End If
 record.close
%>

			</table>
</form>
		</td>
	  </tr>
	 </table>
<BR>
<A href="admin_art.asp?action=add&law=<%=law%>">Adauga articol</a>

<%
End If

If action="edit_confirm" Then
 law=Request.QueryString("law")
 pl=Request.Form("pl")
 desc=Request.Form("desc")
 conn.open mdb
 conn.Execute("UPDATE legi SET pl='"&pl&"',`desc`='"&desc&"' WHERE id="&law)
 conn.close
 Response.Redirect("admin_law.asp")
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
      <A href="index.asp">Acasa</a> > Administrare legi</P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

Function ttop2
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
      <A href="index.asp">Acasa</a> > <a href="admin_law.asp">Administrare legi</a> > Adauga</P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

Function ttop3(txt)
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
      <A href="index.asp">Acasa</a> > <a href="admin_law.asp">Administrare legi</a> > Modifica <%=txt%></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function
%>