<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Administrare articole</TITLE>
</HEAD>
<BODY>

<%
mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
action=Request.QueryString("action")

If action="edit" Then
 art=Request.QueryString("art")
 record.ActiveConnection=mdb
 record.Source="SELECT * from articole WHERE id="&art
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 law=record.fields.item("lege")
 law_art=record.fields.item("number")
 record.close
 
 record.ActiveConnection=mdb
 record.Source="SELECT * from legi WHERE id="&law
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 ttop1 law,record.fields.item("pl"),law_art
 record.close
 
 record.ActiveConnection=mdb
 record.Source="SELECT * from articole WHERE id="&art
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 
  %>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_art.asp?action=edit_confirm&art=<%=art%>&law=<%=law%>">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan="2"><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Numar.</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descriere
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle colSpan="2"><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="number" size="20" value="<%=record.fields.item("number")%>">
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
			</table>
</form>
		</td>
	  </tr>
	 </table>
<%
record.close
End If

If action="edit_confirm" Then
 art=Request.QueryString("art")
 law=Request.QueryString("law")
 nr=Request.Form("number")
 desc=Request.Form("desc")
 conn.open mdb
 conn.Execute("UPDATE articole SET number="&nr&",`desc`='"&desc&"' WHERE id="&art)
 conn.close
 Response.Redirect("admin_law.asp?action=edit&law="&law)
End If

If action="add" Then
 law=Request.QueryString("law")
 record.ActiveConnection=mdb
 record.Source="SELECT * from legi WHERE id="&law
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 ttop2 law,record.fields.item("pl")
 record.close
   %>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_art.asp?action=add_confirm&law=<%=law%>">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan="2"><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Numar</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Descriere
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle colSpan="2"><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="number" size="20">
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <textarea rows="2" name="desc" cols="53" align="left"></textarea>
                </FONT></TD></TR>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="submit" value="Modifica" name="B1">
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
 law=Request.QueryString("law")
 desc=Request.Form("desc")
 nr=Request.Form("number")
 conn.open mdb
 conn.Execute("INSERT INTO articole (lege,number,`desc`) VALUES ("&law&","&nr&",'"&desc&"')")
 conn.close
 Response.Redirect("admin_law.asp?action=edit&law="&law)
End If

%>

</BODY>
</HTML>

<%
Function ttop1(id,txt,txt2)
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
      <A href="index.asp">Acasa</a> > <a href="admin_law.asp">Administrare legi</a> > <a href="admin_law.asp?action=edit&law=<%=id%>">Modifica <%=txt%></a> > Modifica articol <%=txt2%></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

Function ttop2(id,txt)
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
      <A href="index.asp">Acasa</a> > <a href="admin_law.asp">Administrare legi</a> > <a href="admin_law.asp?action=edit&law=<%=id%>">Modifica <%=txt%></a> > Adauga articol</P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

%>