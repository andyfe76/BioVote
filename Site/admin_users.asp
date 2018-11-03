<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Administrare utilizatori</TITLE>
<OBJECT ID="data1" CLASSID="clsid:7739872A-CF44-4BC5-A8B5-F1086A0585D6" CODEBASE="fapicl.cab"> 

 
</OBJECT> 

<SCRIPT LANGUAGE="VBScript">
<!--
Sub enroll_OnClick
 Div1.InnerHTML = "<BR><BR><center><table border=""3"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#0000FF""><tr><td width=""100%"" bgcolor=""#808080"" bordercolor=""#0000FF"" bordercolorlight=""#0000FF"" bordercolordark=""#0000FF"" style=""border: 3px groove #0000FF""><b><font face=""Verdana"" size=""4"" color=""#FFFFFF"">Puneti degetul pe senzor</font></b></td></tr></table></center>"
 min=data1.capture()
 document.form1.biodata.value=min
End Sub

-->
</SCRIPT>

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
 record.Source="SELECT * from user"
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
                  size=1><B>Voturi</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Nume
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
                  <%
                	Set record1=CreateObject("adodb.recordset")
                	record1.ActiveConnection=mdb
 					record1.Source="SELECT * from vot WHERE user="&record.fields.item("id")
 					record1.CursorType=0
 					record1.CursorLocation=2
 					record1.LockType=1
 					record1.Open()
 					nr=0
 					If record1.eof=False Then
 					Do
 					 nr=nr+1
 					 record1.movenext
 					Loop Until record1.eof 
 					End If
 					record1.close
 					Set record1=Nothing
                  %>
                  <A href="admin_users.asp?action=view_vote&user=<%=record.fields.item("id")%>"><%=nr%></a>
                  </FONT></TD>
                <TD colSpan=2>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <A href="admin_users.asp?action=edit&id=<%=record.fields.item("id")%>"><%=record.fields.item("prenume")+" "+record.fields.item("nume")%></a>
                </FONT></TD>
                <td>
                <A href="admin_users.asp?action=delete&id=<%=record.fields.item("id")%>">
                <IMG border="0" src="images/del_1.jpg" height="10" align="centre">
                </a>
                </td>
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
	 <A href="admin_users.asp?action=add">Adauga utilizator</a>
<%
End If

If action="add" Then
ttop2
%>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_users.asp?action=add_confirm" name="form1">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Nume</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Prenume</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Amprenta
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="nume" size="20">
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="prenume" size="20">
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <input type="text" name="biodata" cols="20">
                <input type="button" value="Enroll" name="enroll">
                </FONT></TD></TR>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="submit" value="Adauga" name="B1">
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                
                </FONT></TD></TR>

			</table>
</form>
		</td>
	  </tr>
	 </table>

<div id=Div1>
asas
</div>

<%
End If

If action="add_confirm" Then
 nume=Request.Form("nume")
 prenume=Request.Form("prenume")
 biodata=Request.Form("biodata")
 conn.open(mdb)
 conn.Execute("INSERT INTO user (nume,prenume,biodata) VALUES ('"&nume&"','"&prenume&"','"&biodata&"')")
 conn.close
 Response.Redirect("admin_users.asp")
End If



If action="edit" Then
ttop3
 id=Request.QueryString("ID")
 record.ActiveConnection=mdb
 record.Source="SELECT * from user where id="&id
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()

%>
      <TABLE bgColor=#cccccc border=0>
        <TBODY>
        <TR>
          <TD>
<form method="POST" action="admin_users.asp?action=edit_confirm&id=<%=id%>" name="form1">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Nume</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Prenume</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Amprenta
                  </B></FONT></TD></TR>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="nume" size="20" value="<%=record.fields.item("nume")%>">
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="text" name="prenume" size="20" value="<%=record.fields.item("prenume")%>">
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <input type="text" name="biodata" size="20" value="<%=record.fields.item("biodata")%>">
                <input type="button" value="Enroll" name="enroll">
                </FONT></TD></TR>

              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <input type="submit" value="Modifica" name="B1">
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  
                  </FONT></TD>
                <TD colSpan=3>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                
                </FONT></TD></TR>

			</table>
</form>
		</td>
	  </tr>
	 </table>


</form>
		</td>
	  </tr>
	 </table>
<div id=Div1>

</div>


<%
End If

If action="edit_confirm" Then
 id=Request.QueryString("ID")
 nume=Request.Form("nume")
 prenume=Request.Form("prenume")
 biodata=Request.Form("biodata")
 conn.open(mdb)
 conn.Execute("UPDATE user SET nume='"&nume&"',prenume='"&prenume&"',biodata='"&biodata&"' WHERE ID="&id)
 conn.close
 Response.Redirect("admin_users.asp")
End If

If action="delete" Then
 id=Request.QueryString("ID")
 conn.open(mdb)
 conn.Execute("DELETE FROM user WHERE ID="&id)
 conn.close
 Response.Redirect("admin_users.asp")
End If




If action="view_vote" Then
 id=Request.QueryString("user")
 record.ActiveConnection=mdb
 record.Source="SELECT * from user where id="&id
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 txt=record.fields.item("prenume")&" "&record.fields.item("nume")
 record.close
 ttop4(txt)
 
 record.ActiveConnection=mdb
 record.Source="SELECT * from vot where user="&id
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
                  size=1><B>Lege</B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>Articol</B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  Vot
                  </B></FONT></TD></TR>
 <%
 Do
 %>
              <TR vAlign=top bgColor=#ffffff>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%
                  Set record1=CreateObject("adodb.recordset")
                  record1.ActiveConnection=mdb
 				  record1.Source="SELECT * from legi where id="&record.fields.item("lege")
 				  record1.CursorType=0
 				  record1.CursorLocation=2
 				  record1.LockType=1
 				  record1.Open()
                  Response.Write(record1.fields.item("pl"))
                  record1.close
                  Set record1=nothing
                  %>
                  </FONT></TD>
                <TD align=middle><FONT 
                  face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                  <%
                  If record.fields.item("articol")<> 0 Then
                   Set record1=CreateObject("adodb.recordset")
                   record1.ActiveConnection=mdb
 				   record1.Source="SELECT * from articole where id="&record.fields.item("articol")
 				   record1.CursorType=0
 				   record1.CursorLocation=2
 				   record1.LockType=1
 				   record1.Open()
 				   Response.Write(record1.fields.item("number"))
 				   record1.close
 				   Set record1=Nothing
 				  Else
 				   Response.Write("Toate")
 				  End If

                  %>
                  </FONT></TD>
                <TD colSpan=2>
                <FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
                <%=record.fields.item("vot")%>
                </FONT></TD>
                <td>
                <A href="admin_users.asp?action=delete_vot&id=<%=record.fields.item("id")%>">
                <IMG border="0" src="images/del_1.jpg" height="10" align="centre">
                </a>
                </td>
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
 
 
End If


If action="delete_vot" Then
 id=Request.QueryString("id")
 conn.open(mdb)
 conn.Execute("DELETE FROM vot WHERE id="&id)
 conn.close
 Response.Redirect("admin_users.asp")
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
      Administrare utilizatori
      </FONT></P></B></TD></TR>
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
      <A href="index.asp">Acasa</a>
       >
      <A href="admin_users.asp">Administrare utilizatori</a>
      >
      Adauga
      </FONT></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function


Function ttop3
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
      <A href="admin_users.asp">Administrare utilizatori</a>
      >
      Editeaza
      </FONT></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

Function ttop4(txt)
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
      <A href="admin_users.asp">Administrare utilizatori</a>
      >
      Situatie voturi pentru <%=txt%>
      </FONT></P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

%>