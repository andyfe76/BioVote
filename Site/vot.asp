<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Voteaza</TITLE>

<OBJECT ID="data1" CLASSID="clsid:7739872A-CF44-4BC5-A8B5-F1086A0585D6" CODEBASE="fapicl.cab"> 

 
</OBJECT> 

<SCRIPT LANGUAGE="VBScript">
<!--
Sub da_OnClick
 Div1.InnerHTML ="<BR><BR><center><table border=""3"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#0000FF""><tr><td width=""100%"" bgcolor=""#808080"" bordercolor=""#0000FF"" bordercolorlight=""#0000FF"" bordercolordark=""#0000FF"" style=""border: 3px groove #0000FF""><b><font face=""Verdana"" size=""4"" color=""#FFFFFF"">Puneti degetul pe senzor</font></b></td></tr></table></center>"
 min=data1.capture()
 document.form1.bio.value=min
 document.form1.vot.value="DA"
 document.form1.submit()
End Sub

Sub abtin_OnClick
 Div1.InnerHTML = "<BR><BR><center><table border=""3"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#0000FF""><tr><td width=""100%"" bgcolor=""#808080"" bordercolor=""#0000FF"" bordercolorlight=""#0000FF"" bordercolordark=""#0000FF"" style=""border: 3px groove #0000FF""><b><font face=""Verdana"" size=""4"" color=""#FFFFFF"">Puneti degetul pe senzor</font></b></td></tr></table></center>"
 min=data1.capture()
 document.form1.bio.value=min
 document.form1.vot.value="ABTIN"
 document.form1.submit()
End Sub

Sub nu_OnClick
 Div1.InnerHTML ="<BR><BR><center><table border=""3"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#0000FF""><tr><td width=""100%"" bgcolor=""#808080"" bordercolor=""#0000FF"" bordercolorlight=""#0000FF"" bordercolordark=""#0000FF"" style=""border: 3px groove #0000FF""><b><font face=""Verdana"" size=""4"" color=""#FFFFFF"">Puneti degetul pe senzor</font></b></td></tr></table></center>"
 min=data1.capture()
 document.form1.bio.value=min
 document.form1.vot.value="NU"
 document.form1.submit()
End Sub
-->
</SCRIPT>

<script language="JavaScript">
 function abtin()
 {
  document.getElementById('Div1').innerHTML = "Capture in progress ....";
  min = data1.capture()
  document.form1.bio.value=min
  document.form1.vot.value="ABTIN"
  document.form1.submit()
  } 
</script>

</HEAD>
<BODY>



<%
'set signer=CreateObject("Scripting.Signer")
'file="C:\Inetpub\wwwroot\vote3\vot.asp"
'cert="C:\Inetpub\wwwroot\vote3\andy2.cer"
'store="my"
'Signer.SignFile File, Cert, Store

mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
law=Request.QueryString("law")
art=Request.QueryString("art")

record.ActiveConnection=mdb
record.Source="SELECT * from legi where id="&law
record.CursorType=0
record.CursorLocation=2
record.LockType=1
record.Open()
law_pl=record.fields.item("pl")
record.close

If art="" Then
 ttop1 "Legea "&law_pl,law
Else
 record.ActiveConnection=mdb
 record.Source="SELECT * from articole where id="&art
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 ttop2 "Legea "&law_pl,art,law
 record.close
End If


%>

      <TABLE bgColor=#cccccc border=0 align="center">
        <TBODY>
        <TR>
          <TD>
<form method="POST" name="form1" action="vote_action.asp?law=<%=law%>&art=<%=art%>">
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                  <input type="button" value="DA" name="da">
                  </B></FONT></TD>
                <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                  <input type="button" value="NU" name="nu">
                  </B></FONT></TD>
                <TD colSpan=3><FONT face="Verdana, Geneva, Arial, Sans-Serif" 
                  size=1><B>
                  &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                  <input type="button" value="Ma abtin" name="abtin" >
                  <input type="hidden" name="bio">
                  <input type="hidden" name="vot">
                  </B></FONT></TD></TR>
			</table>
</form>
		</td>
	  </tr>
	  
	 </table>

<div id=Div1>
aa
</div>

</BODY>
</HTML>

<%
Function ttop1(txt,id)
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
      <A href="index.asp">Acasa</a> >
      <A href="law.asp?id=<%=id%>"><%=txt%></a> > Voteaza

      </P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function

Function ttop2(txt,art,law)
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
      <A href="index.asp">Acasa</a> >
      <A href="law.asp?id=<%=law%>"><%=txt%></a> > Articolul <%=art%> > Voteaza

      </P></B></TD></TR>
  <TR>
    <TD height=10><IMG height=2 
      src="images/pix_cccccc.gif" width=600 border=0></TD></TR>
  <TR>
    <TD><FONT face="Verdana, Geneva, Arial, Sans-Serif" size=1>
<%
End Function
%>