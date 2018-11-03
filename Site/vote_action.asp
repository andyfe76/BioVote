<%@ LANGUAGE="VBSCRIPT" %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="SAPIEN Technologies PrimalSCRIPT(TM)">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Vote</TITLE>
</HEAD>
<BODY>

<%
mdb="dsn=vote"
Set record=CreateObject("adodb.recordset")
Set conn=CreateObject("adodb.connection")
law=Request.QueryString("law")
art=Request.QueryString("art")
If art="" Then art="0"
da=Request.Form("da")
nu=Request.Form("nu")
abtin=Request.Form("abtin")
biodata=Request.Form("bio")

vot=Request.Form("vot")

raspuns=""

record.ActiveConnection=mdb
record.Source="SELECT * from legi where id="&law
record.CursorType=0
record.CursorLocation=2
record.LockType=1
record.Open()
law_txt=record.fields.item("pl")
record.close

record.ActiveConnection=mdb
record.Source="SELECT * from user"
record.CursorType=0
record.CursorLocation=2
record.LockType=1
record.Open()
Do
 user_bio=record.fields.item("biodata")
 if compare(user_bio,biodata)=0 Then
  user_id=record.fields.item("id")
  nume=record.fields.item("prenume")&" "&record.fields.item("nume")
  Exit Do
 End If 
 record.movenext
Loop Until record.eof
record.close

If user_id<>0 Then
 id=user_id
 record.ActiveConnection=mdb
 record.Source="SELECT * from vot where user="&id&" AND lege="&law&" AND articol="&art
 record.CursorType=0
 record.CursorLocation=2
 record.LockType=1
 record.Open()
 If record.eof=True Then
  conn.open(mdb)
  conn.Execute("INSERT INTO vot (lege,articol,user,vot) VALUES ("&law&","&art&","&id&",'"&vot&"')")
  conn.close
  raspuns=nume&", va multumim ca ati votat cu "&vot
 Else
  raspuns=nume&", ati mai votat o data!"
 End If
Else
 raspuns="Nu ati fost recunoscut!"
End If 


ttop1 "Legea "&law_txt,law
%>
      <TABLE bgColor=#cccccc border=0 align="center">
        <TBODY>
        <TR>
          <TD>
            <TABLE cellSpacing=1 border=0>
              <TBODY>
              <TR vAlign=center align=middle bgColor=#ffffff>
                <TD colSpan=5><FONT face="Verdana, Geneva, Arial, Sans-Serif" size="3" 
                  size=1><B><center><%=raspuns%></center></B></FONT></TD>
                </TR>
			</table>
		</td>
	  </tr>
	 </table>


</BODY>
</HTML>


<%
Function compare(user_bio,biodata)
 Set t1 = CreateObject("Fapisrv.Biosrv")
 compare=t1.compare(user_bio,biodata)
 set t1=nothing
End Function

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
%>