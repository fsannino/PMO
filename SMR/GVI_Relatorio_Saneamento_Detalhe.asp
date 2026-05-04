<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim strOpcao
Dim strArquivo

strOpcao = Request("strOpcao")

Select Case strOpcao
	Case "AB"
		strArquivo = "abast.gif"
	Case "EP"
		strArquivo = "ep.gif"
	Case "FI"
		strArquivo = "finanças.gif"
	Case "GE"
		strArquivo = "ge.gif"
	Case "IT"
		strArquivo = "inter.gif"
	Case "SC"
		strArquivo = "serv.gif"
End Select
%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Saneamento_Filtro" id="frmRelatorio_Saneamento_Filtro" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<p>

	<table  width=100%>
		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<img src="img_Saneamento/<%=strArquivo%>" name="Saneamento" border="0">
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

	</table>

	<BR>

	<table>
		<tr  align=center>
			<TD width=675px>
				&nbsp;
			</TD>
			<TD align=right>
				<a href="javascript:history.go(-1);" ><img src="img/000024.gif" width="73" height="16" border="0" align="absmiddle" tabindex="4"></a>
			</TD>
		</tr>

	</table>


		<p align="right">
		<!--<input type="button" name="cmdSubmit" value="Enviar" onclick="Confirmar();">-->
		<BR>
		<hr>

</FORM>
</body>
</html>