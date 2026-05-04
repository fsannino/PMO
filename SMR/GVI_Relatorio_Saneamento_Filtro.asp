<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000
	
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
			<TD  width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<b>
					<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
						TEXTO
					</font>
				</b>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<img src="img_Saneamento/saneamento.gif" name="Saneamento" border="0">
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>


		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=AB" class="conf">
						ABAST
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=EP" class="conf">
						E&P
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=FI" class="conf">
						FINANÇAS
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>


		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=GE" class="conf">
						GAS & ENERGIA
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=IT" class="conf">
						INTERNACIONAL
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
			</TD>
		</tr>

		<tr  align=center>
			<TD width=33%>
				&nbsp;
			</TD>
			<td  align=center>
				<font face="Arial" size="1">
					<a href="./GVI_Relatorio_Saneamento_Detalhe.asp?strOpcao=SC" class="conf">
						SERVIÇOS COMPARTILHADOS
					</a>
				</font>
			</td>
			<TD width=33%>
				&nbsp;
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