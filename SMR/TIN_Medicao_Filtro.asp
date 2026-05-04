<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Dim rs1
Dim rs2
Dim strUsuario

Dim Busca

If trim(session("Usuario")) = "" Then
	Response.Redirect ("./TIN_Login.asp")
End if

strUsuario = session("Usuario")

'Abrindo uma conexão com o BD
set conConexao = TIN_AbrirConexaoBD()

strSql = "SP_LISTAR_CASES " 

set rs1 = Server.CreateObject("ADODB.RecordSet") '

rs1.OPEN STRSQL, conConexao

%>

<html>

<head>
	<title>Projeto Sinergia </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/sinergia.css">
<body topmargin="0" leftmargin="0" bgcolor=White text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="center">
			<p><b>
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema de Teste Integrado</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<form method="post" name="frmMedicao_Filtro" id="frmMedicao_Filtro">

<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0">
				<tr>
			    	<td bgcolor=White align="right" colspan="2" height="30">
						<p align="center">
				  		<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				  		<b>Por favor selecione o filtro</b>
				  		</font>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cases e Cenários Complementares:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcCase" style="WIDTH:550px;">
								<%If not rs1.EOF Then
									While not rs1.EOF %>
										<OPTION value="<%=UCase(rs1("Cases"))%>">
										<%=UCase(rs1("CasesXCASE"))%>
										<%=UCase("- ")%>
										<%=UCase(rs1("Desc_Case"))%>
										</OPTION>
										<%rs1.MoveNext
									WEnd			
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>
			</table>
		<td width="10%">&nbsp;</td>
	</tr>
</table>
<br>
<br>
<table width="100%" border="0">
	<td width="10%">&nbsp;</td>
	<td width="80%" align="center">
		<a href="javascript:Confirmar();"><img src="img/000050.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>
	</td>
	<td width="10%">&nbsp;</td>
</table>
<BR>
<BR>
<input type="hidden" id="strUsuario" name="strUsuario" value="<%=strUsuario%>">
</form>

<%
'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)
%>

<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmMedicao_Filtro.action = "TIN_Medicao_Detalhe.asp";
	document.frmMedicao_Filtro.submit();
}

function Redefinir()
{
	document.frmMedicao_Filtro.reset();
}


</SCRIPT>

</body>
</html>