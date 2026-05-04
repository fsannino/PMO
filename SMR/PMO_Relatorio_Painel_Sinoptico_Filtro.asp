<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->


<html>
<head>
	<title>Projeto Sinergia </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/sinergia.css">
<body topmargin="0" leftmargin="0" bgcolor=white text="#000000" link="#0000ff" vlink="#0000ff" alink="#0000ff">
<form method="post" name="frmRelatorio_Painel_Sinoptico_Filtro" id="frmRelatorio_Painel_Sinoptico_Filtro">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="middle">
			<p><b>
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Painel Sinóptico</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0" >
				<tr>
			    	<td bgcolor=white align="right" colspan="2" height="30">
						<p align="center">
				  			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">
				  				<b>Por favor selecione o filtro</b>
				  			</font>
				  		</p>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="middle">
						<b>
						<font color=white size="1" face="Georgia, Times New Roman, Times, serif">Diretoria:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
							<SELECT size=1 name="slcDiretoria" style="WIDTH:303px;">
								<OPTION value="" selected>--------------------------- Todos--------------------------</OPTION> 
								<OPTION value="CO">CONSOLIDADO</OPTION>
								<OPTION value="AB">ABAST</OPTION> 
								<OPTION value="EP">E&P</OPTION> 
								<OPTION value="FI">FINANÇAS</OPTION> 
								<OPTION value="GE">GAS E ENERGIA</OPTION> 
								<OPTION value="IN">INTERNACIONAL</OPTION> 
								<OPTION value="SE">SERVIÇOS</OPTION>
								<OPTION value="PR">PRESIDENCIA</OPTION>
							</SELECT>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b><!--<font color="#000000" size="1" face="Arial">-->
						<font color=white size="1" face="Georgia, Times New Roman, Times, serif">Atividade:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc><!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcAtividade" style="WIDTH: 303px">
								<option value="" selected>--------------------------- Todos --------------------------</option>
<!--								<option value="TC">TESTE DE CAMPO</option>-->
<!--				<option value="SA">SANEAMENTO</option>-->
						        <option value="TR">TREINAMENTO</option>
<!--								<option value="CO">COMUNICAÇÃO</option>-->
<!--								<option value="MO">MOBILIZAÇÃO</option>-->
<!--								<option value="IM">IMPACTOS</option>-->
								<option value="SP">SUPORTE</option>
								<option value="SP-CAP-PER">&nbsp;&nbsp;- CAP. Perfil</option>
								<option value="SP-CAP-SUP">&nbsp;&nbsp;- CAP. Suporte</option>
<!--								<option value="SP-SL-CTR">&nbsp;&nbsp;- Sala de Controle</option>-->
								<option value="IN">INFRA</option>
								<option value="IN-GERAL">&nbsp;&nbsp;- GERAL</option>
								<option value="IN-NF">&nbsp;&nbsp;- NF</option>

						    </select>
						</font>
					</td>
			  	</tr>
			  	<br>
				<br>
			</table>
		</TD>
		<td width="10%">&nbsp;</td>
	</TR>
</TABLE>
<BR>
<table width="100%" border="0">
	<td width="10%">&nbsp;</td>
	<td width="80%" align="center">
		<a href="javascript:Confirmar();"><img src="img/000050.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>
	</td>
	<td width="10%">&nbsp;</td>
</table>

</form>

<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmRelatorio_Painel_Sinoptico_Filtro.action = "PMO_Relatorio_Painel_Sinoptico_Detalhe.asp";
	document.frmRelatorio_Painel_Sinoptico_Filtro.submit();
}

function Redefinir()
{
	document.frmRelatorio_Painel_Sinoptico_Filtro.reset();
}


</SCRIPT>

</body>
</html>
