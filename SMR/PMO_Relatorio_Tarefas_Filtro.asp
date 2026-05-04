<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

Dim strFrente
Dim vetDados

strFrente = Trim(Request("slcFrente"))

If strFrente <> "" Then

	Select Case strFrente
		Case "TI"
			vetDados = Array("TI")
		Case "GM"
			vetDados = Array("GM-AO", "GM-CO", "GM-MO", "GM-SUP", "GM-TR")
		Case "INT"
			vetDados = Array("INT")
		Case "FI"
			vetDados = Array("INT-ECCS", "INT-FT", "INT-FI")
		Case "SD+P&C"
			vetDados = Array("INT-SD", "INT-PC")
		Case "PX+MES"
			vetDados = Array("INT-LT", "INT-MES", "INT-CBS", "INT-PM", "INT-PO", "INT-PP", "INT-PS", "INT-QM")
	End Select

End If


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
			<!--<font size="3" face="Verdana" color="#000000">-->
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe Sinergia</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<form method="post" name="frmRelatorio_Tarefas_Filtro" id="frmRelatorio_Tarefas_Filtro">

<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0">
				<tr>
			    	<td bgcolor=White align="right" colspan="2" height="30">
						<p align="center">
				  		<!--<font face="Arial" size="2" color="#FFFFFF">-->
				  		<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				  		<b>Por favor selecione o filtro</b>
				  		</font>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Frente/Sub-Frente:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcFrente" style="WIDTH:300px;" onchange="Selfrente();">
								<OPTION value=""></OPTION>
								<OPTION value="TI" <%If strFrente = "TI" Then%>Selected<%End If%> >TI</OPTION>
								<OPTION value="GM" <%If strFrente = "GM" Then%>Selected<%End If%>>GM</OPTION>
								<OPTION value="INT" <%If strFrente = "INT" Then%>Selected<%End If%>>INT</OPTION>
								<OPTION value="FI" <%If strFrente = "FI" Then%>Selected<%End If%>>FI</OPTION>
								<OPTION value="SD+P&C" <%If strFrente = "SD+P&C" Then%>Selected<%End If%>>SD+P&C</OPTION>
								<OPTION value="PX+MES" <%If strFrente = "PX+MES" Then%>Selected<%End If%>>PX+MES</OPTION>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipes:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcEquipe" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
								<%If IsArray(vetDados) Then
									For I = 0 to ubound(vetDados)%>	
										<OPTION value="<%=UCase(vetDados(I))%>">
										<%=UCase(vetDados(I))%>
										</OPTION>
								  <%Next
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>
			  	<tr>
			  	
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
		<a href="javascript:ConfirmarExcel();"><img src="img/000047.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>
	</td>
	<td width="10%">&nbsp;</td>
</table>
<BR>
<BR>

</form>

<SCRIPT language=JavaScript>

function Selfrente()
{
	document.frmRelatorio_Tarefas_Filtro.action = "PMO_Relatorio_Tarefas_Filtro.asp";
	document.frmRelatorio_Tarefas_Filtro.submit();

}


function Confirmar()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Tarefas_Filtro.action = "PMO_Relatorio_Tarefas_Detalhe.asp";
		document.frmRelatorio_Tarefas_Filtro.submit();

	}
}

function ConfirmarExcel()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Tarefas_Filtro.action = "PMO_Relatorio_Tarefas_Detalhe_Excel.asp";
		document.frmRelatorio_Tarefas_Filtro.submit();
	}

}

function Redefinir()
{
	document.frmRelatorio_Tarefas_Filtro.reset();
}

function ValidarCombos()
{
	if (document.frmRelatorio_Tarefas_Filtro.slcFrente.value == "")
	{
		alert("Selecionar no mínimo uma frente/sub-frente");
		return false;
	}
	else
	{
		return true;
	}		

}

</SCRIPT>

</body>
</html>
