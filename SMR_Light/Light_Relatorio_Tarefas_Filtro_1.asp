<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

Dim strFrente
Dim rs
Dim rs1

Dim cmdResultado

	strFrente = Trim(Request("slcFrente"))

	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_FRENTES"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

	If strFrente <> "" Then

	    Set cmdResultado = Server.CreateObject("ADODB.Command")
	        
	    With cmdResultado
	    
	        .ActiveConnection = conConexao
	        .CommandType = 4
			.CommandTimeout = 600
	        .CommandText = "SP_LISTAR_EQUIPES_PROJ"

			.Parameters.Refresh

			.Parameters(1).Value = strFrente

	    End With

		set rs1 = Server.CreateObject("ADODB.RecordSet")

		set rs1 = cmdResultado.Execute()

	End If


%>

<html>

<head>
	<title>Projeto BRACUSS </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/Light.css">
<body topmargin="0" leftmargin="0" bgcolor=White text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="center">
			<p><b>
			<!--<font size="3" face="Verdana" color="#000000">-->
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe BRACUSS</font>
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
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Frente:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcFrente" style="WIDTH:300px;" onchange="Selfrente();">
								<OPTION value="" <%If strFrente = "" Then%>Selected<%End If%>></OPTION>
								<%Do While Not rs.EOF%>
									<OPTION value="<%=rs("Frente")%>" <%If strFrente = rs("Frente") Then%>Selected<%End If%>><%=rs("Frente")%></OPTION>
									<%rs.MoveNext%>
								<%Loop%>
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
								<%If strFrente <> "" Then%>
									<%Do While Not rs1.EOF%>
										<OPTION value="<%=rs1("Equipe")%>"><%=rs1("Equipe")%></OPTION>
										<%rs1.MoveNext%>
									<%Loop%>
								<%End If%>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcOpcao" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
						        <option value="1">HOJE</option>
						        <option value="2"> + 7 DIAS</option>
						        <option value="3"> + 14 DIAS</option>
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
	document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Filtro.asp";
	document.frmRelatorio_Tarefas_Filtro.submit();

}


function Confirmar()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Detalhe.asp";
		document.frmRelatorio_Tarefas_Filtro.submit();

	}
}

function ConfirmarExcel()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Detalhe_Excel.asp";
		document.frmRelatorio_Tarefas_Filtro.submit();
	}

}

function Redefinir()
{

	document.frmRelatorio_Tarefas_Filtro.slcFrente.value = "";
	document.frmRelatorio_Tarefas_Filtro.slcEquipe.value = "";
		
	document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Filtro.asp";
	document.frmRelatorio_Tarefas_Filtro.submit();

/*	document.frmRelatorio_Tarefas_Filtro.reset();*/
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
