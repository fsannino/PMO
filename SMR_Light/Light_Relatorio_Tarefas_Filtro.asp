<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

Dim strProj
Dim strFrente
Dim strEquipe
Dim strResponsavel
Dim strOpcao

Dim rs
Dim rs1
Dim rs2
Dim rs3

Dim cmdResultado

	strProj   = Trim(Request("slcProjeto"))
	strFrente = Trim(Request("slcFrente"))
	strEquipe = Trim(Request("slcEquipe"))
	strResponsavel = Trim(Request("slcResponsavel"))
	strOpcao = Trim(Request("slcOpcao"))


	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

' ----------Projetos---------------
    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_PROJETO_MEDICAO"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()


'----------------Frentes---------------------
    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_FRENTES"

		.Parameters.Refresh
		
		If Trim(strProj) <> "" Then
			.Parameters(1).Value = strProj
		Else
			.Parameters(1).Value = Null
		End If
		
    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()


'-------------Equipe------------------
	Set cmdResultado = Server.CreateObject("ADODB.Command")
	        
	With cmdResultado
	    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 600
	    .CommandText = "SP_LISTAR_EQUIPES_PROJ"

		.Parameters.Refresh

		If Trim(strProj) <> "" Then
			.Parameters(1).Value = strProj
		Else
			.Parameters(1).Value = Null
		End If
		
		If trim(strFrente) <> "" Then
			.Parameters(2).Value = strFrente
		Else
			.Parameters(2).Value = Null
		End If

	End With

	set rs2 = Server.CreateObject("ADODB.RecordSet")

	set rs2 = cmdResultado.Execute()

'------------Responsavel-------------------------		
	Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
	With cmdResultado
	    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 600
	    .CommandText = "SP_LISTAR_RESPONSAVEL"

		.Parameters.Refresh

		If Trim(strProj) <> "" Then
			.Parameters(1).Value = strProj
		Else
			.Parameters(1).Value = Null
		End If
		
		If trim(strFrente) <> "" Then
			.Parameters(2).Value = strFrente
		Else
			.Parameters(2).Value = Null
		End If

		If trim(strEquipe) <> "" Then
			.Parameters(3).Value = strEquipe
		Else
			.Parameters(3).Value = Null
		End If

	End With

	set rs3 = Server.CreateObject("ADODB.RecordSet")

	set rs3 = cmdResultado.Execute()


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
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcProjeto" style="WIDTH:300px;" onchange="Selfrente();">
						        <option value="" <%If strProj = "" Then%>Selected<%End If%>>--------------------------- Todos --------------------------</option>

								<%Do While Not rs.EOF%>
									<OPTION value="<%=rs("PROJ_ID")%>" <%If strProj = Trim(rs("PROJ_ID")) Then%>Selected<%End If%>><%=rs("PROJ_NAME")%></OPTION>
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
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Frente:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcFrente" style="WIDTH:300px;" onchange="Selfrente();">
						        <option value="" <%If strFrente = "" Then%>Selected<%End If%>>--------------------------- Todos --------------------------</option>

								<%Do While Not rs1.EOF%>
									<OPTION value="<%=rs1("Frente")%>" <%If strFrente = rs1("Frente") Then%>Selected<%End If%>><%=rs1("Frente")%></OPTION>
									<%rs1.MoveNext%>
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
						    <select size="1" name="slcEquipe" style="WIDTH:300px;" onchange="Selfrente();">
						        <option value="" <%If strEquipe = "" Then%>Selected<%End If%>>--------------------------- Todos --------------------------</option>
									<%Do While Not rs2.EOF%>
										<OPTION value="<%=rs2("Equipe")%>" <%If strEquipe = rs2("Equipe") Then%>Selected<%End If%>><%=rs2("Equipe")%></OPTION>
										<%rs2.MoveNext%>
									<%Loop%>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcResponsavel" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
									<%Do While Not rs3.EOF%>
										<OPTION value="<%=rs3("TEXT_VALUE")%>" <%If strResponsavel = rs3("TEXT_VALUE") Then%>Selected<%End If%>><%=rs3("TEXT_VALUE")%></OPTION>
										<%rs3.MoveNext%>
									<%Loop%>
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
						        <option value="" <%If Trim(strOpcao) = "" Then%>Selected<%End If%>>--------------------------- Todos --------------------------</option>
						        <option value="1" <%If Trim(strOpcao) = "1" Then%>Selected<%End If%>>HOJE</option>
						        <option value="2" <%If Trim(strOpcao) = "2" Then%>Selected<%End If%>> + 7 DIAS</option>
						        <option value="3" <%If Trim(strOpcao) = "3" Then%>Selected<%End If%>> + 14 DIAS</option>
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
	document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Detalhe.asp";
	document.frmRelatorio_Tarefas_Filtro.submit();
}

function ConfirmarExcel()
{
		document.frmRelatorio_Tarefas_Filtro.action = "Light_Relatorio_Tarefas_Detalhe_Excel.asp";
		document.frmRelatorio_Tarefas_Filtro.submit();
}

function Redefinir()
{

	document.frmRelatorio_Tarefas_Filtro.slcProjeto.value = "";
	document.frmRelatorio_Tarefas_Filtro.slcFrente.value = "";
	document.frmRelatorio_Tarefas_Filtro.slcEquipe.value = "";
	document.frmRelatorio_Tarefas_Filtro.slcResponsavel.value = "";
	document.frmRelatorio_Tarefas_Filtro.slcOpcao.value = "";

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
