<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

Dim strProj
Dim strFrente
Dim strDataIni
Dim strDataFim

Dim rs
Dim rs1
Dim rs2
Dim rs3

Dim cmdResultado

	strProj   = Trim(Request("slcProjeto"))
	strFrente = Trim(Request("slcFrente"))
	strDataIni = Trim(Request("txtDataIni"))
	strDataFim = Trim(Request("txtDataFim"))


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
<form method="post" name="frmRelatorio_Comentario_Filtro" id="frmRelatorio_Comentario_Filtro">

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
			    	<td bgcolor=#6699cc align="right" >
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc >
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
			    	<td bgcolor=#6699cc align="right" >
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Frente:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc >
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
  	
			</table>
			
			<BR>
			
			<table>

			  	<tr >
			    	<td bgcolor=#6699cc align=center  colspan=3>
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Periodo</font>
						</b>
					</td>
				</tr>
			  	<tr >

			    	<td align="right" bgcolor=#6699cc >
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
							<input id="txtDataIni" name="txtDataIni" value="<%=strDataIni%>" maxLength=10 size=10 onblur="ValidaDatas(this,'I');">
						</font>
					</td>
			    	<td bgcolor=#6699cc align=center >
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;Ate&nbsp;</font>
						</b>
					</td>

			    	<td align="right" bgcolor=#6699cc >
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
	  						<input id="txtDataFim" name="txtDataFim" value="<%=strDataFim%>" maxLength=10 size=10 onblur="ValidaDatas(this,'F');">
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
	document.frmRelatorio_Comentario_Filtro.action = "Light_Relatorio_Comentario_Filtro.asp";
	document.frmRelatorio_Comentario_Filtro.submit();

}


function Confirmar()
{
	document.frmRelatorio_Comentario_Filtro.action = "Light_Relatorio_Comentario_Detalhe.asp";
	document.frmRelatorio_Comentario_Filtro.submit();
}

function ConfirmarExcel()
{
		document.frmRelatorio_Comentario_Filtro.action = "Light_Relatorio_Comentario_Detalhe_Excel.asp";
		document.frmRelatorio_Comentario_Filtro.submit();
}

function Redefinir()
{

	document.frmRelatorio_Comentario_Filtro.slcProjeto.value = "";
	document.frmRelatorio_Comentario_Filtro.slcFrente.value = "";
	document.frmRelatorio_Comentario_Filtro.txtDataIni.value = "";
	document.frmRelatorio_Comentario_Filtro.txtDataFim.value = "";

	document.frmRelatorio_Comentario_Filtro.action = "Light_Relatorio_Comentario_Filtro.asp";
	document.frmRelatorio_Comentario_Filtro.submit();

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


function ValidaDatas(conteudoAux, Tipo)
{

	var conteudo = conteudoAux.value
	var intTamanho = conteudoAux.value.length;
	var msg_erro;
	var bol;
	var strAux;	
	var strDia = conteudo.substring(0, 2)
	var strMes = conteudo.substring(3, 5)
	var strAno = conteudo.substring(6, 10)

	if (Tipo == "I")
	{
		var titulo = "Data Inicio"
	}
	else
	{
		var titulo = "Data Fim"	
	}	

	if (conteudo != "")
	{
		if (intTamanho == 10)
		{
			if (strMes < 1 || strMes > 12) {
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strDia < 1 || strDia > 31) {
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno<1) {
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno < 1990) {
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno > 2100) {
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}

			if (strMes == 4 || strMes == 6 || strMes == 9 || strMes == 11) 
			{
				if (strDia == 31) {
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}		
			}
			if (strMes == 2)
			{
				strAux = parseInt(strAno/4);
				if (isNaN(strAux)) {
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
				if (strDia > 29) {
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
				if (strDia == 29 && ((strAno/4) != parseInt(strAno/4))) {
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
			}
		}
		else
		{
			msg_erro = "O conteúdo do campo "+titulo+" esta fora do formato. (DD/MM/AAAA).";
			bol = true;
		}
	}

	if (bol){
		alert(msg_erro);

		if (Tipo == "I")
		{
			document.frmRelatorio_Comentario_Filtro.txtDataIni.value = "";
			document.frmRelatorio_Comentario_Filtro.txtDataIni.focus();
		}
		else
		{
			document.frmRelatorio_Comentario_Filtro.txtDataFim.value = "";
			document.frmRelatorio_Comentario_Filtro.txtDataFim.focus();
		}				
	    return false;
	}
}

</SCRIPT>

</body>
</html>
