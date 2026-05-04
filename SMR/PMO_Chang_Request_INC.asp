<!--#include file="./head.asp"-->
<!--#include file="./funcoes/Funcoes.inc"-->

<%

Dim strUsuario
Dim strProjeto
Dim strUID
Dim strDtInicio
Dim strDtFim
Dim strPerc
Dim strPredcessora
Dim strSucessora
Dim strObs
Dim strOperacao

Dim cmdResultado
Dim Rs

	strOperacao = Request("hidOperacao")
	strUsuario = Request("strUsuario")
	strProjeto	= Request("strProjeto")
	strUID	= Request("strUID")
	strWBS 	= Request("strWBS")
	strNome = Request("strNome")
	strDtInicio = Request("strDtInicio")
	strDtFim = Request("strDtFim")
	strPerc = Request("strPerc")
	strPredcessora = Request("strPredcessora")
	strSucessora = Request("strSucessora")
	strObs = Request("strObs")

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	If trim(strOperacao) = "A" Then

				
		Set cmdResultado = Server.CreateObject("ADODB.Command")
				    
		With cmdResultado
		    
			.ActiveConnection = conConexao
			.CommandType = 4
			.CommandTimeout = 600
			.CommandText = "SP_INCLUIR_CHANG_REQUESTS_INC"
					    
			.Parameters.Refresh
			.Parameters(1).Value = trim(strUsuario)
			.Parameters(2).Value = strProjeto
			.Parameters(3).Value = strUID
			.Parameters(4).Value = trim(strNome)
			.Parameters(5).Value = strDtInicio
			.Parameters(6).Value = strDtFim
			.Parameters(7).Value = strPerc
			
			if Trim(strPredcessora) <> "" then
				.Parameters(8).Value = strPredcessora
			End If

			if Trim(strSucessora) <> "" then
				.Parameters(9).Value = strSucessora
			End If

			if Trim(strObs) <> "" then
				.Parameters(10).Value = trim(strObs)
			End If
			
		End With

		cmdResultado.Execute()

%>
		<SCRIPT language=JavaScript>
			this.close();
		</SCRIPT>

<%
	End If

	Set cmdResultado = Server.CreateObject("ADODB.Command")
				    
	With cmdResultado
		    
		.ActiveConnection = conConexao
		.CommandType = 4
		.CommandTimeout = 600
		.CommandText = "SP_LISTAR_CHANG_REQUEST_SUMARIA"
					    
		.Parameters.Refresh
		.Parameters(1).Value = trim(strProjeto)
		.Parameters(2).Value = trim(strWBS)
			
	End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()


%>

<html>
<head>
<title>Projeto Sinergia </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<FORM name="frmChangRequest_PMO" id="frmChangRequest_PMO" action="PMO_Chang_Request.asp" method="post">
<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>

	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
		<tr>
			<td height="270" valign="top">
				<table width="780" border="0" cellspacing="0" cellpadding="0">

					<tr>
						<td width="65">&nbsp;</td>
						<td width="715"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></td>
					</tr>
					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15" valign="top">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>Sumaria pertencente:</strong>
							</font>
						</td>
					</tr>
					<tr>
						<td width="65"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715"> 
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong><%=rs("TASK_NAME")%></strong>
							</font>
						</td>
					</tr>

					<tr>
						<td width="65">&nbsp;</td>
						<td width="715"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></td>
					</tr>
					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15" valign="top">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>Abaixo da tarefa: UID - <%=strUID%></strong>
							</font>
						</td>
					</tr>

					<tr>
						<td width="65">&nbsp;</td>
						<td width="715"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></td>
					</tr>
					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15" valign="top">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>Nome da tarefa Adicionada</strong>
							</font>
						</td>
					</tr>
					<tr>
						<td width="65"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715"> 
							<input type="text" name="txtNome"  size=55 maxlength=254>
						</td>
					</tr>


					<tr>
						<td width="65">&nbsp;</td>
						<td width="715"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></td>
					</tr>

					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15" valign="top">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>
									Nova data de in&iacute;cio
									<img src="img/F_01.gif" width="100" height="1">
									Nova data fim
									<img src="img/F_01.gif" width="125" height="1">
									Avan&ccedil;o real 
								</strong>
							</font>
						</td>
					</tr>
					<tr>
						<td width="65"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715">
							<input type="text" name="txtDtInicio" value="">

							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>
									<img src="img/F_01.gif" width="58" height="1"> 
									<input type="text" name="txtDtFim" value="">

									<img src="img/F_01.gif" width="58" height="1"> 
									<input name="txtPercComp" type="text" size="10" value="" onKeyUp="CaracteresValidos('1234567890',this.value);">

									<img src="img/F_01.gif" width="58" height="1"> 
								</strong>
							</font>
						</td>
					</tr>

					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
					</tr>
					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>
									Unique Id predecessora  da tarefa Adicionada
									<img src="img/F_01.gif" width="35" height="1">
									Unique Id sucessora  da tarefa Adicionada
								</strong>
							</font>
						</td>
					</tr>
					<tr>
						<td width="65">&nbsp;</td>
						<td width="715">
							<input type="text" name="txtPredcessora"  size=37>
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>
									<img src="img/F_01.gif" width="45" height="1"> 
									<input type="text" name="txtSucessora" size=37>
								</strong>
							</font>
						</td>
					</tr>

					<tr>
						<td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
						<td width="715" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
					</tr>
					<tr>
						<td width="65">&nbsp;</td>
						<td width="715">
							<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
								<strong>OBS: (Motivo de inclusão da nova tarefa) - Preenchimento obrigatório</strong>
							</font>
						</td>
					</tr>
					<tr>
						<td width="65">&nbsp;</td>
						<td width="715"><textarea name="txtObs" cols="65"></textarea></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td height="15" background="img/F000004.gif"><img src="img/F_01.gif" width="1" height="1"></td>
		</tr>
	</table>

	<input type="hidden" id="hidProj" name="hidProj" value="">
	<input type="hidden" id="hidUid" name="hidUid" value="">

	<table width="650" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td><div align="right"><br>
			<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a><img src="img/F_01.gif" width="40" height="1">
<!--			<a href="javascript:Confirmar();" target="blank"><img src="img/F000005.gif" width="75" height="18"></a><img src="img/F_01.gif" width="40" height="1"><br>-->
			</div></td>
		</tr>
	</table>

<p>&nbsp;</p>

<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
<input type="hidden" id="strUsuario" name="strUsuario" value="<%=strUsuario%>">
<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">
<input type="hidden" id="strWBS" name="strWBS" value="<%=strWBS%>">
<input type="hidden" id="strNome" name="strNome" value="">
<input type="hidden" id="strDtInicio" name="strDtInicio" value="">
<input type="hidden" id="strDtFim" name="strDtFim" value="">
<input type="hidden" id="strPerc" name="strPerc" value="">
<input type="hidden" id="strPredcessora" name="strPredcessora" value="">
<input type="hidden" id="strSucessora" name="strSucessora" value="">
<input type="hidden" id="strObs" name="strObs" value="">


<SCRIPT language=JavaScript>

function Confirmar()
{

	document.frmChangRequest_PMO.style.cursor = "wait";
	document.frmChangRequest_PMO.hidOperacao.value = "";
	
	if (ValidaNome())
	{
		if (ValidaDataInicio())
		{
			if (ValidaDataFim())
			{
				if (ValidaComplete())
				{

					document.frmChangRequest_PMO.hidOperacao.value = 'A';
					document.frmChangRequest_PMO.style.cursor = "";

					document.frmChangRequest_PMO.strNome.value = document.frmChangRequest_PMO.txtNome.value;
					document.frmChangRequest_PMO.strDtInicio.value = document.frmChangRequest_PMO.txtDtInicio.value;
					document.frmChangRequest_PMO.strDtFim.value = document.frmChangRequest_PMO.txtDtFim.value;
					document.frmChangRequest_PMO.strPerc.value = document.frmChangRequest_PMO.txtPercComp.value;
					document.frmChangRequest_PMO.strPredcessora.value = document.frmChangRequest_PMO.txtPredcessora.value;
					document.frmChangRequest_PMO.strSucessora.value = document.frmChangRequest_PMO.txtSucessora.value;
					document.frmChangRequest_PMO.strObs.value = document.frmChangRequest_PMO.txtObs.value;

					document.frmChangRequest_PMO.action = "PMO_Chang_Request_INC.asp";
					document.frmChangRequest_PMO.submit();
				
				}
			}
		}
	}

	document.frmChangRequest_PMO.style.cursor = "";

}
//***************************************************************************************************


function ValidaNome()
{
	var conteudo = document.frmChangRequest_PMO.txtNome.value
	
	if (jTrim(conteudo) == "")
	{
		alert("O conteúdo do campo nome é obrigatório.");
		document.frmChangRequest_PMO.txtNome.focus();
	    return false;
	}
	else
	{
	    return true;
	}

}


function ValidaDataInicio()
{

	var conteudo = document.frmChangRequest_PMO.txtDtInicio.value

	var titulo = "Nova data de inicio"
	var intTamanho = document.frmChangRequest_PMO.txtDtInicio.value.length;
	
	var msg_erro;
	var bol;
	var strAux;	
	var strDia = conteudo.substring(0, 2)
	var strMes = conteudo.substring(3, 5)
	var strAno = conteudo.substring(6, 10)

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
	else
	{
		msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
		bol = true;
	}

	if (bol){
		alert(msg_erro);
		document.frmChangRequest_PMO.txtDtInicio.focus();
	    return false;
	}	
	else
	{
	    return true;
	}

}


function ValidaDataFim()
{

	var conteudo = document.frmChangRequest_PMO.txtDtFim.value;
	var conteudoAux = document.frmChangRequest_PMO.txtDtInicio.value;

	var DataFim = parseInt(conteudo.substring(6, 10) + conteudo.substring(3, 5) + conteudo.substring(0, 2))
	var DataInicio = parseInt(conteudoAux.substring(6, 10) + conteudoAux.substring(3, 5) + conteudoAux.substring(0, 2))

	var titulo = "Nova data de fim";
	var intTamanho = document.frmChangRequest_PMO.txtDtFim.value.length;
	
	var msg_erro;
	var bol;
	var strAux;	
	var strDia = conteudo.substring(0, 2);
	var strMes = conteudo.substring(3, 5);
	var strAno = conteudo.substring(6, 10);

	if (conteudo != "")
	{
		if (intTamanho == 10)
		{
			if (strMes < 1 || strMes > 12) 
			{
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strDia < 1 || strDia > 31) 
			{
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno<1) 
			{
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno < 1990) 
			{
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}
			if (strAno > 2100) 
			{
				msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
				bol = true;
			}

			if (strMes == 4 || strMes == 6 || strMes == 9 || strMes == 11) 
			{
				if (strDia == 31) 
				{
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}		
			}
			if (strMes == 2)
			{
				strAux = parseInt(strAno/4);
				if (isNaN(strAux)) 
				{
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
				if (strDia > 29) 
				{
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
				if (strDia == 29 && ((strAno/4) != parseInt(strAno/4))) 
				{
					msg_erro = "O conteúdo do campo "+titulo+" esta invalido.";
					bol = true;
				}
			}
			
			if (DataFim < DataInicio)
			{
				msg_erro = "O conteúdo do campo "+titulo+" menor que data inicio.";
				bol = true;
			}
		}
		else
		{
			msg_erro = "O conteúdo do campo "+titulo+" esta fora do formato. (DD/MM/AAAA).";
			bol = true;
		}
	}
	else
	{
		msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
		bol = true;		
	}

	if (bol)
	{
		alert(msg_erro);
		document.frmChangRequest_PMO.txtDtFim.focus();

	    return false;
	}	
	else
	{
	    return true;
	}

}

function ValidaComplete()
{

	var conteudo = document.frmChangRequest_PMO.txtPercComp.value
	
	if (jTrim(conteudo) == "")
	{
		alert("O conteúdo do campo Avanço real obrigatório.");
		document.frmChangRequest_PMO.txtPercComp.focus();
		return false;
	}
	else
	{
		if (parseInt(conteudo) < 0)
		{
			alert("O conteúdo do campo Avanço real invalido. deve ser maior ou igual a 0.");
			document.frmChangRequest_PMO.txtPercComp.focus();
			return false;
		}

		if (parseInt(conteudo) > 100)
		{
			alert("O conteúdo do campo Avanço real invalido. deve ser menor ou igual a 100.");
			document.frmChangRequest_PMO.txtPercComp.focus();
			return false;
		}

	}

    return true;

}

</SCRIPT>

</form>
</body>
</html>

