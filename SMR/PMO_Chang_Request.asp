<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->
<%
Dim I

Dim cmdResultado
Dim vetDados
Dim vetDados_Disc
Dim strProj
Dim strUID
Dim strAcao
Dim strDtInicioNov
Dim strDtFinalNov
Dim strPercCompNov
Dim strUIDPred
Dim strUIDSuc
Dim strRazao
Dim strObs

Dim strOperacao
Dim strSql
Dim Cont

	If trim(session("Usuario")) = "" Then
		response.Redirect("./LOGIN.ASP?hidOrigem=./PMO_Chang_Request_Selecao.asp")
	End if

strOperacao = Request("hidOperacao")

If trim(strOperacao) = "A" Then

	vetDados = split(Request("hidDados"),";")

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	For I = 0 to ubound(vetDados)

		vetDados_Disc = split(vetDados(I),"|")

		strProj = vetDados_Disc(0)
		strUID = vetDados_Disc(1)
		strAcao = vetDados_Disc(2)
		strDtInicioNov = vetDados_Disc(3)
		strDtFinalNov = vetDados_Disc(4)
		strPercCompNov = vetDados_Disc(5)
		strRazao = vetDados_Disc(6)
		strObs = vetDados_Disc(7)
		
		Set cmdResultado = Server.CreateObject("ADODB.Command")
		    
		With cmdResultado
    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 600
		    .CommandText = "SP_INCLUIR_CHANG_REQUESTS_ALT"
		    
		    .Parameters.Refresh
			.Parameters(1).Value = trim(session("Usuario"))
			.Parameters(2).Value = strProj
			.Parameters(3).Value = strUID
			.Parameters(4).Value = trim(strAcao)

			If trim(strDtInicioNov) <> "" Then
				.Parameters(5).Value = strDtInicioNov
			End If
			
			If trim(strDtFinalNov) <> "" Then
				.Parameters(6).Value = strDtFinalNov
			End If
			
			If trim(strPercCompNov) <> "" Then
				.Parameters(7).Value = strPercCompNov
			End If

			If trim(strUIDPred) <> "" Then
				.Parameters(8).Value = strUIDPred
			End If

			If trim(strUIDSuc) <> "" Then
				.Parameters(9).Value = strUIDSuc
			End If
			
			If trim(strRazao) <> "" Then
				.Parameters(10).Value = trim(strRazao)
			End If

			If trim(strObs) <> "" Then
				.Parameters(11).Value = trim(strObs)
			End If

		End With

		cmdResultado.Execute()

	Next

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)

	response.Redirect("./PMO_Chang_Request_Detalhe.asp")

Else

	vetDados = split(Request("hidDados"),";")

End If

%>

<html>
<head>
<title>Projeto Sinergia </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<FORM name="frmChangRequest_PMO" id="frmChangRequest_PMO" action="PMO_Chang_Request.asp" method="post">
<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>

<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">&nbsp;&nbsp;Alteração/Exclusão</font></b>

<%Cont = 0%>
<%For I = 0 To ubound(vetDados)%>
	<%
	vetDados_Disc = split(vetDados(I),"|")
	%>
	<p>&nbsp;</p>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="1" background="img/F_0.gif"><img src="img/F_01.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<br>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr> 
	    <td height="23" valign="top" background="img/F000001.gif"> 
	      <table width="780" border="0" cellpadding="0" cellspacing="0" background="img/F000002.gif">
	        <tr> 
	          <td width="167" height="19"> 
	            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Unique 
	              Id</strong></font></div></td>
	          <td width="305" height="19"> 
	            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome 
	              da atividade</strong></font></div></td>
	          <td width="121" height="19"> 
	            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
	              in&iacute;cio</strong></font></div></td>
	          <td width="104" height="19"> 
	            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
	              fim</strong></font></div></td>
				<td width="83" height="19"> 
	            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>%</strong></font></div></td>
	        </tr>
	      </table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="4"><img src="img/F_01.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" height="33" border="0" cellpadding="0" cellspacing="0">
	  <tr>
	    <td valign="middle" background="img/F000003.gif"><table width="780" height="33" border="0" cellpadding="0" cellspacing="0">
	        <tr> 
	          <td width="167" height="19"> <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=vetDados_Disc(1)%></font></div></td>
	          <td width="305" height="19"> <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=vetDados_Disc(2)%></font></div></td>
	          <td width="121" height="19"> <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=vetDados_Disc(3)%></strong></font></div></td>
	          <td width="104" height="19"> 
	            <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=vetDados_Disc(4)%></strong></font></div></td>
				<td width="83" height="19"> 
	            <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=vetDados_Disc(5)%></strong></font></div></td>
	        </tr>
	      </table></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr> 
	    <td height="4"><img src="img/F_01.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
	  <tr>
	    <td height="270" valign="top">
	<table width="780" border="0" cellspacing="0" cellpadding="0">
	        <tr>
	          <td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
	          <td width="705" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
	        </tr>
			<tr>
	          <td width="65" height="50">&nbsp;</td>
	          <td width="715" height="50"> 
	            <select name="slcAcao">
	                <option value=""></option>
	                <option value="Nova data inicio">Nova data inicio</option>
	                <option value="Nova data fim">Nova data fim</option>
	                <option value="Mudança de %">Mudança de %</option>
	                <option value="Eliminar Atividade">Eliminar Atividade</option>
	                <option value="Outros">Outros</option>
	              </select>
	            </td>
	        </tr>
	        <tr>
	          <td width="65" height="15"><img src="img/F_01.gif" width="1" height="1"></td>
	          <td width="715" height="15" valign="top">
				<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>
						Nova data in&iacute;cio
						<img src="img/F_01.gif" width="118" height="1">
						Nova data fim
						<img src="img/F_01.gif" width="125" height="1">
						Fim limite
<!--						<img src="img/F_01.gif" width="150" height="1">
						Avan&ccedil;o real -->
					</strong>
				</font>
			  </td>
	        </tr>
	        <tr>
				<td width="65"><img src="img/F_01.gif" width="1" height="1"></td>
				<td width="715">
					<input type="text" name="txtDtInicio" value="<%=vetDados_Disc(3)%>" onblur="ValidaDataInicio(<%=Cont%>);">
					<input type="hidden" id="hidDtInicio" name="hidDtInicio" value="<%=vetDados_Disc(3)%>">

					<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
							<img src="img/F_01.gif" width="58" height="1"> 
							<input type="text" name="txtDtFim" value="<%=vetDados_Disc(4)%>" onblur="ValidaDataFim(<%=Cont%>);">
							<input type="hidden" id="hidDtFim" name="hidDtFim" value="<%=vetDados_Disc(4)%>">

							<img src="img/F_01.gif" width="58" height="1"> 
							<input type="text" name="txtDtLimite"  disabled value="<%=vetDados_Disc(6)%>">
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
						Avan&ccedil;o real
<!--						Unique Id predecessora-->
						<img src="img/F_01.gif" width="140" height="1">
						Raz&otilde;es da mudan&ccedil;a da atividade (Obrigatório)
<!--						Unique Id sucessora
						<img src="img/F_01.gif" width="94" height="1">
						Raz&otilde;es da mudan&ccedil;a-->
			   		</strong>
			  	</font>
			  </td>
	        </tr>
			<tr>
	          <td width="65">&nbsp;</td>
	          <td width="715">
				<input name="txtPercComp" type="text" size="10" value="<%=vetDados_Disc(5)%>" onKeyUp="CaracteresValidos('1234567890',this.value);" onblur="ValidaComplete(<%=Cont%>);">
				<input type="hidden" id="hidPerc" name="hidPerc" value="<%=vetDados_Disc(5)%>">

<!--				<input type="text" name="txtUIDPred" onKeyUp="CaracteresValidos('1234567890',this.value);">-->
	            <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>
						<img src="img/F_01.gif" width="118" height="1"> 
						<input type="text" name="txtRazao"  size=55 maxlength=254>

<!--						<input type="text" name="txtUIDSuc" onKeyUp="CaracteresValidos('1234567890',this.value);">
						<img src="img/F_01.gif" width="58" height="1"> -->
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
	          <td width="715"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>OBS:&nbsp;&nbsp;(Maiores detalhes sobre a modificação)</strong></font></td>
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

	<input type="hidden" id="hidProj" name="hidProj" value="<%=vetDados_Disc(0)%>">
	<input type="hidden" id="hidUid" name="hidUid" value="<%=vetDados_Disc(1)%>">

	<%Cont = Cont + 1%>

<%Next%>




<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="right"><br>
			<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a><img src="img/F_01.gif" width="40" height="1">
<!--        <a href="javascript:Confirmar();"><img src="img/F000005.gif" width="75" height="18"></a><br>--><img src="img/F_01.gif" width="40" height="1">
      </div></td>
  </tr>
</table>

<p>&nbsp;</p>

<input type="hidden" id="hidCont" name="hidCont" value="<%=Cont%>">
<input type="hidden" id="hidDados" name="hidDados" value="">
<input type="hidden" id="hidOperacao" name="hidOperacao" value="">

<SCRIPT language=JavaScript>

function Confirmar()
{
	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)
	var strAux = ""
	var strErro = ""
	
	document.frmChangRequest_PMO.style.cursor = "wait";
	document.frmChangRequest_PMO.hidDados.value = "";
	document.frmChangRequest_PMO.hidOperacao.value = "";
	
	for(var i = 0; i <= intCont;i++)
	{

		if (ValidaAcao(i))
		{
			if (ValidaRazao(i))
			{
				if (ValidaObs(i))
				{
					if (intCont == 0)
					{
						if (strAux == "")
						{
							strAux = document.frmChangRequest_PMO.hidProj.value + "|" + 
							document.frmChangRequest_PMO.hidUid.value + "|" + 
							document.frmChangRequest_PMO.slcAcao.value + "|" + 
							document.frmChangRequest_PMO.txtDtInicio.value + "|" + 
							document.frmChangRequest_PMO.txtDtFim.value + "|" + 
							document.frmChangRequest_PMO.txtPercComp.value + "|" + 
							document.frmChangRequest_PMO.txtRazao.value + "|" + 
							document.frmChangRequest_PMO.txtObs.value;
						}	
					}
					else
					{
						if (strAux == "")
						{
							strAux = document.frmChangRequest_PMO.hidProj(i).value + "|" + 
							document.frmChangRequest_PMO.hidUid(i).value + "|" + 
							document.frmChangRequest_PMO.slcAcao(i).value + "|" + 
							document.frmChangRequest_PMO.txtDtInicio(i).value + "|" + 
							document.frmChangRequest_PMO.txtDtFim(i).value + "|" + 
							document.frmChangRequest_PMO.txtPercComp(i).value + "|" + 
							document.frmChangRequest_PMO.txtRazao(i).value + "|" + 
							document.frmChangRequest_PMO.txtObs(i).value;
						}	
						else
						{
							strAux = strAux + ";" + document.frmChangRequest_PMO.hidProj(i).value + "|" + 
							document.frmChangRequest_PMO.hidUid(i).value + "|" + 
							document.frmChangRequest_PMO.slcAcao(i).value + "|" + 
							document.frmChangRequest_PMO.txtDtInicio(i).value + "|" + 
							document.frmChangRequest_PMO.txtDtFim(i).value + "|" + 
							document.frmChangRequest_PMO.txtPercComp(i).value + "|" + 
							document.frmChangRequest_PMO.txtRazao(i).value + "|" + 
							document.frmChangRequest_PMO.txtObs(i).value;

						}
					}
									
				}
				else
				{
					if (intCont == 0)
					{
						alert("O conteúdo do campo Obs obrigatório.");
						document.frmChangRequest_PMO.txtObs.focus();
						document.frmChangRequest_PMO.style.cursor = "";
						strErro = "E";
						i = intCont;
					}
					else
					{
						alert("O conteúdo do campo Obs obrigatório.");
						document.frmChangRequest_PMO.txtObs(i).focus();
						document.frmChangRequest_PMO.style.cursor = "";
						strErro = "E";
						i = intCont;					
					}
				}
			}
			else
			{
				if (intCont == 0)
				{
					alert("O conteúdo do campo Razões da mudança obrigatório.");
					document.frmChangRequest_PMO.txtRazao.focus();
					document.frmChangRequest_PMO.style.cursor = "";
					strErro = "E";
					i = intCont;
				}
				else
				{
					alert("O conteúdo do campo Razões da mudança obrigatório.");
					document.frmChangRequest_PMO.txtRazao(i).focus();
					document.frmChangRequest_PMO.style.cursor = "";
					strErro = "E";
					i = intCont;
				}
			}
		}
		else
		{
			if (intCont == 0)
			{
				alert("O conteúdo do campo Ação obrigatório.");
				document.frmChangRequest_PMO.slcAcao.focus();
				document.frmChangRequest_PMO.style.cursor = "";
				strErro = "E";
				i = intCont;
			}
			else
			{
				alert("O conteúdo do campo Ação obrigatório.");
				document.frmChangRequest_PMO.slcAcao(i).focus();
				document.frmChangRequest_PMO.style.cursor = "";
				strErro = "E";
				i = intCont;
			}
		}
	}

	if (strErro == "")
	{
		if (strAux != "")
		{
			document.frmChangRequest_PMO.hidDados.value = strAux;
			document.frmChangRequest_PMO.hidOperacao.value = 'A';
			document.frmChangRequest_PMO.style.cursor = "";
			document.frmChangRequest_PMO.action = "PMO_Chang_Request.asp";
			document.frmChangRequest_PMO.submit();
		}
		else
		{
			alert("Nenhum registro foi alterado");
			document.frmChangRequest_PMO.style.cursor = "";
		}

	}
}
//***************************************************************************************************

function ValidaDataInicio(num)
{

	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)
	var titulo = "Nova data inicio"

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.txtDtInicio.value
		var intTamanho = document.frmChangRequest_PMO.txtDtInicio.value.length;
	
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.txtDtInicio(num).value
		var intTamanho = document.frmChangRequest_PMO.txtDtInicio(num).value.length;
	}

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
		if (intCont == 0)
		{
			if (document.frmChangRequest_PMO.slcAcao.value == 'Nova data inicio')
			{
				msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
				bol = true;		
			}
		}
		else
		{
			if (document.frmChangRequest_PMO.slcAcao(num).value == 'Nova data inicio')
			{
				msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
				bol = true;		
			}
		}
	}

	if (bol)
	{
		if (intCont == 0)
		{
			alert(msg_erro);
			document.frmChangRequest_PMO.txtDtInicio.value = document.frmChangRequest_PMO.hidDtInicio.value;
			document.frmChangRequest_PMO.txtDtInicio.focus();

			return false;
		}
		else
		{
			alert(msg_erro);
			document.frmChangRequest_PMO.txtDtInicio(num).value = document.frmChangRequest_PMO.hidDtInicio(num).value;
			document.frmChangRequest_PMO.txtDtInicio(num).focus();

			return false;		
		}
	}
}

function ValidaDataFim(num)
{

	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)
	var titulo = "Nova data fim";

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.txtDtFim.value;
		var conteudoAux = document.frmChangRequest_PMO.txtDtInicio.value;
		var intTamanho = document.frmChangRequest_PMO.txtDtFim.value.length;
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.txtDtFim(num).value;
		var conteudoAux = document.frmChangRequest_PMO.txtDtInicio(num).value;
		var intTamanho = document.frmChangRequest_PMO.txtDtFim(num).value.length;	
	}

	var DataFim = parseInt(conteudo.substring(6, 10) + conteudo.substring(3, 5) + conteudo.substring(0, 2))
	var DataInicio = parseInt(conteudoAux.substring(6, 10) + conteudoAux.substring(3, 5) + conteudoAux.substring(0, 2))
	
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
		if (intCont == 0)
		{
			if (document.frmChangRequest_PMO.slcAcao.value == 'Nova data fim')
			{
				msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
				bol = true;		
			}
		}
		else
		{
			if (document.frmChangRequest_PMO.slcAcao(num).value == 'Nova data fim')
			{
				msg_erro = "O conteúdo do campo "+titulo+" obrigatório.";
				bol = true;		
			}
		}
	}

	if (bol)
	{
		if (intCont == 0)
		{
			alert(msg_erro);
			document.frmChangRequest_PMO.txtDtFim.value =  document.frmChangRequest_PMO.hidDtFim.value ;
			document.frmChangRequest_PMO.txtDtFim.focus();

			return false;
		}
		else
		{
			alert(msg_erro);
			document.frmChangRequest_PMO.txtDtFim(num).value =  document.frmChangRequest_PMO.hidDtFim(num).value ;
			document.frmChangRequest_PMO.txtDtFim(num).focus();

			return false;
		}
	}
}

function ValidaComplete(num)
{
	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.txtPercComp.value
	
		if (jTrim(conteudo) == "")
		{
			if (document.frmChangRequest_PMO.slcAcao.value == 'Mudança de %')
			{
				alert("O conteúdo do campo Avanço real obrigatório.");
				document.frmChangRequest_PMO.txtPercComp.value = document.frmChangRequest_PMO.hidPerc.value;
				document.frmChangRequest_PMO.txtPercComp.focus();
				return false;
			}
		}
		else
		{
			if (parseInt(conteudo) < 0)
			{
				alert("O conteúdo do campo Avanço real invalido. deve ser maior ou igual a 0.");
				document.frmChangRequest_PMO.txtPercComp.value = document.frmChangRequest_PMO.hidPerc.value;
				document.frmChangRequest_PMO.txtPercComp.focus();
				return false;
			}

			if (parseInt(conteudo) > 100)
			{
				alert("O conteúdo do campo Avanço real invalido. deve ser menor ou igual a 100.");
				document.frmChangRequest_PMO.txtPercComp.value = document.frmChangRequest_PMO.hidPerc.value;
				document.frmChangRequest_PMO.txtPercComp.focus();
				return false;
			}

		}
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.txtPercComp(num).value
	
		if (jTrim(conteudo) == "")
		{
			if (document.frmChangRequest_PMO.slcAcao(num).value == 'Mudança de %')
			{
				alert("O conteúdo do campo Avanço real obrigatório.");
				document.frmChangRequest_PMO.txtPercComp(num).value = document.frmChangRequest_PMO.hidPerc(num).value;
				document.frmChangRequest_PMO.txtPercComp(num).focus();
				return false;
			}
		}
		else
		{
			if (parseInt(conteudo) < 0)
			{
				alert("O conteúdo do campo Avanço real invalido. deve ser maior ou igual a 0.");
				document.frmChangRequest_PMO.txtPercComp(num).value = document.frmChangRequest_PMO.hidPerc(num).value;
				document.frmChangRequest_PMO.txtPercComp(num).focus();
				return false;
			}

			if (parseInt(conteudo) > 100)
			{
				alert("O conteúdo do campo Avanço real invalido. deve ser menor ou igual a 100.");
				document.frmChangRequest_PMO.txtPercComp(num).value = document.frmChangRequest_PMO.hidPerc(num).value;
				document.frmChangRequest_PMO.txtPercComp(num).focus();
				return false;
			}

		}
	}
}

function ValidaAcao(num)
{
	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.slcAcao.value
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.slcAcao(num).value
	}
	
	if (jTrim(conteudo) == "")
	{
		return (false);
	}
	else
	{
		return (true);
	}
}

function ValidaRazao(num)
{
	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.txtRazao.value
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.txtRazao(num).value
	}

	if (jTrim(conteudo) == "")
	{
		return (false);
	}
	else
	{
		return (true);
	}
}

function ValidaObs(num)
{
	var intCont	= (document.frmChangRequest_PMO.hidCont.value - 1)

	if (intCont == 0)
	{
		var conteudo = document.frmChangRequest_PMO.txtObs.value
	}
	else
	{
		var conteudo = document.frmChangRequest_PMO.txtObs(num).value
	}
	
	if ((jTrim(conteudo) == "") && (document.frmChangRequest_PMO.slcAcao(num).value == 'Outros'))
	{
		return (false);
	}
	else
	{
		return (true);
	}
}


</SCRIPT>

</form>
</body>
</html>

