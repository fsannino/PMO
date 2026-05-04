<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim cmdResultado
Dim strProjeto
Dim strUID
Dim strID
Dim strSql
Dim strOperacao
Dim strDtInicio
Dim strDtFim
Dim strInicio
Dim strFim
Dim strLogin
Dim strNome
Dim strSequencia
Dim strMotivo
Dim rs

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,7,4) + mid(strData,4,2) + mid(strData,1,2)
End Function

	strProjeto = Request("strProjeto")
	strUID	   = Request("strUID")
	strID	   = Request("strID")
'	strInicio  = Request("strDataIni")
'	strFim     = Request("strDataFim")
	strLogin   = Request("strLogin")

	strNome        = Request("txtNome")
	strDtInicio    = Request("txtDtInicio")
	strDtFim       = Request("txtDtFim")
	strSequencia   = Request("txtSequencia")
	strMotivo	   = Request("txtMotivo")

	strOperacao	= Request("hidOperacao")
	
	'Abrindo uma conexão com o BD
	set conConexao = CUT_AbrirConexaoBD()

	if strOperacao <> "" Then

		strSql = "SP_LISTAR_SEQUENCIA " & strProjeto & ", " & strUID & ", " & Trim(strSequencia)

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao
		
		If not rs.eof Then

	%>	<SCRIPT language=JavaScript>
			alert("Sequencia já existente.")
		</SCRIPT>

	<%
		Else

		    Set cmdResultado = Server.CreateObject("ADODB.Command")

			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 600
			    .CommandText = "SP_INCLUIR_NOVAS_TAREFAS"
			    
			    .Parameters.Refresh
				.Parameters(1).Value = strProjeto
				.Parameters(2).Value = strUID
				.Parameters(3).Value = strID
				.Parameters(4).Value = strNome
				.Parameters(5).Value = strDtInicio
				.Parameters(6).Value = strDtFim
				.Parameters(7).Value = strSequencia
				.Parameters(8).Value = strLogin
				.Parameters(9).Value = strMotivo

			End With

			cmdResultado.Execute()


		    Set cmdResultado = Server.CreateObject("ADODB.Command")

			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 600
			    .CommandText = "SP_INCLUIR_LOG_NOVAS_TAREFAS"
			    
			    .Parameters.Refresh
				.Parameters(1).Value = strProjeto
				.Parameters(2).Value = strUID
				.Parameters(3).Value = strID
				.Parameters(4).Value = strNome
				.Parameters(5).Value = strDtInicio
				.Parameters(6).Value = strDtFim
				.Parameters(7).Value = strSequencia
				.Parameters(8).Value = strLogin
				.Parameters(9).Value = strMotivo

			End With

			cmdResultado.Execute()

		End if

	End if

	strSql = "SP_LISTAR_NOVAS_TAREFAS " & strProjeto & ", " & strUID & " "

	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.CursorLocation = 3

	rs.OPEN STRSQL, conConexao, 3, 1, 1
	
%>
<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmIncluirTarefa" ID="frmIncluirTarefa" action="CUT_Incluir_Tarefas.asp" method="post">

	<center>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>

	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
	    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Dt.Inicio</font></b></td>
	    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Dt.Fim</font></b></td>
	    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Seq</font></b></td>

	  </tr>

		<%Do While Not rs.EOF%>
			
			<tr height="17" style="height:12.75pt">
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width=300px>
			  	<font face="Arial" size="1">
			  		<%=rs("NOME")%>&nbsp;
			  	</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="85px"  align=center>
			  	<font face="Arial" size="1">
					<%=FormatarDataMon(rs("DataIni"))%>&nbsp;
			  	</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="85px" align=center>
			  	<font face="Arial" size="1">
					<%=FormatarDataMon(rs("DataFim"))%>&nbsp;
			  	</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="40px" align=right>
			  	<font face="Arial" size="1">
 			  		<%=rs("SEQUENCIA")%>&nbsp;
			  	</font>
			  </td>

			</tr>
		<%rs.MoveNext%>
		<%Loop%>
	</table>
	<BR>

	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtNome" name="txtNome" value="" maxLength=255 size=65>
			</font>
		</td>
		</tr>
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data Inicio</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtDtInicio" name="txtDtInicio" value="" maxLength=10 size=10 onblur="ValidaDatas(this,'I');">
			</font>
		</td>
		</tr>
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data Fim</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtDtFim" name="txtDtFim" value="" maxLength=10 size=10 onblur="ValidaDatas(this,'F');">
			</font>
		</td>
		</tr>

	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Sequencia</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="100px" align=left>
			<font face="Arial" size="1">
				<input id="txtSequencia" name="txtSequencia" value="" maxLength=5 size=10>
			</font>
		</td>
		</tr>

	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Motivo</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtMotivo" name="txtMotivo" value="" maxLength=255 size=65>
			</font>
		</td>
		</tr>

	</table>

	<br>

	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<TD></TD>
			<td align=right>
				<a href="javascript:Confirmar();"><img src="img/000049.gif" align="absmiddle" border="0">
				<a href="javascript:this.close();" ><img src="img/000023.gif" width="73" height="16" border="0" align="absmiddle"></a>
			</td>
		</tr>

	</table>

	</center>

	<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
	<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
	<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">
	<input type="hidden" id="strID" name="strID" value="<%=strID%>">
	<input type="hidden" id="strLogin" name="strLogin" value="<%=strLogin%>">
	<input type="hidden" id="strDataIni" name="strDataIni" maxLength=255 size=65 value="<%=strInicio%>" >
	<input type=hidden id="strDataFim" name="strDataFim" maxLength=255 size=65 value="<%=strFim%>" >

</form>
<SCRIPT language=JavaScript>

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
			document.frmIncluirTarefa.txtDtInicio.focus();
		}
		else
		{
			document.frmIncluirTarefa.txtDtFim.focus();
		}				
	    return false;
	}
}


function verifica_numerico(checkStr,Verifica){
	var checkOK = "0123456789";
	var decPoints = 0;
	var allNum = "";
	for (i = 0; i < checkStr.length; i++){
	   ch = checkStr.charAt(i);
   		for (j = 0; j < checkOK.length; j++)
       		if (ch == checkOK.charAt(j)) break;
   			if (j == checkOK.length) {
		        Verifica = "Ko";
		        break;
			}
		 	allNum += ch;
		}
	return (Verifica);
}

function Validar()
{
	var erro = ""
	
	if (document.frmIncluirTarefa.txtNome.value == "")
	{
		erro = erro + "Preencha o campo Nome." + "\n"
	}


	if (document.frmIncluirTarefa.txtDtInicio.value == "")
	{
		erro = erro + "Preencha o campo Data Inicio." + "\n"
	}


	if (document.frmIncluirTarefa.txtDtFim.value == "")
	{
		erro = erro + "Preencha o campo Data Fim." + "\n"
	}

	if (isNaN(document.frmIncluirTarefa.txtSequencia.value))
	{
		erro = erro + "Campo Seqüência deverá ser numérico." + "\n"
	}
	else
	{
		if (verifica_numerico(document.frmIncluirTarefa.txtSequencia.value,false) == "Ko")
		{
			erro = erro + "Campo Seqüência deverá ser inteiro." + "\n"
		}

	}

	if (document.frmIncluirTarefa.txtSequencia.value == "")
	{
		erro = erro + "Preencha o campo Seqüência." + "\n"
	}
	else
	{
		if (document.frmIncluirTarefa.txtSequencia.value == 0)
		{
			erro = erro + "Campo Seqüência deverá ser diferente de zero." + "\n"
		}
	}

	if (document.frmIncluirTarefa.txtNome.value == "")
	{
		erro = erro + "Preencha o campo motivo." + "\n"
	}


	if (erro != "") 
	{
		alert(erro)
		return false;
	}
	else
	{
		return true;
	}		

return true;

}


function Confirmar()
{
	if (Validar())
	{
		document.frmIncluirTarefa.hidOperacao.value = 'A';
		document.frmIncluirTarefa.submit();
	}
}

</SCRIPT>
</body>
</html>

