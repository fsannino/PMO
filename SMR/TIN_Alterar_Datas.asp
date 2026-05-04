<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strProjeto
Dim strUID
Dim strSql
Dim strOperacao
Dim strDataIni
Dim strDataFim
Dim strLogin

Dim rs

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,7,4) + mid(strData,4,2) + mid(strData,1,2)
End Function

	strProjeto	= Request("strProjeto")
	strUID	= Request("strUID")
	strInicio = Request("strDataIni")
	strFim = Request("strDataFim")
	strLogin = Request("strLogin")
	
	strOperacao	= Request("hidOperacao")
	strDataIni = Request("txtDataIni")
	strDataFim = Request("txtDataFim")
	
	'Abrindo uma conexão com o BD
	set conConexao = TIN_AbrirConexaoBD()

	if strOperacao <> "" Then

		strSql = "EXEC SP_ATUALIZAR_DATAS_TAREFAS " & strProjeto & ", " & strUID & ", '" & FormatarDataSQL(strDataIni) & "', '" & FormatarDataSQL(strDataFim) & "' "
	
		conConexao.execute strSql

		strSql = "EXEC SP_INCLUIR_LOG_ALTERACAO_TAREFAS '" & strLogin & "', " & strProjeto & ", " & strUID & ", '" & FormatarDataSQL(strDataIni) & "', '" & FormatarDataSQL(strDataFim) & "' "
	
		conConexao.execute strSql


	Else

		strSql = "SP_LISTAR_TAREFAS " & strProjeto & ", " & strUID & " "

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao
	
	End if

%>
<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmAlterarDatas" ID="frmAlterarDatas" action="TIN_Alterar_Datas.asp" method="post">
<%if trim(strOperacao) = "" Then%>
	<center>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>


	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
	    <td class="xl27" width="340px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
	    <td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
	    <td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
	    <td class="xl27" width="30px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
	  </tr>

		<%Do While Not rs.EOF%>
			
			<tr height="17" style="height:12.75pt">
			  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
			  	<font face="Arial" size="1">
			  		<%=rs("TASK_UID")%>
			  	</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="340px">
			  	<font face="Arial" size="1">
			  		<%=rs("TASK_NAME")%>&nbsp;
			  	</font>
			  </td>
			  <td class="xl28" style="border: 1 solid #666666" width="80px" align="right" >
			  	<font face="Arial" size="1">
			  		<%=FormatDateTime(rs("TASK_START_DATE"),2)%>&nbsp;
			  	</font>
			  </td>
			  <td class="xl28" style="border: 1 solid #666666" width="80px" align="right" >
			  	<font face="Arial" size="1">
			  		<%=FormatDateTime(rs("TASK_FINISH_DATE"),2)%>&nbsp;
			  	</font>
			  </td>
			  <td class="xl28" style="border: 1 solid #666666" width="30px" align="right" >
			  	<font face="Arial" size="1">
			  		<%=rs("TASK_PCT_COMP")%>&nbsp;
			  	</font>
			  </td>
			</tr>
	</table>
	<BR>
	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
		<tr height="17" style="height:12.75pt">
			<td height="17" class="xl27" width=50% style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data Inicio</font></b></td>
			<td height="17" class="xl27" width=50% style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data Fim</font></b></td>
		</tr>
		<tr height="17" style="height:12.75pt">
			<td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
			  	<font face="Arial" size="1" color="#0000FF">
					<%If rs("DataIni") = "01/01/1900" Then%>
  						<input id="txtDataIni" name="txtDataIni" value="" maxLength=10 size=10 onblur="ValidaDatas(this,'I');">
					<%Else%>
  						<input id="txtDataIni" name="txtDataIni" value="<%=rs("DataIni")%>" maxLength=10 size=10 onblur="ValidaDatas(this,'I');">
					<%End If%>					
			  		<input type="hidden" id="hidDataIni" name="hidDataIni" value="<%=rs("DataIni")%>">
			  	</font>
			  </td>

			<td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
			  	<font face="Arial" size="1" color="#0000FF">
					<%If rs("DataFim") = "01/01/1900" Then%>
	  					<input id="txtDataFim" name="txtDataFim" value="" maxLength=10 size=10 onblur="ValidaDatas(this,'F');">
					<%Else%>
	  					<input id="txtDataFim" name="txtDataFim" value="<%=rs("DataFim")%>" maxLength=10 size=10 onblur="ValidaDatas(this,'F');">
					<%End If%>					
			  		<input type="hidden" id="hidDataFim" name="hidDataFim" value="<%=rs("DataFim")%>">
			  	</font>
			  </td>

		</tr>

	</table>

	<%rs.MoveNext%>
	<%Loop%>
	<BR>
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
	<input type="hidden" id="strLogin" name="strLogin" value="<%=strLogin%>">

<%Else%>
	<%strOperacao = ""%>
	<SCRIPT language=JavaScript>
		this.close();
	</SCRIPT>
<%End If%>
</form>
<SCRIPT language=JavaScript>

function Confirmar()
{
	if (Validar())
	{
		document.frmAlterarDatas.hidOperacao.value = 'A';
		document.frmAlterarDatas.submit();
	}
}

function Validar()
{
	var erro = ""
	var datafim = document.frmAlterarDatas.txtDataFim.value.substr(6,4) + document.frmAlterarDatas.txtDataFim.value.substr(3,2) +  document.frmAlterarDatas.txtDataFim.value.substr(0,2)
	var datainicio = document.frmAlterarDatas.txtDataIni.value.substr(6,4) + document.frmAlterarDatas.txtDataIni.value.substr(3,2) +  document.frmAlterarDatas.txtDataIni.value.substr(0,2)
	
	if (datafim < datainicio)
	{
		erro = erro + "Data inicial deverá ser menor ou igual que a final." + "\n"
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
	else
	{
	    return false;
	}

	if (bol){
		alert(msg_erro);

		if (Tipo == "I")
		{
			document.frmAlterarDatas.txtDataIni.value = document.frmAlterarDatas.hidDataIni.value;
			document.frmAlterarDatas.txtDataIni.focus();
		}
		else
		{
			document.frmAlterarDatas.txtDataFim.value = document.frmAlterarDatas.hidDataFim.value;
			document.frmAlterarDatas.txtDataFim.focus();
		}				
	    return false;
	}
}

</SCRIPT>
</body>
</html>
