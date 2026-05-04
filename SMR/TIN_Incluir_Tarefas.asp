<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strProjeto
Dim strUID
Dim strID
Dim strSql
Dim strOperacao
Dim strDataIni
Dim strDataFim
Dim strLogin
Dim strNome
Dim strResposavel
Dim strCenario
Dim strTransacao
Dim strSequencia

Dim rs

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,7,4) + mid(strData,4,2) + mid(strData,1,2)
End Function

	strProjeto = Request("strProjeto")
	strUID	   = Request("strUID")
	strID	   = Request("strID")
	strDataIni = Request("strDataIni")
	strDataFim = Request("strDataFim")
	strLogin   = Request("strLogin")

	if trim(strDataIni) = "" Then
		strDataIni = Request("txtDataIni")
		strDataFim = Request("txtDataFim")
	End If

	strNome        = Request("txtNome")
	strResposavel  = Request("txtResponsavel")
	strCenario     = Request("txtCenario")
	strTransacao   = Request("txtTransacao")
	strSequencia   = Request("txtSequencia")

	strOperacao	= Request("hidOperacao")
	
	'Abrindo uma conexão com o BD
	set conConexao = TIN_AbrirConexaoBD()

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
			strSql = "EXEC SP_INCLUIR_NOVAS_TAREFAS " & strProjeto &_
			         ", " & strUID & ", " &	strID  & ", '" & Trim(strNome) &_
			         "', '" & Trim(strResposavel) & "', '" & Trim(strCenario) &_
			         "', '" & Trim(strTransacao) & "', '" & Trim(strSequencia) & "' "

			conConexao.execute strSql

			strSql = "EXEC SP_INCLUIR_LOG_NOVAS_TAREFAS '" & strLogin & "', " & strProjeto &_
			         ", " & strUID & ", " &	strID  & ", '" & Trim(strNome) &_
			         "', '" & Trim(strResposavel) & "', '" & Trim(strCenario) &_
			         "', '" & Trim(strTransacao) & "', '" & Trim(strSequencia) & "' "

			conConexao.execute strSql

		End if

	end if

	strSql = "SP_LISTAR_NOVAS_TAREFAS " & strProjeto & ", " & strUID & " "

	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.OPEN STRSQL, conConexao
	
	intConta = 0

	do while not rs.EOF
		intConta = intConta + 1
		rs.MoveNext
	loop
	
	strSql = "SP_LISTAR_NOVAS_TAREFAS " & strProjeto & ", " & strUID & " "

	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.OPEN STRSQL, conConexao
	
%>
<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmIncluirTarefa" ID="frmIncluirTarefa" action="TIN_Incluir_Tarefas.asp" method="post">

	<input type="hidden" id="txtDataIni" name="txtDataIni" maxLength=255 size=65 value="<%=strDataIni%>" >
	<input type=hidden id="txtDataFim" name="txtDataFim" maxLength=255 size=65 value="<%=strDataFim%>" >

<%'if trim(strOperacao) = "" Then%>
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
					<%=FormatarDataMon(strDataIni)%>&nbsp;
			  	</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="85px" align=center>
			  	<font face="Arial" size="1">
					<%=FormatarDataMon(strDataFim)%>&nbsp;
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
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cod. Cenário</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtCenario" name="txtCenario" value="" maxLength=11 size=65>
			</font>
		</td>
		</tr>
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cod. Transação</font></b></td>
		<td class="xl28" style="border: 1 solid #666666" width="400px" align=left>
			<font face="Arial" size="1">
				<input id="txtTransacao" name="txtTransacao" value="" maxLength=50 size=65>
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

<%'Else%>
	<%'strOperacao = ""%>
	<SCRIPT language=JavaScript>
		//this.close();
	</SCRIPT>
<%'End If%>
</form>
<SCRIPT language=JavaScript>

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

	if (document.frmIncluirTarefa.txtTransacao.value == "")
	{
		erro = erro + "Preencha o campo Cod. Transação." + "\n"
	}

	if (document.frmIncluirTarefa.txtCenario.value == "")
	{
		erro = erro + "Preencha o campo Cod. Cenário." + "\n"
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

<%
'	If intConta >= 9 Then
%>	
<!--<SCRIPT language=JavaScript>
	alert("Excedeu o limite de nove tarefas.")
	this.close()
</SCRIPT>-->

<%'end if%>