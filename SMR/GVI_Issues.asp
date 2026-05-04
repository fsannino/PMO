<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strProjeto
Dim strUID
Dim strSql
Dim strOperacao
Dim strDescricao

Dim rs

	strProjeto		= Request("strProjeto")
	strUID			= Request("strUID")
	strRedirect		= "./GVI_Issues.asp?strProjeto=" & strProjeto & "þstrUID=" & strUID
	
	if trim(session("Usuario"))="" then
		response.Redirect("./LOGIN.ASP?hidOrigem=" & strRedirect)
	end if

	strProjeto		= Request("strProjeto")
	strUID			= Request("strUID")
	strOperacao		= Request("hidOperacao")
	strNome			= Request("txtNome")
	strDescricao	= Request("txtDescricao")
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	If strOperacao <> "" Then

		strSql = "EXEC SP_INCLUIR_ISSUES_GVI " & _
			strProjeto & ", " & _
			strUID & ", '" & _
			ReplicaPlics(strNome) & "', '" & _
			ReplicaPlics(strDescricao) & "', '" & _
			session("Usuario") & "' "

		conConexao.execute strSql

		strOperacao = ""%>
		<SCRIPT language=JavaScript>
			this.close();
		</SCRIPT>

	<%End If%>

<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmIssues" ID="frmIssues" action="GVI_Issues.asp" method="post">
	<center>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="21" bgcolor="#6699CC">
	<table width="600" border="0" cellspacing="0" cellpadding="0">
	        <tr>
	          <td width="143"> 
	            <div align="right"><img src="img/000002.gif" width="124" height="21"><img src="img/000003.gif" width="19" height="21"></div></td>
	          <td width="456" height="21"><img src="img/_0.gif" width="1" height="1"></td>
	        </tr>
	      </table>
					  
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr> 
		<TD></TD>
		<TD>
			<p><img src="img/000061.gif" width="105" height="40"></p>
		</TD>
		<TD></TD>
	  </tr>
	</table>
	<TABLE  name="tblDados" id="tblDados" border="0" ALIGN="Left">
		<TR>
			<TD>
				<font face=arial size=1>Nome da Issue:</font>
			</TD>
		</TR>
		<TR>
			<TD>
				<INPUT type="text" id="txtNome" name="txtNome" tabindex="1" size=65>
			</TD>
		</TR>
		<TR>
			<TD>
				<font face=arial size=1>Descrição:
			</TD>
		</TR>
		<TR>
			<TD>
				<TEXTAREA id="txtDescricao" name="txtDescricao" cols="70" rows="4" maxlength="512" tabindex="2" 
				onkeyup="if (this.value.length > 512)
				         {
							this.value = this.value.substr(0, 512);
							alert('O campo descrição deve ter no máximo 512 caracteres');
							this.focus();
						}"
				onblur="if (this.value.length > 512)
						{
							alert('O campo descrição deve ter no máximo 512 caracteres');
							this.value = this.value.substr(0, 512);
							this.focus();
						}"></TEXTAREA>
			</TD>  
		</TR>

		<tr>
			<td align=right>
<!--				<input type="Image" name="cmdConfirmar" value="Confirmar" src="img/000049.gif" align="absmiddle" onclick="Confirmar();" tabindex="3">-->
				<a href="javascript:this.Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0" align="absmiddle" tabindex="4"></a>
				<a href="javascript:this.close();" ><img src="img/000023.gif" width="73" height="16" border="0" align="absmiddle" tabindex="4"></a>
			</td>
			<TD></TD>
		</tr>
					
	</TABLE>

	</center>

	<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
	<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
	<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">

</form>
<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmIssues.hidOperacao.value = 'A';
	document.frmIssues.submit();
}

</SCRIPT>
</body>
</html>
