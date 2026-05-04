<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strUsuario, strSenha, rs, strOrigem

	strUsuario	= Request("txtUsuario")
	strSenha	= Request("txtSenha")
	strOrigem	= Request("hidOrigem")
	strOrigem	= Replace(strOrigem, "¬", " ")

	if (strUsuario <> "" and strSenha <> "") then

		strUsuario	= lcase(TRIM(strUsuario))
		strSenha	= TRIM(strSenha)

		'Abrindo uma conexão com o BD
		set conConexao = SMR_AbrirConexaoBD()

		strSql = "SP_LISTAR_USUARIO '" & strUsuario & "'"

		set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open strSql, conConexao

		if rs.EOF Then
			Response.Redirect ("./Erro.asp?Erro=Usuário não existe.&Voltar=true")
		else
			if rs("Senha") <> strSenha Then
				Response.Redirect ("./Erro.asp?Erro=Senha inválida.&Voltar=true")
			else
				Session("Usuario") = rs("Login")
				Session("UsuarioCLI") = trim(rs("Nome_Usuario"))
				Session("Unidade") = rs("Cod_Unidade")
				Session("Integrador") = UCase(trim(rs("Integrador")))
				
				If rs("Flag_CLI") = true Then
					If Trim(strOrigem) = "" Then 
						strOrigem = "./GVI_Medicao_Detalhe.asp"
					Else
						strOrigem = Replace(strOrigem,"þ","&")
					End If
				Else
					If Trim(strOrigem) = "" Then 
						strOrigem = "./PMO_Medicao_Detalhe.asp"
					Else
						strOrigem = Replace(strOrigem,"þ","&")
					End If
				End if
				Response.Redirect (strOrigem)
			end if

		end if
	end if

%>
<html>
<head>

<script language="javascript">
function realizarLogin() {
	if (isCampoPreenchido(frmLogin.txtUsuario, "Usuário") &&
		isCampoPreenchido(frmLogin.txtSenha, "Senha")) {
		frmLogin.action = "./Login.asp";
		return true;
	} else {
		return false;
	}
}

function isCampoPreenchido(controle, nome) {
	if (controle.value=="") {
		controle.focus();
		alert("O campo " + nome + " é obrigatório e não foi preenchido.");
		return false;
	} else {
		return true;
	}
}

</script>
</head>
<body onload="document.frmLogin.txtUsuario.focus();">
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmLogin" ID="frmLogin" method="post" onsubmit="return realizarLogin();">
<INPUT type="hidden" id=hidOrigem name=hidOrigem value=<%=Replace(strOrigem, " ", "¬")%>>

<center>

<table style="border-collapse: collapse; border: 1px solid #C0C0C0" border="0" bordercolor="#111111" cellpadding="0" cellspacing="0">
<tr bgcolor=#6699cc>
	<th>
	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;Informe seus dados para entrar no SMR...</font>
	</th>
</tr>
<tr>
	<td>
		<table>
		<tr>
			<td align="right" width=100px>
				<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
				Usuário:
				</font>
			</td>
			<td align="left"width=190px>
				<input name="txtUsuario" size="15" maxlength="25">
			</td>
		</tr>
		<tr>
			<td align="right" width=100px>
				<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
				Senha:
				</font>
			</td>
			<td align="left" width=190px>
				<input type="password" name="txtSenha" size="15" maxlength="10">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align=right>
		<p align="center">
		<input type="image" name="cmdLogin" value="Confirmar" src="img/000049.gif" onClick="Confirmar();" align="right">
	</td>
</tr>
</table>

</center>
</form>
<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmLogin.action = "Login.asp?hidOrigem=<%=hidOrigem%>";
	document.frmLogin.submit();
}

</SCRIPT>

</body>
</html>
