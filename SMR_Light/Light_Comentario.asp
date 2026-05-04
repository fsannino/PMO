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

Dim cmdResultado
Dim rs

	strProjeto		= Request("strProjeto")
	strUID			= Request("strUID")
	strRedirect		= "./Light_Comentario.asp?strProjeto=" & strProjeto & "þstrUID=" & strUID
	
	if trim(session("Usuario"))="" then
		response.Redirect("./Light_Login.ASP?hidOrigem=" & strRedirect)
	end if

	strProjeto		= Request("strProjeto")
	strUID			= Request("strUID")
	strOperacao		= Request("hidOperacao")
	strNome			= Request("txtNome")
	
	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

	If strOperacao <> "" Then

		strSql = "EXEC SP_ATUALIZAR_COMENTARIOS " & _
			strProjeto & ", " & _
			strUID & ", '" & _
			ReplicaPlics(strNome) & "'"
			
		conConexao.execute strSql

		strOperacao = ""%>
		<SCRIPT language=JavaScript>
			this.close();
		</SCRIPT>
	<%End If%>


<%
	Set cmdResultado = Server.CreateObject("ADODB.Command")
    
	With cmdResultado
    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 480
	    .CommandText = "SP_LISTAR_COMENTARIOS"

		.Parameters.Refresh
		.Parameters(1).Value = strProjeto
		.Parameters(2).Value = strUID

	End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

%>

<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmComentarios" ID="frmComentarios" action="Light_Comentario.asp" method="post">
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
			<p><!--<img src="img/000061.gif" width="105" height="40">--><b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Comentários</font></b></p>
		</TD>
		<TD></TD>
	  </tr>
	</table>
	<BR>
	<TABLE  name="tblDados" id="tblDados" border="0" ALIGN="Left">
		<TR>
			<TD>
				<%If Not rs.EOF Then%>

					<TEXTAREA id="txtNome" name="txtNome" cols="70" rows="4" maxlength="255" tabindex="2" 
					onkeyup="if (this.value.length > 255)
					         {
								this.value = this.value.substr(0, 255);
								alert('O campo descrição deve ter no máximo 512 caracteres');
								this.focus();
							}"
					onblur="if (this.value.length > 255)
							{
								alert('O campo descrição deve ter no máximo 512 caracteres');
								this.value = this.value.substr(0, 255);
								this.focus();
							}"><%=rs("Comentario")%></TEXTAREA>




<!--					<INPUT type="text" id="txtNome" name="txtNome" tabindex="1" size=95 maxlength=255 value="<%=rs("Comentario")%>">-->
				<%Else%>

					<TEXTAREA id="txtNome" name="txtNome" cols="70" rows="4" maxlength="255" tabindex="2" 
					onkeyup="if (this.value.length > 255)
					         {
								this.value = this.value.substr(0, 255);
								alert('O campo descrição deve ter no máximo 512 caracteres');
								this.focus();
							}"
					onblur="if (this.value.length > 255)
							{
								alert('O campo descrição deve ter no máximo 512 caracteres');
								this.value = this.value.substr(0, 255);
								this.focus();
							}"></TEXTAREA>


<!--					<INPUT type="text" id="txtNome" name="txtNome" tabindex="1" size=95 maxlength=255 >-->
				<%End If%>				
			</TD>
		</TR>

		<tr>
			<td align=right>
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
	document.frmComentarios.hidOperacao.value = 'A';
	document.frmComentarios.submit();
}

</SCRIPT>
</body>
</html>
