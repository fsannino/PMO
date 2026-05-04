<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strCodIssue
Dim strProjeto
Dim strUID
Dim strIntegrador
Dim strTipo
Dim strOperacao

Dim cmdResultado
Dim strSql
Dim rs
Dim rs1

Dim strComentario
Dim strFechamento
Dim strTitulo

	strCodIssue		= Request("strCodIssue")
	strProjeto		= Request("strProjeto")
	strUID			= Request("strUID")
	strIntegrador			= Request("strIntegrador")
	strTipo			= Request("strTipo")

	strOperacao		= Request("hidOperacao")
	strComentario	= Request("txtComentario")
	strFechamento	= Request("chkFechamento")
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	If strTipo = "A" Then

		strTitulo = "Atualização do Problema"
		
		strRedirect		= "./GVI_Comentario_Issues.asp?strCodIssue=" & strCodIssue & "þstrProjeto=" & strProjeto & "þstrUID=" & strUID & "þstrIntegrador=" & strIntegrador & "þstrTipo=" & strTipo
	
		if trim(session("Usuario"))="" then
			response.Redirect("./LOGIN.ASP?hidOrigem=" & strRedirect)
		end if

		If isnull(session("Integrador")) OR trim(session("Integrador")) <> strIntegrador then

			Session("Usuario") = ""
			Session("UsuarioCLI") = ""
			Session("Unidade") = ""
			Session("Integrador") = ""
			Session("Desc_Unidade") = ""

			Response.Redirect ("./Erro.asp?Erro=Usuário não tem acesso para atualizar a issue.&Voltar=true&IrPara=GVI_Relatorios_Issues.asp")

		End if
		
		If strOperacao <> "" Then

			Set cmdResultado = Server.CreateObject("ADODB.Command")
    
			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 240
			    .CommandText = "SP_ATUALIZAR_ISSUES_GVI"

			    .Parameters.Refresh

				.Parameters(1).Value = cint(strCodIssue)
				.Parameters(2).Value = strComentario
				
				If Trim(strFechamento) <> "" Then
					.Parameters(3).Value = 1
					.Parameters(4).Value = Trim(Session("Usuario"))
				Else
					.Parameters(3).Value = Null
					.Parameters(4).Value = Null
				End If
						
			End With

			cmdResultado.Execute()
			
			strOperacao = ""

			response.Redirect("./GVI_Relatorios_Issues.asp?")
			
		End If
	Else
		strTitulo = "Visualização do Problema"	
	End If

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 240
        .CommandText = "SP_LISTAR_TAREFAS_ISSUES_GVI"

        .Parameters.Refresh

		.Parameters(1).Value = cint(strProjeto)

		If isnumeric(strUID) Then
			.Parameters(2).Value = strUID
		Else
			.Parameters(2).Value = cint(strUID)		
		End If			
				
    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()


    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 240
        .CommandText = "SP_LISTAR_ISSUES_GVI"

        .Parameters.Refresh

		.Parameters(1).Value = cint(strProjeto)
		If isnumeric(strUID) Then
			.Parameters(2).Value = strUID
		Else
			.Parameters(2).Value = cint(strUID)		
		End If			
		.Parameters(3).Value = cint(strCodIssue)
				
    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()
	
	
	%>

<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmComentarioIssues" ID="frmComentarioIssues" action="GVI_Comentario_Issues.asp" method="post">

	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<%If Not rs.EOF Then%>
		<p>

		<table cellspacing="0" cellpadding="0" align=center>
			<tr height="17"  >
				<td align=left width="750px">
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><%=strTitulo%> &gt;<%=rs("PROJ_NAME")%></font></b>
				</td>
			</tr>
		</table>

		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0"  align=center>
				  
			<tr height="17" style="height:12.75pt">
			<td height="17" class="xl27" width=50px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
			<td class="xl27" width="300px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
			<td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
			<td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
			<td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
			<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Integrador</font></b></td>
			<td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">CLI</font></b></td>
			</tr>

			<tr height="17" style="height:12.75pt">
			  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=50px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="300px">
					<font face="Arial" size="1">
							<%=rs("TASK_NAME")%>&nbsp;
					</font>
				</td>

			  <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>
			  <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>

				<td class="xl30" style="border: 1 solid #666666" width="100px" align=center>
				<%if isnull(rs("Integrador")) or trim(rs("Integrador")) = "" Then%>
					<font face="Arial" size="1">&nbsp;</font></td>
				<%Else%>
					<font face="Arial" size="1"><%=rs("Integrador")%></font>
				<%End If%>
				</td>


			  <td class="xl22" style="border: 1 solid #666666" width="150px" align=center>
			  <%if isnull(rs("CLI")) or trim(rs("CLI")) = "" Then%>
					<font face="Arial" size="1">&nbsp;</font></td>
			  <%Else%>
					<font face="Arial" size="1"><%=rs("CLI")%></font></td>
			  <%End If%>

			</tr>
		</table>

		<BR>

		<table cellspacing="0" cellpadding="0" align=center>
			<tr height="17"  >
				<td align=left width="750px">
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Dados do Problema</font></b>
				</td>
			</tr>
		</table>

		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0"  align=center>
				  
			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=200px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				  <font color=White size="1" face="Georgia, Times New Roman, Times, serif">Título do Issue</font></b></td>

				<td class="xl28" style="border: 1 solid #666666" width="575px">
				  	<font face="Arial" size="1">
				  			<%=rs1("Nome_Issue")%>&nbsp;
				  	</font>
				  </td>
			</tr>
			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=200px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				  <font color=White size="1" face="Georgia, Times New Roman, Times, serif">Descrição do Issue</font></b></td>

				<td class="xl28" style="border: 1 solid #666666" width="575px">
				  	<font face="Arial" size="1">
						<TEXTAREA id="txtDescricao" name="txtDescricao" cols="69" rows="5" maxlength="512" tabindex="2" 
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
								}" readonly><%=rs1("Desc_Issue")%></TEXTAREA>

				  	</font>
				  </td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=200px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				  <font color=White size="1" face="Georgia, Times New Roman, Times, serif">Comentário do Issue</font></b></td>

				<td class="xl28" style="border: 1 solid #666666" width="575px">
				  	<font face="Arial" size="1">
						<%If  strTipo = "A" Then%>
							<TEXTAREA id="txtComentario" name="txtComentario" cols="69" rows="5" maxlength="255" tabindex="2" 
							onkeyup="if (this.value.length > 255)
							         {
										this.value = this.value.substr(0, 255);
										alert('O campo comentário deve ter no máximo 255 caracteres');
										this.focus();
									}"
							onblur="if (this.value.length > 255)
									{
										alert('O campo comentário deve ter no máximo 255 caracteres');
										this.value = this.value.substr(0, 255);
										this.focus();
									}"><%=rs1("Coment_Issue")%></TEXTAREA>

						<%Else%>
							<TEXTAREA id="txtComentario" name="txtComentario" cols="69" rows="5" maxlength="255" tabindex="2" 
							onkeyup="if (this.value.length > 255)
							         {
										this.value = this.value.substr(0, 255);
										alert('O campo comentário deve ter no máximo 255 caracteres');
										this.focus();
									}"
							onblur="if (this.value.length > 255)
									{
										alert('O campo comentário deve ter no máximo 255 caracteres');
										this.value = this.value.substr(0, 255);
										this.focus();
									}"  readonly><%=rs1("Coment_Issue")%></TEXTAREA>

						<%End If%>
				  	</font>
				  </td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=200px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				  <font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fechamento</font></b></td>

				<td class="xl28" style="border: 1 solid #666666" width="575px">
				  	<font face="Arial" size="1">
						<%If  strTipo = "A" Then%>
							<INPUT type="checkbox" id=chkFechamento name=chkFechamento>
						<%Else%>
							<INPUT type="checkbox" id=chkFechamento name=chkFechamento  disabled>
						<%End If%>
	
				  	</font>
				  </td>
			</tr>
		</table>
		<BR>
		<table align=right>
			<tr>
				<td align=right>
					<p align="center">
					<%If  strTipo = "A" Then%>
						<a href="JavaScript:Confirmar();">
							<img src="img/000049.gif" name="Issues" border="0">
						</a>

<!--						<input type="image" name="cmdConfirmar" value="Confirmar" src="img/000049.gif" onClick="Confirmar();" align="right">-->
					<%End If%>
				</td>

				<td align=right>
					<p align="center">
					<a href="JavaScript:Voltar();">
						<img src="img/000024.gif" name="Issues" border="0">
					</a>

<!--					<input type="image" name="cmdVoltar" value="Confirmar" src="img/000024.gif" onClick="Voltar();" align="right">-->
				</td>
				<td width="110px">
				</td>
			</tr>
		</table>


	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)
%>

		<input type="hidden" id="strCodIssue" name="strCodIssue" value="<%=strCodIssue%>">
		<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
		<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">
		<input type="hidden" id="strIntegrador" name="strIntegrador" value="<%=strIntegrador%>">
		<input type="hidden" id="strTipo" name="strTipo" value="<%=strTipo%>">
		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">

</form>
<SCRIPT language=JavaScript>

function Confirmar()
{

	if (document.frmComentarioIssues.chkFechamento.checked == true && 
		document.frmComentarioIssues.txtComentario.value == '')
	{
		alert("Comentário obrigatório para fechamento.");
	}
	else
	{
		document.frmComentarioIssues.hidOperacao.value = 'A';
		document.frmComentarioIssues.action = 'GVI_Comentario_Issues.asp';
		document.frmComentarioIssues.submit();
	}
}

function Voltar()
{
	document.frmComentarioIssues.action = 'GVI_Relatorios_Issues.asp' ;
	document.frmComentarioIssues.submit();
}

</SCRIPT>
</body>
</html>
