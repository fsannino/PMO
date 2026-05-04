<!--#include file="funcoes/Funcoes.inc"-->

<% 
'Ordena ao browser a abertura do MS-Excel
Response.ContentType = "application/vnd.ms-excel"

dim strTodosProjetos
dim strProjeto
dim strUsuario
dim strTipo
dim strStatus
dim strPrioridade
dim strComite
'dim strDataLimite
dim strEquipe
dim strRDO
dim strTipoAux
dim strAcesso
dim strPesquisaAvancada
dim strOpLogico1
dim strCampoAPesquisar
dim strInfoASerPesquisada

dim rsRelatorios

'Recuperando dados da página de filtro
strTodosProjetos = Request.Form("chkTodosProjetos")
strProjeto = Request.Form("chkProjeto")
strUsuario = Request.Form("slcUsuario")
strTipo = Request.Form("slcTipo")
strStatus = Request.Form("slcStatus")
strPrioridade = Request.Form("slcPrioridade")
strComite = Request.Form("slcComite")
'strDataLimite = request.form("slcDataLimite")
strEquipe = Request.Form("slcEquipe")
strRDO = ucase(trim(Request.Form("chkRDO")))
strAcesso = "PUBLIC"
'strPesquisaAvancada = Request.Form("chkPesquisaAvancada")
'strOpLogico1 = Request.Form("slcOpLogico1")
strCampoAPesquisar = Request.Form("slcCampoAPesquisar")
strInfoASerPesquisada = Request.Form("txtInfoASerPesquisada")

If Trim(strCampoAPesquisar) <> "" And Trim(strInfoASerPesquisada) <> "" Then
	strPesquisaAvancada = "True"
	strOpLogico1 = "AND"
End If


'Organizando os códigos dos projetos para que eles possam ser enviados para a Stored Procedure
if ucase(trim(strTodosProjetos)) <> "TRUE" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strProjeto = TrataStringSP(strProjeto, ",", "S")
else
	strProjeto = trim(strTodosProjetos)
end if

'Organizando os tipos para que eles possam ser enviados para a Stored Procedure
if trim(strTipo) <> "" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strTipo = TrataStringSP(strTipo, ",", "S")
end if

'Organizando os códigos dos usuários para que eles possam ser enviados para a Stored Procedure
if trim(strUsuario) <> "" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strUsuario = TrataStringSP(strUsuario, ",", "N")
end if

'Organizando os códigos das equipes para que eles possam ser enviados para a Stored Procedure
if trim(strEquipe) <> "" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strEquipe = TrataStringSP(strEquipe, ",", "N")
end if

'Organizando os códigos dos comitês para que eles possam ser enviados para a Stored Procedure
if trim(strComite) <> "" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strComite = TrataStringSP(strComite, "," ,"N")
end if

'Organizando os códigos das prioridades para que eles possam ser enviados para a Stored Procedure
if trim(strPrioridade) <> "" then	
	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	strPrioridade = TrataStringSP(strPrioridade, ",", "S")
end if

'Buscando informações nas tabelas de acordo com o filtro escolhido
'strSQL = "EXECUTE SP_RELATORIOS " & _ 
'		 "'" & strProjeto &  "', " & _
'		 "'" & strUsuario &  "', " & _
'		 "'" & strTipo &  "', " & _
'		 "'" & strStatus &  "', " & _
'		 "'" & strPrioridade &  "', " & _
'		 "'" & strComite &  "', " & _
'		 "'" & strDataLimite &  "', " & _
'		 "'" & strEquipe &  "', " & _
'		 "'" & strRDO & "', " & _
'		 "'" & strAcesso & "'"

strSQL = "EXECUTE SP_RELATORIOS " & _ 
		 "'" & strProjeto &  "', " & _
		 "'" & strUsuario &  "', " & _
		 "'" & strTipo &  "', " & _
		 "'" & strStatus &  "', " & _
		 "'" & strPrioridade &  "', " & _
		 "'" & strComite &  "', " & _
		 "'" & strEquipe &  "', " & _
		 "'" & strRDO & "', " & _
		 "'" & strAcesso & "', " & _
		 "'" & strPesquisaAvancada & "', " & _
		 "'" & strOpLogico1 & "', " & _
		 "'" & strCampoAPesquisar & "', " & _
		 "'" & strInfoASerPesquisada & "'"

'Abrindo uma conexão com o BD
set conConexao = LIGHT_AbrirConexaoBD()

set rsRelatorios = Server.CreateObject("ADODB.RecordSet")

rsRelatorios.Open strSQL, conConexao

%>

<html>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<% 
if not rsRelatorios.EOF then
%>
	<div align="center">
		<center>
			<table border="0" width="90%" cellspacing="0" cellpadding="0">
				<tr>
					<td align="center">
						<font face="Arial" size="1">PROJETO</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">USUÁRIO</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">ABERTO EM</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">ID</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">ETAPA DE WORKFLOW</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">DATA LIMITE</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">NOME</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">PRIORIDADE</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">DATA DE FECHAMENTO</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">EQUIPE</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">COMITÊ</font>
					</td>
					<td align="center">
						<font face="Arial" size="1">STATUS</font>
					</td>

					<%If Trim(strTipo) = """Risks""" Then%>

						<td align="center">
							<font face="Arial" size="1">RDO</font>
						</td>
						<td align="center">
							<font face="Arial" size="1">TIPO DE RISCO</font>
						</td>
						<td align="center">
							<font face="Arial" size="1">DASHBOARD</font>
						</td>

					<%End If%>

				</tr>
	<%
	do while not rsRelatorios.EOF 
	%>
				<tr>
      				<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Desc_Projeto")%>&nbsp;</font>
					</td>
	      			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Nome_Funcionario")%>&nbsp;</font>
					</td>
				    <td align="right">
						<font face="Arial" size="1">
						<%
						if (rsRelatorios("Aberto_em") <> cdate(0)) and (rsRelatorios("Aberto_em") <> cdate("01/01/1900")) then
							Response.Write rsRelatorios("Aberto_em")
						else
							Response.Write "&nbsp;"
						end if
						%>
						</font>
					</td>
    	  			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("ID")%>&nbsp;</font>
					</td>
     					<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Workflow")%>&nbsp;</font>
					</td>
     					<td align="right">
						<font face="Arial" size="1">
						<%
						if (rsRelatorios("Data_Limite") <> cdate(0)) and (rsRelatorios("Data_Limite") <> cdate("01/01/1900")) then
							Response.Write rsRelatorios("Data_Limite")
						else
							Response.Write "&nbsp;"
						end if
						%>
						</font>
					</td>
   		  			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Name")%>&nbsp;</font>
					</td>
    	  			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Priority")%>&nbsp;</font>
					</td>
    	  			<td align="right">
						<font face="Arial" size="1">
						<%
						if (rsRelatorios("Data_Fechamento") <> cdate(0)) and (rsRelatorios("Data_Fechamento") <> cdate("01/01/1900")) then
							Response.Write rsRelatorios("Data_Fechamento")
						else
							Response.Write "&nbsp;"
						end if
						%>
						</font>
					</td>
	      			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Desc_Eqp")%>&nbsp;</font>
					</td>
   					<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Desc_Comite")%>&nbsp;</font>
					</td>
    	  			<td align="left">
						<font face="Arial" size="1"><%=rsRelatorios("Overall_Status")%>&nbsp;</font>
					</td>

					<%If Trim(strTipo) = """Risks""" Then%>

	      				<td align="left">
							<font face="Arial" size="1"><%=rsRelatorios("Modulo")%>&nbsp;</font>
						</td>
   						<td align="left">
							<font face="Arial" size="1"><%=rsRelatorios("Tipo_Risco")%>&nbsp;</font>
						</td>
    	  				<td align="left">
							<font face="Arial" size="1"><%=rsRelatorios("Dashboard")%>&nbsp;</font>
						</td>

					<%End If%>

    			</tr>
		<% 
		
			rsRelatorios.MoveNext
			
	loop
	%>
  			</table>
	  	</center>
	</div>
<%
end if

'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)

%>
</body>
</html>