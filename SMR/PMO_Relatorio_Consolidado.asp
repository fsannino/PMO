<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<% 
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
strEquipe = Request.Form("slcEquipe")
strRDO = ucase(trim(Request.Form("chkRDO")))
strAcesso = "PUBLIC"
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
set conConexao = SMR_AbrirConexaoBD()

set rsRelatorios = Server.CreateObject("ADODB.RecordSet")

rsRelatorios.Open strSQL, conConexao

%>

<html>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<link rel="stylesheet" href="estilos/sinergia.css">

<table width="100%" border="0">
	<tr>
		<td width="30%" height="10px">&nbsp;</td>
		<td width="40%" align="center" height="10px">
			<p><b><font size="3" face="Georgia, Times New Roman, Times, serif" color="#666666">Sistema de Relatórios</font></b></p>
		</td>
		<td width="30%" height="10px">&nbsp;</td>
	</tr>
</table>

<% 
if not rsRelatorios.EOF then

	strTipoAux = mid(rsRelatorios("ID"),1,3)
	
	response.write "<table width=100% border=0>"
	response.write "	<tr>"
	response.write "		<td width=30% height=10px>&nbsp;</td>"
	response.write "		<td width=40% align=center height=10px>"
	response.write "			<br><b><font face=Georgia, Times New Roman, Times, serif size=2 color=#666666>" & Descricao_Tipo(strTipoAux) & "</font></b><br>"
	response.write "		</td>"
	response.write "		<td width=30% height=10px>&nbsp;</td>"
	response.write "	</tr>"
	response.write "</table>"
	
	do while not rsRelatorios.EOF 

		if mid(rsRelatorios("ID"),1,3) <> strTipoAux then
			strTipoAux = mid(rsRelatorios("ID"),1,3)
			response.write "<table width=100% border=0>"
			response.write "	<tr>"
			response.write "		<td width=30% height=10px>&nbsp;</td>"
			response.write "		<td width=40% align=center height=10px>"
			response.write "			<br><b><font face=Georgia, Times New Roman, Times, serif size=2 color=#666666>" & Descricao_Tipo(strTipoAux) & "</font></b><br>"
			response.write "		</td>"
			response.write "		<td width=30% height=10px>&nbsp;</td>"
			response.write "	</tr>"
			response.write "</table>"
		end if
%>
		<div align="center">
			<center>
				<table border="0" width="90%" cellspacing="0" cellpadding="0" style="border: 1 solid #000000">
					<tr>
						<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">PROJETO</font>
						</td>
						<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">USUÁRIO</font>
						</td>
	    	  			<td width="34%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">ABERTO EM</font>
						</td>
		    		</tr>
				    <tr>
	      				<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Desc_Projeto")%>&nbsp;</font>
						</td>
		      			<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Nome_Funcionario")%>&nbsp;</font>
						</td>
					    <td width="34%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1">
							<%
							if (rsRelatorios("Aberto_em") <> cdate(0)) and (rsRelatorios("Aberto_em") <> cdate("01/01/1900")) then
								response.write rsRelatorios("Aberto_em")
							else
								response.write "&nbsp;"
							end if
							%>
							</font>
						</td>
    				</tr>
	    			<tr>
	    	  			<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">ID</font>
						</td>
	      				<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">ETAPA DE WORKFLOW</font>
						</td>
      					<td width="34%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">DATA LIMITE</font>
						</td>
	    			</tr>
    				<tr>
	    	  			<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("ID")%>&nbsp;</font>
						</td>
      					<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Workflow")%>&nbsp;</font>
						</td>
      					<td width="34%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1">
							<%
							if (rsRelatorios("Data_Limite") <> cdate(0)) and (rsRelatorios("Data_Limite") <> cdate("01/01/1900")) then
								response.write rsRelatorios("Data_Limite")
							else
								response.write "&nbsp;"
							end if
							%>
							</font>
						</td>
    				</tr>
	    			<tr>
    	  				<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">NOME</font>
						</td>
    	  				<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">PRIORIDADE</font>
						</td>
    		  			<td width="34%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">DATA DE FECHAMENTO</font>
						</td>
		    		</tr>
		    		<tr>
    		  			<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Name")%>&nbsp;</font>
						</td>
	    	  			<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Priority")%>&nbsp;</font>
						</td>
	    	  			<td width="34%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1">
							<%
							if (rsRelatorios("Data_Fechamento") <> cdate(0)) and (rsRelatorios("Data_Fechamento") <> cdate("01/01/1900")) then
								response.write rsRelatorios("Data_Fechamento")
							else
								response.write "&nbsp;"
							end if
							%>
							</font>
						</td>
	    			</tr>
		    		<tr>
    		  			<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">EQUIPE</font>
						</td>
		      			<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">COMITÊ</font>
						</td>
		      			<td width="34%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">STATUS</font>
						</td>
		    		</tr>
    				<tr>
		      			<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Desc_Eqp")%>&nbsp;</font>
						</td>
      					<td width="33%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Desc_Comite")%>&nbsp;</font>
						</td>
	    	  			<td width="34%" align="center" style="border: 1 solid #000000">
							<font face="Arial" size="1"><%=rsRelatorios("Overall_Status")%>&nbsp;</font>
						</td>
	    			</tr>

					<%If Trim(strTipo) = """Risks""" Then%>
		    			<tr>
    		  				<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
								<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">RDO</font>
							</td>
		      				<td width="33%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
								<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">TIPO DE RISCO</font>
							</td>
		      				<td width="34%" align="center" bgcolor="#639ACE" style="border: 1 solid #000000">
								<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">DASHBOARD</font>
							</td>
		    			</tr>
    					<tr>
		      				<td width="33%" align="center" style="border: 1 solid #000000">
								<font face="Arial" size="1"><%=rsRelatorios("Modulo")%>&nbsp;</font>
							</td>
      						<td width="33%" align="center" style="border: 1 solid #000000">
								<font face="Arial" size="1"><%=rsRelatorios("Tipo_Risco")%>&nbsp;</font>
							</td>
	    	  				<td width="34%" align="center" style="border: 1 solid #000000">
								<font face="Arial" size="1"><%=rsRelatorios("Dashboard")%>&nbsp;</font>
							</td>
	    				</tr>
					<%End If%>

	  			</table>
		  	</center>
		</div>
		<hr size="3" color="#000000">
		
		<% rsRelatorios.movenext
		
	loop

else

	response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
	
end if

'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)

%>
</body>
</html>
