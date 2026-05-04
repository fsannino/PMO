<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 5000

Dim rs
Dim rs1

Dim cmdResultado

Dim Cont
Dim intProj_Aux
Dim strIntegrador_Aux

Dim strAtualizacao
Dim strVisualizacao

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 240
        .CommandText = "SP_LISTAR_TAREFAS_ISSUES_GVI"

        .Parameters.Refresh
		
    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Issues" id="frmRelatorio_Issues" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<%If Not rs.EOF Then%>

		<table cellspacing="0" cellpadding="0" align=center>
			<tr height="17"  >
				<td align=left width="600px">
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Relatório de Problemas &gt;<%=rs("PROJ_NAME")%></font></b>
				</td>
			</tr>
		</table>

		<table cellspacing="0" cellpadding="0" align=center>
			<tr height="17"  >
				<td align=left width="600px">
					<b><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Integrador &gt;<%=rs("Integrador")%></font></b>
				</td>
			</tr>
		</table>

		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
			<tr height="17" style="height:12.75pt" >

			<td class="xl27" width="500px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Título do Problema</font></b></td>

			<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atualizar Problema</font></b></td>

			<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Visualizar Problema</font></b></td>

			</tr>



		<%Cont = 0%>
		<%
		intProj_Aux = rs("PROJ_ID")
		strIntegrador_Aux = rs("Integrador")
		%>

		<%Do While Not rs.EOF%>

			<%If rs("PROJ_ID") <> intProj_Aux Then
				
				intProj_Aux = rs("PROJ_ID")				
				strIntegrador_Aux = rs("Integrador")%>
				<table cellspacing="0" cellpadding="0" align=center>
					<tr height="17"  >
						<td align=left width="600px">
							<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Relatório de Problemas &gt;<%=rs("PROJ_NAME")%></font></b>
						</td>
					</tr>
				</table>

				<table cellspacing="0" cellpadding="0" align=center>
					<tr height="17"  >
						<td align=left width="600px">
							<b><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Integrador &gt;<%=rs("Integrador")%></font></b>
						</td>
					</tr>
				</table>


				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					<tr height="17" style="height:12.75pt" >

					<td class="xl27" width="500px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Título do Problema</font></b></td>

					<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atualizar Problema</font></b></td>

					<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Visualizar Problema</font></b></td>

					</tr>


			<%End If%>


			<%If rs("Integrador") <> strIntegrador_Aux Then
				
				strIntegrador_Aux = rs("Integrador")%>
				<table cellspacing="0" cellpadding="0" align=center>
					<tr height="17"  >
						<td align=left width="600px">
							<b><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Integrador &gt;<%=rs("Integrador")%></font></b>
						</td>
					</tr>
				</table>


				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					<tr height="17" style="height:12.75pt" >

					<td class="xl27" width="500px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Título do Problema</font></b></td>

					<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atualizar Problema</font></b></td>

					<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Visualizar Problema</font></b></td>

					</tr>

			<%End If%>

				<%				
				Set cmdResultado = Server.CreateObject("ADODB.Command")
    
				With cmdResultado
    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 240
				    .CommandText = "SP_LISTAR_ISSUES_GVI"

				    .Parameters.Refresh
					.Parameters(1).Value = rs("PROJ_ID")
					.Parameters(2).Value = rs("TASK_UID")
						
				End With

				set rs1 = Server.CreateObject("ADODB.RecordSet")

				set rs1 = cmdResultado.Execute()
				%>

				<%Do While Not rs1.EOF%>
					<tr height="17" style="height:12.75pt" >

					  <td class="xl28" style="border: 1 solid #666666" width="500px" nowrap>
							<font face="Arial" size="1">
								<%=rs1("Nome_Issue")%>&nbsp;
							</font>
					  </td>

					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%
									strAtualizacao = "GVI_Comentario_Issues.asp?strCodIssue=" & trim(rs1("Cod_Issue")) & "&strProjeto=" & trim(rs1("PROJ_ID")) & "&strUID=" & trim(rs1("UID")) & "&strIntegrador=" & trim(rs("Integrador")) & "&strTIPO=A"
									strVisualizacao = "GVI_Comentario_Issues.asp?strCodIssue=" & trim(rs1("Cod_Issue")) & "&strProjeto=" & trim(rs1("PROJ_ID")) & "&strUID=" & trim(rs1("UID")) & "&strIntegrador=" & trim(rs("Integrador")) & "&strTIPO=V"
								%>
								<a href="JavaScript:Carregar_Pagina('<%=strAtualizacao%>')">
									<img src="icones/ro.gif" alt="Atualizar Problema" name="Issues" border="0">
								</a>
							</font>
					  </td>

					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<a href="JavaScript:Carregar_Pagina('<%=strVisualizacao%>')">
									<img src="icones/detalhar.gif" alt="Visualizar Problema" name="Issues" border="0">
								</a>
							</font>
					  </td>

					</tr>

				    <%rs1.MoveNext%>
				<%Loop%>

			<%rs.MoveNext%>

			<%If Not rs.EOF Then%>
				<%If rs("Integrador") <> strIntegrador_Aux Then%>
					</table>
					<BR>
				<%End If%>
			<%End If%>			

		<%Loop%>

		<p align="right">
		<input type="hidden" id="hidCont" name="hidCont" value="<%=Cont%>">

		<input type="hidden" id="hidDados" name="hidDados" value="">
		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">

	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para exibição</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)%>


<SCRIPT Language="JavaScript">
	function Carregar_Pagina(strPagina)
	{
		frmRelatorio_Issues.action = strPagina ;
		frmRelatorio_Issues.submit();
	}
</SCRIPT>

</FORM>
</body>
</html>