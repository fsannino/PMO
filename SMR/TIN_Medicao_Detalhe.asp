<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim strDupla,strSimples
Dim rs
Dim cmdResultado
Dim strCase
Dim X

Dim Cont
Dim intProj_Aux
Dim strLogin

If trim(session("Usuario")) = "" Then
	Response.Redirect ("./TIN_Login.asp")
End if

strDupla = """"
strSimples = "´"


strLogin = trim(session("Login"))

Function VerificaHB()
Dim Dia
Dim Hora

	VerificaHB = true

	Hora = CDate(FormatDateTime(Now(), 3))

	If Hora > CDate("20:00:00") And Hora < CDate("23:00:00") Then
	    VerificaHB = false
	End If

End Function

if VerificaHB then

	strCase	= Request("slcCase")
	
	'Abrindo uma conexão com o BD
	set conConexao = TIN_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_MEDICAO_DETALHE"
        
        .Parameters.Refresh
		.Parameters(1).Value = trim(strCase)

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmMedicao_Detalhe" id="frmMedicao_Detalhe" action="TIN_Medicao_Detalhe.asp" method="post">
	<input type="hidden" id="slcCase" name="slcCase" value="<%=strCase%>">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema de Teste Integrado &gt;<%=rs("PROJ_NAME")%></font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="500px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cenário</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Resposável</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Seq.</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Alt.Dt</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inc.Tar.</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Exc.Tar.</font></b></td>
		  </tr>

		<%Do While Not rs.EOF%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor="#C0C0C0">
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="500px">
					<font face="Arial" size="1">
						<%For X = 1 to rs("TASK_OUTLINE_LEVEL")%>
							&nbsp;
						<%Next%>
						<%=rs("TASK_NAME")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=rs("Cenario")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align=center style="border: 1 solid #666666" width="100px">
					<font face="Arial" size="1">
						<%=rs("RESP")%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align=center style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<% if not isnull(rs("SEQ")) Then%>
							<%=FormatNumber(rs("SEQ"),0)%>&nbsp;
						<% Else%>
							<%=rs("SEQ")%>&nbsp;
						<% End If%>
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
					<%If Not isnull(rs("Flag")) Then%>
						<a href="JavaScript:abreJanelaTesteIntegr('TIN_Alterar_Datas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strLogin=<%=strLogin%>')">
							<img src="icones/ro.gif" alt="Alterar Datas da Tarefa" name="Issues" border="0">
						</a>
					<%Else%>
						&nbsp;
					<%End If%>
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<%If Not isnull(rs("Flag")) Then%>
							<%If Trim(session("Perfil")) = "A" Then%>
								<a href="JavaScript:abreJanelaTesteIntegr('TIN_Incluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=Replace(rs("TASK_NAME"),strDupla,strSimples)%>&strLogin=<%=strLogin%>')">
									<img src="icones/ro.gif" alt="Incluir Tarefa após essa linha" name="Issues" border="0">
								</a>
							<%Else%>
								&nbsp;
							<%End If%>
						<%Else%>
							&nbsp;
						<%End If%>
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						&nbsp;
					</font>
				</td>

			  </tr>
			<%Else%>

				<tr height="17" style="height:12.75pt">
				  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
						<font face="Arial" size="1">
							<%=rs("TASK_UID")%>&nbsp;
						</font>
					</td>

				  <td class="xl28" style="border: 1 solid #666666" width="500px">
						<font face="Arial" size="1">
							<%For X = 1 to rs("TASK_OUTLINE_LEVEL")%>
								&nbsp;
							<%Next%>
							<%=rs("TASK_NAME")%>&nbsp;
						</font>
					</td>

					<td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=rs("Cenario")%>&nbsp;
						</font>
					</td>

				  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
						</font>
				  </td>
				  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
						</font>
				  </td>

				  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
						<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
						</font>
					</td>

					<td class="xl23" align=center style="border: 1 solid #666666" width="100px">
						<font face="Arial" size="1">
							<%=rs("RESP")%>&nbsp;
						</font>
					</td>

					<td class="xl23" align=center style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
						<% if not isnull(rs("SEQ")) Then%>
							<%=FormatNumber(rs("SEQ"),0)%>&nbsp;
						<% Else%>
							<%=rs("SEQ")%>&nbsp;
						<% End If%>
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
						<%If Not isnull(rs("Flag")) Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('TIN_Alterar_Datas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strNome=<%=Replace(rs("TASK_NAME"),strDupla,strSimples)%>&strLogin=<%=strLogin%>')">
								<img src="icones/ro.gif" alt="Alterar Datas da Tarefa" name="Issues" border="0">
							</a>
						<%Else%>
							&nbsp;
						<%End If%>
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
						<%If Trim(session("Perfil")) = "A" Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('TIN_Incluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=Replace(rs("TASK_NAME"),strDupla,strSimples)%>&strLogin=<%=strLogin%>')">
								<img src="icones/ro.gif" alt="Incluir Tarefa após essa linha" name="Issues" border="0">
							</a>
						<%Else%>
							&nbsp;
						<%End If%>
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
						<%If Trim(session("Perfil")) = "A" and isnull(rs("Excluido")) Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('TIN_Excluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=Replace(rs("TASK_NAME"),strDupla,strSimples)%>&strLogin=<%=strLogin%>')">
								<img src="icones/012001.gif" alt="Excluir essa Tarefa" name="Issues" border="0">
							</a>
						<%ElseIf Trim(session("Perfil")) = "A" and not isnull(rs("Excluido")) Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('TIN_Excluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=Replace(rs("TASK_NAME"),strDupla,strSimples)%>&strLogin=<%=strLogin%>')">
								<img src="icones/012002.gif" alt="Cancela Exclusão da Tarefa" name="Issues" border="0">
							</a>
						<%Else%>
							&nbsp;
						<%End If%>
						</font>
					</td>

				</tr>

			<%End If

			rs.MoveNext
			
		  Loop%>
		
		</table>

		<p align="right">
		<BR>
		<img src="img/_0.gif" width="2" height="2">
		<hr>

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)

Else
	response.write "<p><b><center><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>&nbsp;&nbsp;Medição só pode ser atualizada apartir de quinta 06:00 AM</font></center></b></p>"
End If%>

</FORM>
</body>
</html>