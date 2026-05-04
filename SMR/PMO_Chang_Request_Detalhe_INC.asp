<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs
Dim rs1
Dim rs2

Dim cmdResultado
Dim strFrente
Dim strEquipe
Dim vetDados
Dim vetDados_Disc
Dim strProj
Dim strID
Dim strUID
Dim strComplete
Dim strSql
Dim Cont
Dim intProj_Aux


	If trim(session("Usuario")) = "" Then
		response.Redirect("./LOGIN.ASP?hidOrigem=./PMO_Chang_Request_Selecao.asp")
	End if
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_CHANG_REQUEST_DETALHE_INC"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmChang_Request_Detalhe_PMO" id="frmChang_Request_Detalhe_PMO" action="PMO_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Chang Requests &gt;<%=rs("PROJ_NAME")%>&nbsp;&nbsp;&nbsp;(Adição de Novas Atividades)</font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="650px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
		    <td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasada</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Incluir C.R</font></b></td>

		  </tr>

		<%Cont = 0%>
		<%intProj_Aux = rs("PROJ_ID")%>
		<%Do While Not rs.EOF%>

			<%If rs("PROJ_ID") <> intProj_Aux Then
				
				intProj_Aux = rs("PROJ_ID")%>
				<BR>
				<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Chang Requests &gt;<%=rs("PROJ_NAME")%>&nbsp;&nbsp;&nbsp;(Adição de Novas Atividades)</font></b>
				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
					<tr height="17" style="height:12.75pt">
						<td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
						<td class="xl27" width="650px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
						<td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
						<td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
						<td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
						<td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasada</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Incluir C.R</font></b></td>

					
					</tr>
				  
			<%End If%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor=#ebebeb>
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="650px">
					<font face="Arial" size="1">
						<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
						<%=left("                              ",rs("task_outline_level")) %>
						<%=rs("TASK_NAME")%>&nbsp;
						</div>
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;<%IF rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						&nbsp;
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<a href="JavaScript:abreJanelaChangRequests('PMO_Chang_Request_INC.asp?strUsuario=<%=trim(session("Usuario"))%>&strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDtInicio=<%=rs("TASK_START_DATE")%>&strDtFim=<%=rs("TASK_FINISH_DATE")%>&strPerc=<%=rs("TASK_PCT_COMP")%>&strWBS=<%=rs("TASK_OUTLINE_NUM")%>')">
							<img src="icones/ro.gif" alt="Incluir tarefa abaixo dessa linha" name="Issues" border="0">
						</a>
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

					  <td class="xl28" style="border: 1 solid #666666" width="650px">
							<font face="Arial" size="1">
								<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
								<%=left("                              ",rs("task_outline_level")) %>
								<%=rs("TASK_NAME")%>&nbsp;
								</div>
							</font>
						</td>

					<td class="xl30" align=center style="border: 1 solid #666666" width="85px">
						<font face="Arial" size="1">
							<%=rs("DURACAO")%>&nbsp;<%IF rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
						</font>
					</td>

					  <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
							<font face="Arial" size="1">
								<input type="hidden" id="hiddtStart" name="hiddtStart" value="<%=rs("TASK_START_DATE")%>">
								<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
							</font>
					  </td>
					  <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
							<font face="Arial" size="1">
								<input type="hidden" id="hiddtFinish" name="hiddtFinish" value="<%=rs("TASK_FINISH_DATE")%>">
								<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
							</font>
					  </td>

					  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
							<font face="Arial" size="1">
								<%=rs("TASK_PCT_COMP")%>&nbsp;
							</font>
						</td>

						<td class="xl30" align=center style="border: 1 solid #666666" width="75px">
							<font face="Arial" size="1">
								<%IF rs("TASK_PCT_COMP") < 100 Then%>
									<%If cdate(rs("TASK_FINISH_DATE")) < date() Then%>
										<img src="icones/Vermelho.gif" alt="Atrasada" name="Atrasada" border="0" width=18>
									<%Else%>
										<img src="icones/Verde.gif" alt="No Prazo" name="Atrasada" border="0" width=18>
									<%End If%>
								<%Else%>
									&nbsp;
								<%End If%>
							</font>
						</td>

						<td class="xl30" align=center style="border: 1 solid #666666" width="75px">
							<font face="Arial" size="1">
								<a href="JavaScript:abreJanelaChangRequests('PMO_Chang_Request_INC.asp?strUsuario=<%=trim(session("Usuario"))%>&strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDtInicio=<%=rs("TASK_START_DATE")%>&strDtFim=<%=rs("TASK_FINISH_DATE")%>&strPerc=<%=rs("TASK_PCT_COMP")%>&strWBS=<%=rs("TASK_OUTLINE_NUM")%>')">
									<img src="icones/ro.gif" alt="Incluir tarefa abaixo dessa linha" name="Issues" border="0">
								</a>
							</font>
						</td>


					</tr>

					<%Cont = Cont + 1%>

			<%End If

			rs.MoveNext
			
			If Not rs.EOF Then
				If (rs("PROJ_ID") <> intProj_Aux) Then%>
					</table>
				<%End If
			End If%>
			
		<%Loop%>
		
		</table>

		<p align="right">
		<!--<input type="Image" name="cmdConfirmar" value="Confirmar" src="img/000049.gif" align="absmiddle" onclick="Confirmar();">-->
		<!--<input type="button" name="cmdConfirmar" value="Confirmar" onclick="Confirmar();">-->
		<!--<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a>-->
		<BR>
		<table cellspacing="0" cellpadding="0" align=center>
			<tr align=center>
				<td></td>
				<td align=center><a href="javascript:ListarAdicionadas();" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Listar Tarefas Adicionadas</font></a></td>
				<td></td>
			</tr>
		</table>
		<img src="img/_0.gif" width="2" height="2">
		<hr>
		<input type="hidden" id="hidCont" name="hidCont" value="<%=Cont%>">
		<input type="hidden" id="hidDados" name="hidDados" value="">
		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
		
		<input type="hidden" id="strUsuario" name="strUsuario" value="<%=trim(session("Usuario"))%>">

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)
%>

<SCRIPT language=JavaScript>

function ListarAdicionadas()
{
	document.frmChang_Request_Detalhe_PMO.action = "PMO_Chang_Request_Lista_INC.asp";
	document.frmChang_Request_Detalhe_PMO.submit();

}
</SCRIPT>

</FORM>
</body>
</html>
