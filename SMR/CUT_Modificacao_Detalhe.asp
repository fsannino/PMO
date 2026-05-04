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
Dim strPlano
Dim strModulo
Dim strWBS

Dim Cont
Dim intProj_Aux
Dim strLogin

strDupla = """"
strSimples = "´"

	If trim(session("Login")) = "" Then
'		Response.Redirect ("./Erro.asp?Erro=Sua sessão expirou. Por-favor, logue-se novamente.&Voltar=true&IrPara=./CUT_LOGIN.ASP")
		Response.Redirect ("./Erro.asp?Erro=Sua sessão expirou. Por-favor, logue-se novamente.&Voltar=true&IrPara=./CUT_LOGIN.ASP?hidOrigem=./CUT_Modificacao_Filtro.asp")
	Else
		strLogin = trim(session("Login"))
	End if

	strPlano  = Trim(Request("slcPlano"))
	strModulo = Trim(Request("slcModulo"))

	If strModulo <> "" Then
		strWBS = strModulo
	Else
		strWBS = strPlano
	End If

	'Abrindo uma conexão com o BD
	set conConexao = CUT_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_INCLUIR_MODIFICACAO_DETALHE_TB_TEMP"
        
        .Parameters.Refresh

		.Parameters(1).Value = strLogin
		.Parameters(2).Value = strWBS
        
    End With

	cmdResultado.Execute()


'********************************************************************************

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_SUMARIAS"
        
        .Parameters.Refresh

		.Parameters(1).Value = strLogin
		.Parameters(2).Value = "A"

    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_TB_TEMP"
       
        .Parameters.Refresh

		.Parameters(1).Value = strLogin
		.Parameters(2).Value = "A"
	
    End With

	set rs2 = Server.CreateObject("ADODB.RecordSet")

	set rs2 = cmdResultado.Execute()

	Do While Not rs2.EOF and Not rs1.EOF 

		IF rs1("PROJ_ID") = rs2("PROJ_ID") Then

			IF trim(rs1("Task_Outline_num") & ".") = Left(rs2("Task_Outline_num"),LEN(trim(rs1("Task_Outline_num")))+1) Then

				Set cmdResultado = Server.CreateObject("ADODB.Command")
				    
				With cmdResultado
    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 600
				    .CommandText = "SP_INCLUIR_TB_TEMP"
				    
				    .Parameters.Refresh
					.Parameters(1).Value = rs1("TASK_OUTLINE_NUM")
					.Parameters(2).Value = rs1("TASK_NAME")
					.Parameters(3).Value = rs1("TASK_OUTLINE_LEVEL")
					.Parameters(4).Value = rs1("TASK_BASE_START")
					.Parameters(5).Value = rs1("TASK_BASE_FINISH")
					.Parameters(6).Value = rs1("TASK_START_DATE")
					.Parameters(7).Value = rs1("TASK_FINISH_DATE")
					.Parameters(8).Value = rs1("TASK_PCT_COMP")
					.Parameters(9).Value = rs1("TASK_IS_SUMMARY")
					.Parameters(10).Value = rs1("TASK_IS_EXTERNAL")
					.Parameters(11).Value = rs1("PROJ_ID")
					.Parameters(12).Value = rs1("PROJ_NAME")
					.Parameters(13).Value = rs1("TASK_UID")
					.Parameters(14).Value = rs1("TASK_ID")
					.Parameters(15).Value = rs1("TASK_DUR")
					.Parameters(16).Value = rs1("CLI")
					.Parameters(17).Value = rs1("Frente")
					.Parameters(18).Value = rs1("Equipe")
					.Parameters(19).Value = rs1("Governanca")
					.Parameters(20).Value = rs1("RespInt")
					.Parameters(21).Value = rs1("Area")
					.Parameters(22).Value = rs1("Unidade")
					.Parameters(23).Value = Null
					.Parameters(24).Value = Null
					.Parameters(25).Value = strLogin
					.Parameters(26).Value = "A"

				End With

				cmdResultado.Execute()

				rs1.MoveNext
		
			ElseIF trim(rs1("Task_Outline_num") & ".") < Left(rs2("Task_Outline_num"),LEN(trim(rs1("Task_Outline_num")))+1) Then

				rs1.MoveNext	
			Else 

				rs2.MoveNext	
			End If

		ElseIF rs1("PROJ_ID") > rs2("PROJ_ID") Then

			rs2.MoveNext

		Else 

			rs1.MoveNext

		End If
				
	Loop

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_MODIFICACAO_DETALHE"
        
        .Parameters.Refresh

		.Parameters(1).Value = strLogin
		.Parameters(2).Value = "A"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

'*********************************************************************************

		
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmModificacao_Detalhe" id="frmModificacao_Detalhe" action="CUT_Modificacao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe Sinergia &gt;<%=rs("PROJ_NAME")%></font></b>

<!--		<table>
			<tr  align=right>
				<td  width=725px>
					&nbsp;
				</td>
				<td  align=right>
					<a href="JavaScript:abreJanelaTesteIntegr('CUT_Listar_Alterar_Datas.asp?strLogin=<%=strLogin%>')">
						Listar Alterações Solicitadas Anteriomente
					</a>
				</td>
			</tr>
		</table>-->

		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
					  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="400px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Resposável</font></b></td>
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

			    <td class="xl28" style="border: 1 solid #666666" width="400px">
					<font face="Arial" size="1">
						<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
						<%=left("                              ",rs("task_outline_level")) %>
						<%=rs("TASK_NAME")%>
						</div>
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align=center style="border: 1 solid #666666" width="200px">
					<font face="Arial" size="1">
						<%=rs("CLI")%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						&nbsp;
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<%If Not isnull(rs("Flag")) Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('CUT_Incluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=ReplicaPlics(rs("TASK_NAME"))%>&strLogin=<%=strLogin%>')">
								<img src="icones/ro.gif" alt="Incluir Tarefa após essa linha" name="Issues" border="0">
							</a>
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

				  <td class="xl28" style="border: 1 solid #666666" width="400px">
						<font face="Arial" size="1">
							<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
							<%=left("                              ",rs("task_outline_level")) %>
							<%=rs("TASK_NAME")%>
							</div>
						</font>
					</td>

				  <td class="xl30" align=center style="border: 1 solid #666666" width="85px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
						</font>
				  </td>
				  <td class="xl30" align=center style="border: 1 solid #666666" width="85px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
						</font>
				  </td>

				  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
						<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
						</font>
					</td>

					<td class="xl23" align=center style="border: 1 solid #666666" width="200px">
						<font face="Arial" size="1">
							<%=rs("CLI")%>&nbsp;
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
							<a href="JavaScript:abreJanelaTesteIntegr('CUT_Alterar_Datas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=ReplicaPlics(rs("TASK_NAME"))%>&strLogin=<%=strLogin%>')">
								<img src="icones/ro.gif" alt="Alterar Datas da Tarefa" name="Issues" border="0">
							</a>
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
							<a href="JavaScript:abreJanelaTesteIntegr('CUT_Incluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=ReplicaPlics(rs("TASK_NAME"))%>&strLogin=<%=strLogin%>')">
								<img src="icones/ro.gif" alt="Incluir Tarefa após essa linha" name="Issues" border="0">
							</a>
						</font>
					</td>

					<td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
						<%If isnull(rs("Excluido")) Then%>
							<a href="JavaScript:abreJanelaTesteIntegr('CUT_Excluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=ReplicaPlics(rs("TASK_NAME"))%>&strLogin=<%=strLogin%>&strPlano=<%=strPlano%>&strModulo=<%=strModulo%>')">
								<img src="icones/012001.gif" alt="Excluir essa Tarefa" name="Issues" border="0">
							</a>
						<%Else%>
							<a href="JavaScript:abreJanelaTesteIntegr('CUT_Excluir_Tarefas.asp?strProjeto=<%=rs("PROJ_ID")%>&strUID=<%=rs("TASK_UID")%>&strDataIni=<%=rs("TASK_START_DATE")%>&strDataFim=<%=rs("TASK_FINISH_DATE")%>&strID=<%=rs("TASK_ID")%>&strNome=<%=ReplicaPlics(rs("TASK_NAME"))%>&strLogin=<%=strLogin%>&strPlano=<%=strPlano%>&strModulo=<%=strModulo%>')">
								<img src="icones/012002.gif" alt="Cancela Exclusão da Tarefa" name="Issues" border="0">
							</a>
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

%>

	<input type="hidden" id="slcPlano" name="slcPlano" value="<%=strPlano%>">
	<input type="hidden" id="slcModulo" name="slcModulo" value="<%=strModulo%>">

</FORM>
</body>
</html>