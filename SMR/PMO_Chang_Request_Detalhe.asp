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
        .CommandText = "SP_INCLUIR_CHANG_REQUEST_DETALHE_TB_TEMP"
        
        .Parameters.Refresh
		.Parameters(1).Value = trim(session("Usuario"))

    End With

	cmdResultado.Execute()


'********************************************************************************

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_SUMARIAS"

		.Parameters(1).Value = trim(session("Usuario"))
		.Parameters(2).Value = "R"
        
        .Parameters.Refresh

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

		.Parameters(1).Value = trim(session("Usuario"))
		.Parameters(2).Value = "R"
	
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
					.Parameters(16).Value = Null
					.Parameters(17).Value = Null
					.Parameters(18).Value = Null
					.Parameters(19).Value = Null
					.Parameters(20).Value = Null
					.Parameters(21).Value = Null
					.Parameters(22).Value = Null
					.Parameters(23).Value = Null
					.Parameters(24).Value = Null
					.Parameters(25).Value = trim(session("Usuario"))
					.Parameters(26).Value = "R"


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

'*********************************************************************************

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_CHANG_REQUEST_DETALHE"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Usuario"))
		.Parameters(2).Value = "R"

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
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Chang Requests &gt;<%=rs("PROJ_NAME")%>&nbsp;&nbsp;&nbsp;(Alteração/Exclusão de Atividades)</font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="700px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasada</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">C.R</font></b></td>

		  </tr>

		<%Cont = 0%>
		<%intProj_Aux = rs("PROJ_ID")%>
		<%Do While Not rs.EOF%>

			<%If rs("PROJ_ID") <> intProj_Aux Then
				
				intProj_Aux = rs("PROJ_ID")%>
				<BR>
				<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Chang Requests &gt;<%=rs("PROJ_NAME")%>&nbsp;&nbsp;&nbsp;(Alteração/Exclusão de Atividades)</font></b>
				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
					<tr height="17" style="height:12.75pt">
						<td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
						<td class="xl27" width="700px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
						<td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasada</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">C.R</font></b></td>

					
					</tr>
				  
			<%End If%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor=#ebebeb>
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="700px">
					<font face="Arial" size="1">
						<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
						<%=left("                              ",rs("task_outline_level")) %>
						<%=rs("TASK_NAME")%>
						</div>
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;
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

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						&nbsp;
					</font>
				</td>


			  </tr>
			  
			<%Else%>

					<input type="hidden" id="hidProj" name="hidProj" value="<%=rs("PROJ_ID")%>">
					<input type="hidden" id="hidId" name="hidId" value="<%=rs("TASK_ID")%>">
					<input type="hidden" id="hidUid" name="hidUid" value="<%=rs("TASK_UID")%>">
					<input type="hidden" id="hidNome" name="hidNome" value="<%=rs("TASK_NAME")%>">
					<input type="hidden" id="hidDtInicio" name="hidDtInicio" value="<%=rs("TASK_START_DATE")%>">
					<input type="hidden" id="hidDtFim" name="hidDtFim" value="<%=rs("TASK_FINISH_DATE")%>">
					<input type="hidden" id="hidPerc" name="hidPerc" value="<%=rs("TASK_PCT_COMP")%>">
					<input type="hidden" id="hidDtLimite" name="hidDtLimite" value="<%=rs("Data_Limite")%>">

					<tr height="17" style="height:12.75pt">
					  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
							<font face="Arial" size="1">
								<%=rs("TASK_UID")%>&nbsp;
							</font>
						</td>

					  <td class="xl28" style="border: 1 solid #666666" width="700px">
							<font face="Arial" size="1">
								<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
								<%=left("                              ",rs("task_outline_level")) %>
								<%=rs("TASK_NAME")%>
								</div>
							</font>
						</td>

					<td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=rs("DURACAO")%>&nbsp;
						</font>
					</td>

					  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
							<font face="Arial" size="1">
								<input type="hidden" id="hiddtStart" name="hiddtStart" value="<%=rs("TASK_START_DATE")%>">
								<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
							</font>
					  </td>
					  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
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
								<INPUT type="checkbox" id=chkChangRequest name=chkChangRequest>
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
		<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a>
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
	document.frmChang_Request_Detalhe_PMO.action = "PMO_Chang_Request_Lista.asp";
	document.frmChang_Request_Detalhe_PMO.submit();

}

function Confirmar()
{
	var intCont	= (document.frmChang_Request_Detalhe_PMO.hidCont.value - 1)
	var strAux = ""

	document.frmChang_Request_Detalhe_PMO.style.cursor = "wait";
	document.frmChang_Request_Detalhe_PMO.hidDados.value = "";
	document.frmChang_Request_Detalhe_PMO.hidOperacao.value = "";
	
	for(var i = 0; i <= intCont;i++)
	{

		if (document.frmChang_Request_Detalhe_PMO.chkChangRequest(i).checked == true)
		{
			if (strAux == "")
			{
				strAux = document.frmChang_Request_Detalhe_PMO.hidProj(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidUid(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidNome(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtInicio(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtFim(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidPerc(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtLimite(i).value;
			}	
			else
			{
				strAux = strAux + ";" + document.frmChang_Request_Detalhe_PMO.hidProj(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidUid(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidNome(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtInicio(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtFim(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidPerc(i).value + "|" + 
				document.frmChang_Request_Detalhe_PMO.hidDtLimite(i).value;

			}
		}
	}

	if (strAux != "")
	{
		document.frmChang_Request_Detalhe_PMO.hidDados.value = strAux;
		document.frmChang_Request_Detalhe_PMO.hidOperacao.value = 'C';
		document.frmChang_Request_Detalhe_PMO.style.cursor = "";
		document.frmChang_Request_Detalhe_PMO.action = "PMO_Chang_Request.asp";
		document.frmChang_Request_Detalhe_PMO.submit();
	}
	else
	{
		alert("Nenhum registro foi alterado");
		document.frmChang_Request_Detalhe_PMO.style.cursor = "";
	}

}
</SCRIPT>
</FORM>
</body>
</html>
