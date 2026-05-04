<!--#include file="funcoes/Funcoes.inc"-->
<%
'Ordena ao browser a abertura do MS-Excel
Response.ContentType = "application/vnd.ms-excel"

Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs
Dim strFrente
Dim strEquipe
Dim strGovernaca
Dim strArea
Dim strUnidade
Dim strUsuarioCLI
Dim strCompleto
Dim strDtInicio

Dim Cont
Dim intProj_Aux

	strFrente		= Request("slcFrente")
	strEquipe		= Request("slcEquipe")
	strGovernaca	= Request("slcGovernaca")
	strArea			= Request("slcArea")
	strUnidade		= Request("slcUnidade")
	strUsuarioCLI	= Request("slcUsuarioCLI")
	strCompleto		= Request("slcCompleto")
	strDtInicio		= Request("slcDataInicio")

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_INCLUIR_REL_MEDICAO_DETALHE_TB_TEMP"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Usuario"))

       	If trim(strUsuarioCLI) <> "" then
			.Parameters(2).Value = trim(strUsuarioCLI)
		Else
			.Parameters(2).Value = Null
		End if
		
       	If trim(strFrente) <> "" then
			.Parameters(3).Value = trim(strFrente)
		Else
			.Parameters(3).Value = Null
		End if
       
       	If trim(strEquipe) <> "" then
			.Parameters(4).Value = trim(strEquipe)
		Else
			.Parameters(4).Value = Null
		End if

      	If trim(strGovernaca) <> "" then
			.Parameters(5).Value = trim(strGovernaca)
		Else
			.Parameters(5).Value = Null
		End if

       	If trim(strUnidade) <> "" then
			.Parameters(6).Value = trim(strUnidade)
		Else
			.Parameters(6).Value = Null
		End if

       	If trim(strArea) <> "" then
			.Parameters(7).Value = trim(strArea)
		Else
			.Parameters(7).Value = Null
		End if

       	If trim(strCompleto) <> "" then
			.Parameters(8).Value = trim(strCompleto)
		Else
			.Parameters(8).Value = Null
		End if

       	If trim(strDtInicio) <> "" then
			.Parameters(9).Value = trim(strDtInicio)
		Else
			.Parameters(9).Value = Null
		End if
        
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
		.Parameters(2).Value = "C"
	
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
					.Parameters(23).Value = null
					.Parameters(24).Value = null
					.Parameters(25).Value = trim(session("Usuario"))
					.Parameters(26).Value = "C"


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
        .CommandText = "SP_LISTAR_RELATORIO_MEDICAO_DETALHE"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Usuario"))
		.Parameters(2).Value = "C"

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
	<FORM name="frmRelatorio_Medicao_Detalhe" id="frmRelatorio_Medicao_Detalhe" action="Relatorio_Medicao_Detalhe.asp" method="post">
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe Sinergia &gt;<%=rs("PROJ_NAME")%></font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=30px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="450px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
		    <td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipe</font></b></td>

		    <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável</font></b></td>
		  </tr>

		<%Cont = 0%>
		<%intProj_Aux = rs("PROJ_ID")%>
		<%Do While Not rs.EOF%>

			<%If rs("PROJ_ID") <> intProj_Aux Then
				
				intProj_Aux = rs("PROJ_ID")%>
				<BR>
				<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe Sinergia &gt;<%=rs("PROJ_NAME")%></font></b>
				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
					<tr height="17" style="height:12.75pt">
					  <td height="17" class="xl27" width=30px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
					  <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
					  <td class="xl27" width="450px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
					  <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
					  <td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
					  <td class="xl27" width="65px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
					  <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipe</font></b></td>

					  <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável</font></b></td>
					</tr>
				  
			<%End If%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor="#C0C0C0">
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						
						<%If rs("PROJ_ID") = 2 And rs("TASK_PCT_COMP") = 0 Then%>
							&nbsp;
						<%Else%>
							<%=rs("TASK_PCT_COMP")%>
						<%End if%>
						
						<%'=rs("TASK_PCT_COMP")%>
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="450px">
					<font face="Arial" size="1">

						<%For X = 1 to rs("TASK_OUTLINE_LEVEL")%>
							&nbsp;
						<%Next%>
						<%=rs("TASK_NAME")%>

					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>
					</font>
				</td>

			    <td class="xl30" style="border: 1 solid #666666" width="100px" align=center>
					<font face="Arial" size="1"><%=rs("Equipe")%></font>
				</td>

			    <td class="xl22" style="border: 1 solid #666666" width="200px" align=center>
					<font face="Arial" size="1"><%=rs("CLI")%></font>
				</td>
			  </tr>
			<%Else%>

				<%'If rs("TASK_IS_EXTERNAL") = True Then%>

					<tr height="17" style="height:12.75pt">
					  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
							<font face="Arial" size="1">
								<%=rs("TASK_UID")%>
							</font>
						</td>

					  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
							<font face="Arial" size="1">
								<%If rs("PROJ_ID") = 2 And rs("TASK_PCT_COMP") = 0 Then%>
									&nbsp;
								<%Else%>
									<%=rs("TASK_PCT_COMP")%>
								<%End if%>
								
								<%'=rs("TASK_PCT_COMP")%>
							</font>
						</td>

					  <td class="xl28" style="border: 1 solid #666666" width="450px">
							<font face="Arial" size="1">

								<%For X = 1 to rs("TASK_OUTLINE_LEVEL")%>
									&nbsp;
								<%Next%>
								<%=rs("TASK_NAME")%>

							</font>
						</td>

					  <td class="xl23" align="right" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
							</font>
						</td>

					  <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
							<font face="Arial" size="1">
								<%=FormatarDataMon(rs("TASK_START_DATE"))%>
							</font>
						</td>
					  <td class="xl30" align="right" style="border: 1 solid #666666" width="65px">
							<font face="Arial" size="1">
								<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>
							</font>
						</td>

						<td class="xl30" style="border: 1 solid #666666" width="100px" align=center>
							<font face="Arial" size="1"><%=rs("Equipe")%></font>
						</td>

					  <td class="xl22" style="border: 1 solid #666666" width="200px" align=center>
							<font face="Arial" size="1"><%=rs("CLI")%></font>
						</td>

					</tr>

<%
				'End If

			End If

			rs.MoveNext
			
			If Not rs.EOF Then
				If (rs("PROJ_ID") <> intProj_Aux) Then%>
					</table>
				<%End If
			End If%>
			
		<%Loop%>
		
		</table>

		<p align="right">
		<!--<input type="button" name="cmdSubmit" value="Enviar" onclick="Confirmar();">-->

	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)%>

</FORM>
</body>
</html>