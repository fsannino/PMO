<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim strProj
Dim strFrente
Dim strEquipe
Dim strResponsavel
Dim strOpcao

Dim cmdResultado
Dim rs


Function RetornaCriticidade(strCriticidade)
Dim Valor

	RetornaCriticidade = ""

	If Not IsNull(strCriticidade) Then
		Valor = Cint(strCriticidade)
	
		If  Valor >=1 AND Valor <= 4 Then
			RetornaCriticidade = "icones/Verde.gif"
		ElseIf  Valor >=5 AND Valor <= 13 Then
			RetornaCriticidade = "icones/Amarelo.gif"
		ElseIf  Valor >=14 AND Valor <= 21 Then
			RetornaCriticidade = "icones/Vermelho.gif"
		ElseIf  Valor >=22 AND Valor <= 100 Then
			RetornaCriticidade = "icones/Preto.gif"
		End If
	End If
End Function

	strProj   = Trim(Request("slcProjeto"))
	strFrente = Trim(Request("slcFrente"))
	strEquipe = Trim(Request("slcEquipe"))
	strResponsavel = Trim(Request("slcResponsavel"))
	strOpcao = Trim(Request("slcOpcao"))


'	Response.Write "strProj - " & strProj & "<BR>"
'	Response.Write "strFrente - " & strFrente & "<BR>"
'	Response.Write "strEquipe - " & strEquipe & "<BR>"
'	Response.Write "strResponsavel - " & strResponsavel & "<BR>"
'	Response.Write "strOpcao - " & strOpcao & "<BR>"

	If trim(session("Login")) = "" Then
		session("Login") = "PAINEL"
	End If


	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")

    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_INCLUIR_REL_PAINEL_SINOPTICO_TB_TEMP"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Login"))

       	If trim(strProj) <> "" then
			.Parameters(2).Value = trim(strProj)
		Else
			.Parameters(2).Value = Null		
		End If

       	If trim(strFrente) <> "" then
			.Parameters(3).Value = trim(strFrente)
		Else
			.Parameters(3).Value = Null
		End If
		       
       	If trim(strEquipe) <> "" then
			.Parameters(4).Value = trim(strEquipe)
		Else
			.Parameters(4).Value = Null
		End if

       	If trim(strResponsavel) <> "" then
			.Parameters(5).Value = trim(strResponsavel)
		Else
			.Parameters(5).Value = Null
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

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = "C"

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

		.Parameters(1).Value = trim(session("Login"))
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
					.Parameters(17).Value = Null
					.Parameters(18).Value = rs1("Equipe")
					.Parameters(19).Value = Null
					.Parameters(20).Value = Null
					.Parameters(21).Value = Null
					.Parameters(22).Value = Null
					.Parameters(23).Value = Null
					.Parameters(24).Value = Null
					.Parameters(25).Value = trim(session("Login"))
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
        .CommandText = "SP_LISTAR_PAINEL_SINOPTICO"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = "C"

       	If trim(strOpcao) <> "" then
			.Parameters(3).Value = trim(strOpcao)
		Else
			.Parameters(3).Value = Null
		End if

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

'*********************************************************************************


'    Set cmdResultado = Server.CreateObject("ADODB.Command")
'   
'    With cmdResultado
'
'        .ActiveConnection = conConexao
'        .CommandType = 4
'		.CommandTimeout = 600
'        .CommandText = "SP_LISTAR_PAINEL_SINOPTICO"
'        
'        .Parameters.Refresh
'
'		.Parameters(1).Value = strProjeto
'
'    End With
'
'	set rs = Server.CreateObject("ADODB.RecordSet")
'
'	set rs = cmdResultado.Execute()

%>

	<html>

	<head>
	<title>Projeto BRACUSS </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Saneamento_Filtro" id="frmRelatorio_Saneamento_Filtro" method="post">
	<link rel="stylesheet" href="estilos/Light.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<p>


<%If Not rs.EOF Then%>
	<table width="100%" border="0">
		<tr>
			<td width="30%">&nbsp;</td>
			<td width="30%" align="middle">
				<p><b>
				<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Painel Sinóptico Projeto <%=rs("PROJ_NAME")%></font>
				</b></p>
			</td>
			<td width="30%">&nbsp;</td>
		</tr>
	</table>

	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0"  align=center>
				  
	  <tr height="17" style="height:12.75pt">
	    <td height="17" class="xl27" width=30px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
	    <td class="xl27" width="450px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
	    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
	    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
	    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
	    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Plan</font></b></td>
	    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		<td class="xl27" width="70px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
		  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Criticidade</font></b></td>
	  </tr>

		<%Do While Not rs.EOF%>

			<%If rs("TASK_IS_SUMMARY") = True Then %>
			  <tr height="17" style="height:12.75pt" bgcolor=LightGrey>
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="450px">
					<font face="Arial" size="1">
						<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
						<%=left("                              ",rs("task_outline_level")) %>
						<%=rs("TASK_NAME")%>
						</div>
					</font>
				</td>


			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
					</font>
				</td>

			    <td class="xl30" align="center" style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="center" style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
							<%=rs("Planejado")%>&nbsp;
					</font>
				</td>


			    <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>


			    <td class="xl30" style="border: 1 solid #666666" width="70px" align=center>
					<font face="Arial" size="1">
						<img src="<%=RetornaCriticidade(rs("CRITICIDADE"))%>" name="Bola" border="0" width=18>
					</font>
				</td>

			  </tr>
			<%Else%>

					<tr height="17" style="height:12.75pt">
					  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
							<font face="Arial" size="1">
								<%=rs("TASK_UID")%>
							</font>
						</td>

					  <td class="xl28" style="border: 1 solid #666666" width="450px">
							<font face="Arial" size="1">
								<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
									<%=left("                              ",rs("task_outline_level")) %>
									<%=rs("TASK_NAME")%>
								</div>
							</font>
						</td>

					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
							</font>
						</td>

					  <td class="xl30" align="center" style="border: 1 solid #666666" width="85px">
							<font face="Arial" size="1">
								<%=FormatarDataMon(rs("TASK_START_DATE"))%>
							</font>
						</td>
					  <td class="xl30" align="center" style="border: 1 solid #666666" width="85px">
							<font face="Arial" size="1">
								<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>
							</font>
						</td>


					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%=rs("Planejado")%>&nbsp;
							</font>
						</td>

					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%=rs("TASK_PCT_COMP")%>&nbsp;
							</font>
						</td>


						<td class="xl30" style="border: 1 solid #666666" width="70px" align=center>
							<font face="Arial" size="1">
								<img src="<%=RetornaCriticidade(rs("CRITICIDADE"))%>" name="Bola" border="0" width=18>
							</font>
						</td>

					</tr>
<%
			End If

			rs.MoveNext
			
			Loop%>

	</table>

<%else
	response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
end if%>

	<p align="right">
	<BR>
	<hr>

</FORM>
</body>
</html>