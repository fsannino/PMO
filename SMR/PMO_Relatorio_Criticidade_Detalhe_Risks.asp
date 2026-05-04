<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 5000

Dim rs
Dim rs1
Dim rs2
Dim rs3
Dim rs4

Dim cmdResultado
Dim strEquipe
Dim strDescEquipe
Dim StrData
Dim strHist

Dim intCont
Dim TotalCarteira1
Dim TotalCarteira2
Dim TotalCarteira3
Dim TotalNaoDefinido
Dim TotalGeral

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,4,2) & "/" & mid(strData,1,2) & "/" & mid(strData,7,4) 
End Function

strEquipe = Request("slcEquipe")
strComite = Request("strComite")
strHist = Request("strHist")

intCont = 0
TotalCarteira1 = 0
TotalCarteira2 = 0
TotalCarteira3 = 0
TotalNaoDefinido = 0
TotalGeral = 0

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
	If trim(strEquipe) <> "" Then
		With cmdResultado
    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 480
		    .CommandText = "SP_LISTAR_EQUIPES"
		    
			.Parameters.Refresh

			.Parameters(1).Value = strEquipe

		End With

		set rs = Server.CreateObject("ADODB.RecordSet")

		set rs = cmdResultado.Execute()

		strDescEquipe = rs("Desc_Eqp")
		
	Else
		strDescEquipe = "Todos"
	End If

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 480
        .CommandText = "SP_LISTAR_DATAS_CRITICIDADE_RISKS"

		If Trim(strHist) = "S" Then
			.Parameters(1).Value = strHist
		End If

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Criticidade_Detalhe" id="frmRelatorio_Criticidade_Detalhe" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<table width="100%" border="0">
		<tr>
			<td width="30%" height="10px">&nbsp;</td>
			<td width="40%" align="center" height="10px">
				<p><b><font size="3" face="Georgia, Times New Roman, Times, serif" color="#666666">Matriz de Criticidade - <%=strDescEquipe%></font></b></p>
			</td>
			<td width="30%" height="10px">&nbsp;</td>
		</tr>
	</table>

	<%If Not rs.EOF Then%>
		<p>
		<table cellspacing="0" cellpadding="0" align=center>

		  <tr height="17" style="height:12.75pt" >
				<td height="17" class="xl27" width="200px" align=left bgcolor=White nowrap>
					<b>
					<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
						Data:&nbsp;<%=rs("Data_Score")%>&nbsp;
					</font>
					</b>
				</td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				  
		  </tr>
		</table>

		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>

		  <tr height="17" style="height:12.75pt" >

				<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Status resolução / Criticidade</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Baixo</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Moderado</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Alto</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Critico</font></b></td>

		  </tr>

		<%StrData = rs("Data_Score")%>
		<%Do While Not rs.EOF%>


			<%  Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
				With cmdResultado
				    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 480
				    .CommandText = "SP_LISTAR_CRITICIDADE_BAIXO_RISKS"

					.Parameters.Refresh

					'.Parameters(1).Value = FormatarDataSQL(rs("Data_Score"))
					.Parameters(1).Value = rs("Data_Score")
					
					If trim(strEquipe) <> "" then
						.Parameters(2).Value = trim(strEquipe) 
					End if

				End With

				set rs1 = Server.CreateObject("ADODB.RecordSet")

				set rs1 = cmdResultado.Execute()

				Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
				With cmdResultado
				    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 480
				    .CommandText = "SP_LISTAR_CRITICIDADE_MODERADO_RISKS"

					.Parameters.Refresh

					'.Parameters(1).Value = FormatarDataSQL(rs("Data_Score"))
					.Parameters(1).Value = rs("Data_Score")


					If trim(strEquipe) <> "" then
						.Parameters(2).Value = trim(strEquipe) 
					End if

				End With

				set rs2 = Server.CreateObject("ADODB.RecordSet")

				set rs2 = cmdResultado.Execute()

				Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
				With cmdResultado
				    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 480
				    .CommandText = "SP_LISTAR_CRITICIDADE_ALTO_RISKS"

					.Parameters.Refresh

					'.Parameters(1).Value = FormatarDataSQL(rs("Data_Score"))
					.Parameters(1).Value = rs("Data_Score")

					If trim(strEquipe) <> "" then
						.Parameters(2).Value = trim(strEquipe) 
					End if

				End With

				set rs3 = Server.CreateObject("ADODB.RecordSet")

				set rs3 = cmdResultado.Execute()

				Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
				With cmdResultado
				    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 480
				    .CommandText = "SP_LISTAR_CRITICIDADE_CRITICO_RISKS"

					.Parameters.Refresh

					'.Parameters(1).Value = FormatarDataSQL(rs("Data_Score"))
					.Parameters(1).Value = rs("Data_Score")

					If trim(strEquipe) <> "" then
						.Parameters(2).Value = trim(strEquipe) 
					End if

				End With

				set rs4 = Server.CreateObject("ADODB.RecordSet")

				set rs4 = cmdResultado.Execute()

			%>

			<%If rs("Data_Score") <> StrData Then
				
				StrData = rs("Data_Score")%>
				<BR>
				<table cellspacing="0" cellpadding="0" align=center>
					<tr height="17" style="height:12.75pt" >
							<td height="17" class="xl27" width="200px" align=left bgcolor=White nowrap>
								<b>
								<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
									Data:&nbsp;<%=rs("Data_Score")%>&nbsp;
								</font>
								</b>
							</td>
							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

					</tr>
				</table>

				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>

					<tr height="17" style="height:12.75pt" >

							<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Status resolução / Criticidade</font></b></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Baixo</font></b></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Moderado</font></b></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Alto</font></b></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Critico</font></b></td>

					</tr>
				  
			<%End If%>

			<tr height="17" style="height:12.75pt" >
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" nowrap>
					<font face="Arial" size="1">
						No prazo
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=YellowGreen>
					<font face="Arial" size="1">

						<%If rs1.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs1("Andamento")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=A&strCriticidade=Baixo')">
									<%TotalCarteira3 = TotalCarteira3 + cint(rs1("Andamento"))%>
									<%=rs1("Andamento")%>
									</a>
								<%End If%>
							<%Else%>
								<%TotalCarteira3 = TotalCarteira3 + cint(rs1("Andamento"))%>
								<%=rs1("Andamento")%>
							<%End IF%>
						<%End If%>


					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=YellowGreen>
					<font face="Arial" size="1">

						<%If rs2.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs2("Andamento")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=A&strCriticidade=Medio')">
									<%TotalCarteira3 = TotalCarteira3 + cint(rs2("Andamento"))%>
									<%=rs2("Andamento")%>
									</a>
								<%End If%>
							<%Else%>
								<%TotalCarteira3 = TotalCarteira3 + cint(rs2("Andamento"))%>
								<%=rs2("Andamento")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If rs3.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>		
								<%If cint(rs3("Andamento")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=A&strCriticidade=Alto')">
										<%TotalCarteira2 = TotalCarteira2 + cint(rs3("Andamento"))%>
										<%=rs3("Andamento")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira2 = TotalCarteira2 + cint(rs3("Andamento"))%>
								<%=rs3("Andamento")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=#ff7372>
					<font face="Arial" size="1">
						<%If rs4.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs4("Andamento")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=A&strCriticidade=Critico')">
										<%TotalCarteira1 = TotalCarteira1 + cint(rs4("Andamento"))%>
										<%=rs4("Andamento")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira1 = TotalCarteira1 + cint(rs4("Andamento"))%>
								<%=rs4("Andamento")%>
							<%End IF%>
						<%End If%>						
			   		</font>
			  </td>


		</tr>


			<tr height="17" style="height:12.75pt" >
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" nowrap>
					<font face="Arial" size="1">
						Atrasados (menos de 10 dias)
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=YellowGreen>
					<font face="Arial" size="1">
						<%If rs1.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs1("Atrasadas10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=AT10&strCriticidade=Baixo')">
										<%TotalCarteira3 = TotalCarteira3 + cint(rs1("Atrasadas10"))%>
										<%=rs1("Atrasadas10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira3 = TotalCarteira3 + cint(rs1("Atrasadas10"))%>
								<%=rs1("Atrasadas10")%>
							<%End IF%>						
						<%End If%>
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If rs2.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs2("Atrasadas10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=AT10&strCriticidade=Medio')">
										<%TotalCarteira2 = TotalCarteira2 + cint(rs2("Atrasadas10"))%>
										<%=rs2("Atrasadas10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira2 = TotalCarteira2 + cint(rs2("Atrasadas10"))%>
								<%=rs2("Atrasadas10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If rs3.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs3("Atrasadas10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=AT10&strCriticidade=Alto')">
										<%TotalCarteira2 = TotalCarteira2 + cint(rs3("Atrasadas10"))%>
										<%=rs3("Atrasadas10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira2 = TotalCarteira2 + cint(rs3("Atrasadas10"))%>
								<%=rs3("Atrasadas10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=#ff7372>
					<font face="Arial" size="1">
						<%If rs4.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
 								<%If cint(rs4("Atrasadas10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=AT10&strCriticidade=Critico')">
										<%TotalCarteira1 = TotalCarteira1 + cint(rs4("Atrasadas10"))%>
										<%=rs4("Atrasadas10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira1 = TotalCarteira1 + cint(rs4("Atrasadas10"))%>
								<%=rs4("Atrasadas10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>


		</tr>

			<tr height="17" style="height:12.75pt" >
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" nowrap>
					<font face="Arial" size="1">
						Atrasados (mais de 10 dias)
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If rs1.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs1("AtrasadasMais10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=ATM10&strCriticidade=Baixo')">
										<%TotalCarteira2 = TotalCarteira2 + cint(rs1("AtrasadasMais10"))%>
										<%=rs1("AtrasadasMais10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira2 = TotalCarteira2 + cint(rs1("AtrasadasMais10"))%>
								<%=rs1("AtrasadasMais10")%>
							<%End IF%>
						<%End If%>
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If rs2.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs2("AtrasadasMais10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=ATM10&strCriticidade=Medio')">
										<%TotalCarteira2 = TotalCarteira2 + cint(rs2("AtrasadasMais10"))%>
										<%=rs2("AtrasadasMais10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira2 = TotalCarteira2 + cint(rs2("AtrasadasMais10"))%>
								<%=rs2("AtrasadasMais10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=#ff7372>
					<font face="Arial" size="1">
						<%If rs3.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs3("AtrasadasMais10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=ATM10&strCriticidade=Alto')">
										<%TotalCarteira1 = TotalCarteira1 + cint(rs3("AtrasadasMais10"))%>
										<%=rs3("AtrasadasMais10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira1 = TotalCarteira1 + cint(rs3("AtrasadasMais10"))%>
								<%=rs3("AtrasadasMais10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right  bgcolor=#ff7372>
					<font face="Arial" size="1">
						<%If rs4.EOF Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<%If cint(rs4("AtrasadasMais10")) = 0 Then%>
									0&nbsp;
								<%Else%>
									<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Risks_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strTipo=ATM10&strCriticidade=Critico')">
										<%TotalCarteira1 = TotalCarteira1 + cint(rs4("AtrasadasMais10"))%>
										<%=rs4("AtrasadasMais10")%>
									</a>
								<%End IF%>
							<%Else%>
								<%TotalCarteira1 = TotalCarteira1 + cint(rs4("AtrasadasMais10"))%>
								<%=rs4("AtrasadasMais10")%>
							<%End IF%>
						<%End If%>
			   		</font>
			  </td>


		</tr>

<%

			rs.MoveNext
			
			If Not rs.EOF Then
				If (rs("Data_Score") <> StrData) Then%>
					</table>
					<br>
					<table cellspacing="0" cellpadding="0" align=center>
						<tr height="17" style="height:12.75pt" >
							<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
							<b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						   			Total Carteira Nivel 1
						   		</font>
							</b>
							</td>
						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=#ff7372>
								<font face="Arial" size="1">
									<%If TotalCarteira1 = 0 Then%>
										0&nbsp;
									<%Else%>
										<%If intCont = 0 Then%>
											<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=rs("Data_Score")%>&strCarteira=C1&strComite=<%=strComite%>')">
						   					<%=TotalCarteira1%>
											</a>
										<%Else%>
						   					<%=TotalCarteira1%>
										<%End If%>
									<%End If%>
						   		</font>
						  </td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>
						
						</tr>
						
						<tr height="17" style="height:12.75pt" >

							<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
							<b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						   			Total Carteira Nivel 2
						   		</font>
							</b>
							</td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=Yellow>
								<font face="Arial" size="1">
									<%If TotalCarteira2 = 0 Then%>
										0&nbsp;
									<%Else%>
										<%If intCont = 0 Then%>
											<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=StrData%>&strCarteira=C2&strComite=<%=strComite%>')">
											<%=TotalCarteira2%>
											</a>
										<%Else%>
											<%=TotalCarteira2%>
										<%End If%>
									<%End If%>
						   		</font>
						  </td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

						
						</tr>

						<tr height="17" style="height:12.75pt" >
							<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
							<b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						   			Total Carteira Nivel 3
						   		</font>
							</b>
							</td>
						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=YellowGreen>
								<font face="Arial" size="1">
									<%If TotalCarteira3 = 0 Then%>
										0&nbsp;
									<%Else%>
										<%If intCont = 0 Then%>
											<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=StrData%>&strCarteira=C3&strComite=<%=strComite%>')">
											<%=TotalCarteira3%>
											</a>
										<%Else%>
											<%=TotalCarteira3%>
										<%End If%>
									<%End If%>
						   		</font>
						  </td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

						
						</tr>
						

						<tr height="17" style="height:12.75pt" >

							<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
							<b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
									Total Geral
								</font>
							</b>
							</td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
								<font face="Arial" size="1">
						   			<%=(TotalCarteira1+TotalCarteira2+TotalCarteira3+TotalNaoDefinido)%>
						   		</font>
						  </td>


							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

							<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>


									
						</tr>
						
					</table>

				<%	intCont = intCont + 1
					TotalCarteira1 = 0
					TotalCarteira2 = 0
					TotalCarteira3 = 0
					TotalNaoDefinido = 0%>

				<%End If
			End If%>
			
		<%Loop%>

		</table>
		<BR>
		<table cellspacing="0" cellpadding="0" align=center>

			<tr height="17" style="height:12.75pt" >
				<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
				<b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						Total Carteira Nivel 1
					</font>
				</b>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=#ff7372>
					<font face="Arial" size="1">
						<%If TotalCarteira1 = 0 Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=StrData%>&strCarteira=C1&strComite=<%=strComite%>')">
								<%=TotalCarteira1%>
								</a>
							<%Else%>
								<%=TotalCarteira1%>							
							<%End If%>
						<%End If%>
			   		</font>
			  </td>
						
				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>


			</tr>
			<tr height="17" style="height:12.75pt" >

				<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
				<b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						Total Carteira Nivel 2
					</font>
				</b>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=Yellow>
					<font face="Arial" size="1">
						<%If TotalCarteira2 = 0 Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=StrData%>&strCarteira=C2&strComite=<%=strComite%>')">
								<%=TotalCarteira2%>
								</a>
							<%Else%>
								<%=TotalCarteira2%>
							<%End If%>
						<%End If%>
			   		</font>
			  </td>


				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

						
			</tr>

			<tr height="17" style="height:12.75pt" >

				<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
				<b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						Total Carteira Nivel 3
					</font>
				</b>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right bgcolor=YellowGreen>
					<font face="Arial" size="1">
						<%If TotalCarteira3 = 0 Then%>
							0&nbsp;
						<%Else%>
							<%If intCont = 0 Then%>
								<a href="JavaScript:abreJanelaRelRisks('PMO_Relatorio_Criticidade_Comite_Filtro_Risks.asp?strEquipe=<%=strEquipe%>&strData=<%=StrData%>&strCarteira=C3&strComite=<%=strComite%>')">
								<%=TotalCarteira3%>
								</a>
							<%Else%>
								<%=TotalCarteira3%>
							<%End If%>
						<%End If%>
			   		</font>
			  </td>
						
				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

			</tr>


			<tr height="17" style="height:12.75pt" >

				<td height="17" class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap>
				<b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
						Total Geral
					</font>
				</b>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
			   			<%=(TotalCarteira1+TotalCarteira2+TotalCarteira3+TotalNaoDefinido)%>
			   		</font>
			  </td>


				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=White nowrap></td>

									
			</tr>

		</table>

		<%If intCont = 0 Then%>
			<BR>
			<table cellspacing="0" cellpadding="0" align=center>
				<tr align=center>
					<td></td>
					<td align=center><a href="JavaScript:Confirmar();" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Historico</font></a></td>
					<td>&nbsp;</td>

					<td>&nbsp;</td>
					<td align=center><a href="JavaScript:abreJanelaRelRisks('PMO_Tab_Av_Critic_Issue.asp')" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Legenda</font></a></td>
					<td></td>

				</tr>
			</table>
		<%Else%>
			<BR>
			<table cellspacing="0" cellpadding="0" align=center>
				<tr align=center>
					<td></td>
					<td align=center><a href="JavaScript:abreJanelaRelRisks('PMO_Tab_Av_Critic_Issue.asp')" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Legenda</font></a></td>
					<td></td>
				</tr>
			</table>		
		<%End If%>

		<p align="right">
		<BR>
		<hr>

		<input type="hidden" id="strHist" name="strHist" value="">
		<input type="hidden" id="slcEquipe" name="slcEquipe" value="<%=strEquipe%>">


	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)%>


<script language="JavaScript" type="text/JavaScript">

function abreJanelaRelRisks(Url) 
{
  window.open(Url,'Aviso','');
}

function Confirmar()
{
	document.frmRelatorio_Criticidade_Detalhe.strHist.value = "S";
	document.frmRelatorio_Criticidade_Detalhe.action = "PMO_Relatorio_Criticidade_Detalhe_Risks.asp";
	document.frmRelatorio_Criticidade_Detalhe.submit();
}


    </script>


</FORM>
</body>
</html>
