<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim cmdResultado
Dim rs

'Variaveis de Quantidades - V3
Dim Quant_Carga_FI_V5
Dim Quant_Carga_AC_V5
Dim Quant_Carga_CO_V5
Dim Quant_Carga_IS_V5
Dim Quant_Carga_SD_V5
Dim Quant_Carga_CBS_V5
Dim Quant_Carga_MES_V5
Dim Quant_Carga_PS_V5
Dim Quant_Carga_PO_V5
Dim Quant_Carga_LT_V5
Dim Quant_Carga_PP_V5
Dim Quant_Carga_QM_V5

'Variaveis de Sub-Totais - V3
Dim Quant_SF_FI_V5
Dim Quant_SF_CO_OIL_V5
Dim Quant_SF_PX_MES_V5

'Variaveis de Sub-Totais - V3
Dim Quant_Carga_Total_V5



'Variaveis de Quantidades - V4
Dim Quant_Carga_FI_V4
Dim Quant_Carga_AC_V4
Dim Quant_Carga_CO_V4
Dim Quant_Carga_IS_V4
Dim Quant_Carga_SD_V4
Dim Quant_Carga_CBS_V4
Dim Quant_Carga_MES_V4
Dim Quant_Carga_PS_V4
Dim Quant_Carga_PO_V4
Dim Quant_Carga_LT_V4
Dim Quant_Carga_PP_V4
Dim Quant_Carga_QM_V4

'Variaveis de Sub-Totais - V3
Dim Quant_SF_FI_V4
Dim Quant_SF_CO_OIL_V4
Dim Quant_SF_PX_MES_V4

'Variaveis de Sub-Totais - V3
Dim Quant_Carga_Total_V4

Dim dtInicio
Dim dtFim

	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()


Function RetornaValorPerc(strProj, strUID)

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_TAREFAS"
        
        .Parameters.Refresh

		.Parameters(1).Value = strProj
		.Parameters(2).Value = strUID

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

	If not rs.EOF Then
		RetornaValorPerc = rs("TASK_PCT_COMP")
	Else
		RetornaValorPerc = 	"0"
	End If

End Function


Function RetornaQuantidade(strProj, strUID)

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_QUANTIDADE_TESTE_CARGA"
        
        .Parameters.Refresh

		.Parameters(1).Value = strProj
		.Parameters(2).Value = strUID

    End With

	cmdResultado.Execute()

	RetornaQuantidade = cmdResultado.Parameters(3).Value

End Function


Function RetornaDatas(strProj, strUID, strDataIni, strDataFim)

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_TAREFAS"
        
        .Parameters.Refresh

		.Parameters(1).Value = strProj
		.Parameters(2).Value = strUID

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

	If not rs.EOF Then
		strDataIni = rs("DATA_INICIO")
		strDataFim = rs("DATA_FIM")
	End If

End Function


'V4
Quant_Carga_FI_V4 = RetornaQuantidade(1, 18690)
Quant_Carga_AC_V4 = RetornaQuantidade(1, 18707)

Quant_Carga_CO_V4 = RetornaQuantidade(1, 18717)
Quant_Carga_IS_V4 = RetornaQuantidade(1, 18739)
Quant_Carga_SD_V4 = RetornaQuantidade(1, 18744)

Quant_Carga_CBS_V4 = RetornaQuantidade(1, 18753)
Quant_Carga_MES_V4 = RetornaQuantidade(1, 18764)
Quant_Carga_PS_V4 = RetornaQuantidade(1, 18782)
'Quant_Carga_PO = RetornaQuantidade(1, 17868)
Quant_Carga_LT_V4 = RetornaQuantidade(1, 18787)
Quant_Carga_PP_V4 = RetornaQuantidade(1, 18793)
Quant_Carga_QM_V4 = RetornaQuantidade(1, 18800)

Quant_SF_FI_V4 = CINT(Quant_Carga_FI_V4) + CINT(Quant_Carga_AC_V4)
Quant_SF_CO_OIL_V4 = CINT(Quant_Carga_CO_V4) + CINT(Quant_Carga_IS_V4) + CINT(Quant_Carga_SD_V4)
Quant_SF_PX_MES_V4 = CINT(Quant_Carga_CBS_V4) + CINT(Quant_Carga_MES_V4) + CINT(Quant_Carga_PS_V4) + CINT(Quant_Carga_LT_V4) + CINT(Quant_Carga_PP_V4) + CINT(Quant_Carga_QM_V4)

Quant_Carga_Total_V4 = Quant_SF_FI_V4 + Quant_SF_CO_OIL_V4 + Quant_SF_PX_MES_V4




'**************************************************************
'V5
Quant_Carga_FI_V5 = RetornaQuantidade(1, 19224)
Quant_Carga_AC_V5 = RetornaQuantidade(1, 19238)

Quant_Carga_CO_V5 = RetornaQuantidade(1, 19246)
Quant_Carga_IS_V5 = RetornaQuantidade(1, 19262)
Quant_Carga_SD_V5 = RetornaQuantidade(1, 19267)

Quant_Carga_CBS_V5 = RetornaQuantidade(1, 19276)
Quant_Carga_MES_V5 = RetornaQuantidade(1, 19288)
Quant_Carga_PS_V5 = RetornaQuantidade(1, 19307)
'Quant_Carga_LT_V5 = RetornaQuantidade(1, )
'Quant_Carga_PP_V5 = RetornaQuantidade(1, 17877)
Quant_Carga_QM_V5 = RetornaQuantidade(1, 19311)

Quant_SF_FI_V5 = CINT(Quant_Carga_FI_V5) + CINT(Quant_Carga_AC_V5)
Quant_SF_CO_OIL_V5 = CINT(Quant_Carga_CO_V5) + CINT(Quant_Carga_IS_V5) + CINT(Quant_Carga_SD_V5)
Quant_SF_PX_MES_V5 = CINT(Quant_Carga_CBS_V5) + CINT(Quant_Carga_MES_V5) + CINT(Quant_Carga_PS_V5) + CINT(Quant_Carga_QM_V5)
' + CINT(Quant_Carga_LT_V5) + CINT(Quant_Carga_PP_V5)

Quant_Carga_Total_V5 = Quant_SF_FI_V5 + Quant_SF_CO_OIL_V5 + Quant_SF_PX_MES_V5

'******************************************************************************************


call RetornaDatas(1, 18235, dtInicio, dtFim)

%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Medicao_Detalhe" id="frmRelatorio_Medicao_Detalhe" action="GVI_Relatorio_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>
		<p>
		<table align=center>
			<tr>
				<td width=300px>
					&nbsp;
				</td>
				<TD width=400px  align=center>
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">TESTE DE CARGA - PETROBRAS</font></b>
				</td>
				<td width=300px>
					&nbsp;
				</td>
			</tr>
			<tr>
				<td width=300px>
					&nbsp;
				</td>
				<TD width=400px  align=center>
					<b><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif"><%=FormatarDataMon(dtInicio)%> a <%=FormatarDataMon(dtFim)%></font></b>
				</td>
				<td width=300px>
					&nbsp;
				</td>
			</tr>

		</table>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0"  align=center>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left bgcolor=#6699cc>
					<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif"> 
							&nbsp;
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
					<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
							Realização
						</font>
					</b>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							Teste de Carga - Petrobras
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 18235)%>%
						</font>
					</b>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Planejamento Geral
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 16412)%>%
						</font>
					</b>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Teste de Carga - Petrobras - Ciclo 1 (PEQ 330)
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 16413)%>%
						</font>
					</b>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Teste de Carga - Petrobras - Ciclo 2 (PEQ 330)
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 16414)%>%
						</font>
					</b>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Teste de Carga - Petrobras - Ciclo 3 (PEQ 330)
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 16681)%>%
						</font>
					</b>
				</td>
			</tr>


<%'********************************************************************************************************************%>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Teste de Carga - Petrobras - Ciclo 4 (PEQ 330)
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 16682)%>%
						</font>
					</b>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;Intervanção Inicial da Integração
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 18805)%>%
						</font>
					</b>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif"> 
							&nbsp;
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Quantidade
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Execução
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Verificação
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;Cargas
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_Total_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18688)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18823)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente FI
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_FI_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18689)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18824)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de FI
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_FI_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18690)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18825)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de AC
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_AC_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18707)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18842)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente CO + OIL
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_CO_OIL_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18716)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18851)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de CO
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_CO_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18717)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18852)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de IS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_IS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18739)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18874)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de SD
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_SD_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18744)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18879)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente PX + MES
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_PX_MES_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18752)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18887)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de CBS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_CBS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18753)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18888)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de MES
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_MES_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18764)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18899)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de PS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_PS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18782)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18917)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de LT
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_LT_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18787)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18922)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de PP
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_PP_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18793)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18928)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de QM
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_QM_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18800)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18935)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

<%'********************************************************************************************************************%>

			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;Teste de Carga - Petrobras - Teste de Cutover
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 19201)%>%
						</font>
					</b>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;Intervanção Inicial da Integração
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align="center">
					<b>
						<font face="Arial" size="1">
							<%=RetornaValorPerc(1, 19202)%>%
						</font>
					</b>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif"> 
							&nbsp;
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Quantidade
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Execução
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center bgcolor=#6699cc>
								<b>
									<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
										Verificação
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;Cargas
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_Total_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19222)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19313)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente FI
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_FI_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19223)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19314)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de FI
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_FI_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19224)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19315)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de AC
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_AC_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19238)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19329)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente CO + OIL
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_CO_OIL_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19245)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19336)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de CO
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_CO_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19246)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19337)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de IS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_IS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19262)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19353)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de SD
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_SD_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19267)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19358)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



			<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas sub-frente PX + MES
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_SF_PX_MES_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19275)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19366)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de CBS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_CBS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19276)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19367)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de MES
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_MES_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19288)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19379)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de PS
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_PS_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19307)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19398)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>


<!--			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de LT
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_LT_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18787)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18922)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de PP
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_PP_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18793)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 18928)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>
-->

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl27" width=400px style="border: 1 solid #666666" align=left>
					<b>
						<font face="Arial" size="1">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cargas de QM
						</font>
					</b>
				</td>
				<td height="17" class="xl27" width=300px style="border: 1 solid #666666" align=center>
					<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
					
						<tr height="17" style="height:12.75pt">
							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=Quant_Carga_QM_V4%>&nbsp;
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19311)%>%
									</font>
								</b>
							</td>

							<td height="17" class="xl27" width=100px style="border: 1 solid #666666" align=center>
								<b>
									<font face="Arial" size="1">
										<%=RetornaValorPerc(1, 19402)%>%
									</font>
								</b>
							</td>

						</tr>
					</table>
				</td>
			</tr>



		</table>

		<p align="right">
		<!--<input type="button" name="cmdSubmit" value="Enviar" onclick="Confirmar();">-->
		<BR>
<!--		<table cellspacing="0" cellpadding="0" align=center>
			<tr align=center>
				<td></td>
				<td align=center><a href="./GVI_selecao.asp" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Voltar ao Menu</font></a></td>
				<td></td>
			</tr>
		</table>-->
		<hr>

</FORM>
</body>
</html>