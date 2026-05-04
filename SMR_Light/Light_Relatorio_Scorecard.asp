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

Dim StrData
Dim strHist

Dim intCont
Dim intTotalProj
Dim	intTotalAndamento
Dim	intTotalAtrasadas10
Dim	intTotalAtrasadasMais10
Dim intTotalGeral

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,4,2) & "/" & mid(strData,1,2) & "/" & mid(strData,7,4) 
End Function

strHist = Request("strHist")

intCont = 0
intTotalProj = 0
intTotalAndamento = 0
intTotalAtrasadas10 = 0
intTotalAtrasadasMais10 = 0
intTotalGeral = 0
 
	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

	If Trim(strHist) = "" Then

		Set cmdResultado = Server.CreateObject("ADODB.Command")
    
		With cmdResultado
    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 480
		    .CommandText = "SP_LISTAR_DATAS_CRITICIDADE"

		End With

		set rs1 = Server.CreateObject("ADODB.RecordSet")

		set rs1 = cmdResultado.Execute()

	End If

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 480
        .CommandText = "SP_LISTAR_SCORECARD"

		If Trim(strHist) = "" Then
			'.Parameters(1).Value = FormatarDataSQL(rs1("Data_Score"))
			.Parameters(1).Value = rs1("Data_Score")
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
	<FORM name="frmRelatorio_Scorecard" id="frmRelatorio_Scorecard" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<table width="100%" border="0">
		<tr>
			<td width="30%" height="10px">&nbsp;</td>
			<td width="40%" align="center" height="10px">
				<p><b><font size="3" face="Georgia, Times New Roman, Times, serif" color="#666666">Scorecard de Issues</font></b></p>
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
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipe</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">No Prazo</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas até 10 dias</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas mais de 10 dias</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Total</font></b></td>

		  </tr>

		<%StrData = rs("Data_Score")%>
		<%Do While Not rs.EOF%>

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
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipe</font></b></td>

						<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">No Prazo</font></b></td>

						<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas até 10 dias</font></b></td>

						<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas mais de 10 dias</font></b></td>

						<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Total</font></b></td>

					</tr>
				  
			<%End If%>

			<tr height="17" style="height:12.75pt" >
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" nowrap>
					<font face="Arial" size="1">
						<%=rs("Desc_Eqp")%>&nbsp;
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If intCont = 0 Then%>
							<%If cint(rs("Andamento")) = 0 Then%>
								<%=rs("Andamento")%>&nbsp;
							<%Else%>
								<a href="JavaScript:abreJanelaRelIssues('Light_Relatorio_Issues_Detalhado.asp?strEquipe=<%=rs("Cod_Equipe")%>&strData=<%=rs("Data_Score")%>&strTipo=A')">
								<%=rs("Andamento")%>&nbsp;
								</a>
							<%End If%>
						<%Else%>
							<%=rs("Andamento")%>&nbsp;
						<%End If%>
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If intCont = 0 Then%>
							<%If cint(rs("Atrasadas10")) = 0 Then%>
			   					<%=rs("Atrasadas10")%>&nbsp;
							<%Else%>
								<a href="JavaScript:abreJanelaRelIssues('Light_Relatorio_Issues_Detalhado.asp?strEquipe=<%=rs("Cod_Equipe")%>&strData=<%=rs("Data_Score")%>&strTipo=AT10')">
			   					<%=rs("Atrasadas10")%>&nbsp;
								</a>
							<%End If%>
						<%Else%>
			   				<%=rs("Atrasadas10")%>&nbsp;
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If intCont = 0 Then%>
							<%If cint(rs("AtrasadasMais10")) = 0 Then%>
			   					<%=rs("AtrasadasMais10")%>&nbsp;
							<%Else%>	
								<a href="JavaScript:abreJanelaRelIssues('Light_Relatorio_Issues_Detalhado.asp?strEquipe=<%=rs("Cod_Equipe")%>&strData=<%=rs("Data_Score")%>&strTipo=ATM10')">
			   					<%=rs("AtrasadasMais10")%>&nbsp;
								</a>
							<%End If%>
						<%Else%>
			   				<%=rs("AtrasadasMais10")%>&nbsp;
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%intTotalProj = cint(rs("Andamento")) + cint(rs("Atrasadas10")) + cint(rs("AtrasadasMais10"))%>
			   			<%=intTotalProj%>&nbsp;
			   		</font>
			  </td>


		</tr>

<%

			intTotalAndamento = intTotalAndamento + cint(rs("Andamento"))
			intTotalAtrasadas10 = intTotalAtrasadas10 + cint(rs("Atrasadas10")) 
			intTotalAtrasadasMais10 = intTotalAtrasadasMais10 + cint(rs("AtrasadasMais10"))
			intTotalGeral = intTotalGeral + intTotalProj
			
			rs.MoveNext
			
			If Not rs.EOF Then
				If (rs("Data_Score") <> StrData) Then%>
					
						<tr height="17" style="height:12.75pt" >
						  <!--<td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" bgcolor=#6699cc nowrap>-->
						  <td class="xl27" width="100px" style="border: 1 solid #666666" align=left bgcolor=#6699cc nowrap>
						  <b>
								<font color=White size="1" face="Georgia, Times New Roman, Times, serif">									
									Total&nbsp;
								</font>
						  </b>
						  </td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
								<font face="Arial" size="1">
									<%=intTotalAndamento%>&nbsp;
								</font>
						  </td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
								<font face="Arial" size="1">
						   			<%=intTotalAtrasadas10%>&nbsp;
						   		</font>
						  </td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
								<font face="Arial" size="1">
						   			<%=intTotalAtrasadasMais10%>&nbsp;
						   		</font>
						  </td>

						  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
								<font face="Arial" size="1">
						   			<%=intTotalGeral%>&nbsp;
						   		</font>
						  </td>


					</tr>

					</table>

					<%  intCont = intCont + 1
						intTotalAndamento = 0
						intTotalAtrasadas10 = 0
						intTotalAtrasadasMais10 = 0
						intTotalGeral = 0
					%>
				<%End If
			End If%>
			
		<%Loop%>

			<tr height="17" style="height:12.75pt" >

				<!--<td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" bgcolor=#6699cc nowrap>-->
				<td class="xl27" width="100px" style="border: 1 solid #666666" align=left bgcolor=#6699cc nowrap>
				<b>
				  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">									
				  		Total&nbsp;
				  	</font>
				</b>
				</td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%=intTotalAndamento%>&nbsp;
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
			   			<%=intTotalAtrasadas10%>&nbsp;
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
			   			<%=intTotalAtrasadasMais10%>&nbsp;
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
			   			<%=intTotalGeral%>&nbsp;
			   		</font>
			  </td>


		</tr>
		</table>

		<%If intCont = 0 Then%>
			<BR>
			<table cellspacing="0" cellpadding="0" align=center>
				<tr align=center>
					<td></td>
					<td align=center><a href="JavaScript:Confirmar();" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Historico</font></a></td>
					<td></td>
				</tr>
			</table>
		<%End If%>

		<p align="right">
		<!--<input type="button" name="cmdSubmit" value="Enviar" onclick="Confirmar();">-->
		<BR>
		<hr>

		<input type="hidden" id="strHist" name="strHist" value="">
<!--		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">-->

	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)%>

<script language="JavaScript" type="text/JavaScript">

function abreJanelaRelIssues(Url) 
{
  window.open(Url,'Aviso','');
}

function Confirmar()
{
	document.frmRelatorio_Scorecard.strHist.value = "S";
	document.frmRelatorio_Scorecard.action = "Light_Relatorio_Scorecard.asp";
	document.frmRelatorio_Scorecard.submit();
}


</script>


</FORM>
</body>
</html>
