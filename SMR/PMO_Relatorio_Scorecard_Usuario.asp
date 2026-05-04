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

Dim strData
Dim intTotalUsuario
Dim	intTotalAndamento
Dim	intTotalAtrasadas10
Dim	intTotalAtrasadasMais10
Dim intTotalGeral

intTotalUsuario = 0
intTotalAndamento = 0
intTotalAtrasadas10 = 0
intTotalAtrasadasMais10 = 0
intTotalGeral = 0
 
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	Set cmdResultado = Server.CreateObject("ADODB.Command")
    
	With cmdResultado
    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 480
	    .CommandText = "SP_LISTAR_DATAS_CRITICIDADE"

	End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 480
        .CommandText = "SP_LISTAR_SCORECARD_USUARIO"

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
				<p><b><font size="3" face="Georgia, Times New Roman, Times, serif" color="#666666">Issues por Usuário</font></b></p>
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
						Data:&nbsp;<%=rs1("Data_Score")%>&nbsp;
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
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Usuário</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">No Prazo</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas até 10 dias</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Atrasadas mais de 10 dias</font></b></td>

				<td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc nowrap><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Total</font></b></td>

		  </tr>

		<%Do While Not rs.EOF%>

			<tr height="17" style="height:12.75pt" >
			  <td height="17" class="xl22" align=left style="border: 1 solid #666666" width="200px" nowrap>
					<font face="Arial" size="1">
						<%=rs("Ret_Usuario")%>&nbsp;
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If cint(rs("NoPrazo")) = 0 Then%>
							0&nbsp;
						<%Else%>
							<a href="JavaScript:abreJanelaRelIssues('PMO_Relatorio_Issues_Detalhado.asp?strUsuario=<%=rs("Ret_Usuario")%>&strData=<%=rs1("Data_Score")%>&strTipo=A')">
								<%=rs("NoPrazo")%>&nbsp;
							</a>
						<%End If%>
					</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If cint(rs("Atras10")) = 0 Then%>
							0&nbsp;
						<%Else%>
							<a href="JavaScript:abreJanelaRelIssues('PMO_Relatorio_Issues_Detalhado.asp?strUsuario=<%=rs("Ret_Usuario")%>&strData=<%=rs1("Data_Score")%>&strTipo=AT10')">
				   				<%=rs("Atras10")%>&nbsp;
							</a>
						<%End If%>

			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%If cint(rs("AtrasM10")) = 0 Then%>
							0&nbsp;
						<%Else%>
							<a href="JavaScript:abreJanelaRelIssues('PMO_Relatorio_Issues_Detalhado.asp?strUsuario=<%=rs("Ret_Usuario")%>&strData=<%=rs1("Data_Score")%>&strTipo=ATM10')">
			   					<%=rs("AtrasM10")%>&nbsp;
							</a>
						<%End If%>
			   		</font>
			  </td>

			  <td class="xl28" style="border: 1 solid #666666" width="100px" nowrap align=right>
					<font face="Arial" size="1">
						<%intTotalUsuario = cint(rs("NoPrazo")) + cint(rs("Atras10")) + cint(rs("AtrasM10"))%>
			   			<%=intTotalUsuario%>&nbsp;
			   		</font>
			  </td>


		</tr>

<%

			intTotalAndamento = intTotalAndamento + cint(rs("NoPrazo"))
			intTotalAtrasadas10 = intTotalAtrasadas10 + cint(rs("Atras10")) 
			intTotalAtrasadasMais10 = intTotalAtrasadasMais10 + cint(rs("AtrasM10"))
			intTotalGeral = intTotalGeral + intTotalUsuario
			
			rs.MoveNext
%>
			
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

</script>


</FORM>
</body>
</html>