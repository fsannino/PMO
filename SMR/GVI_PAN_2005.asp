<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim cmdResultado
Dim rs


Function RetornaCriticidade(strCriticidade)
Dim Valor

	RetornaCriticidade = ""

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
	
End Function


	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_PAN2005_DETALHE"
        
    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Saneamento_Filtro" id="frmRelatorio_Saneamento_Filtro" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>

	<p>

	<table width="100%" border="0">
		<tr>
			<td width="30%">&nbsp;</td>
			<td width="30%" align="middle">
				<p><b>
				<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Painel Sinóptico PAN 2005</font>
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
								<%=rs("Planejado")%>
							</font>
						</td>

					  <td class="xl23" align="center" style="border: 1 solid #666666" width="50px">
							<font face="Arial" size="1">
								<%=rs("TASK_PCT_COMP")%>
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

	<p align="right">
	<BR>
	<hr>

</FORM>
</body>
</html>