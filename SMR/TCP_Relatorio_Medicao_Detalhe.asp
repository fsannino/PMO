<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs
Dim cmdResultado

	'Abrindo uma conexão com o BD
	set conConexao = TCP_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_RELATORIO_MEDICAO_DETALHE"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()

	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Medicao_Detalhe" id="frmRelatorio_Medicao_Detalhe" action="GVI_Relatorio_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><%=rs("PROJ_NAME")%></font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=30px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td height="17" class="xl27" width=30px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">ID</font></b></td>
		    <td class="xl27" width="450px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
		    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="85px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Predecessora</font></b></td>
		    <td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Sucessora</font></b></td>
		  </tr>

		<%Do While Not rs.EOF%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor=#ebebeb>
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
					<font face="Arial" size="1">
						<%=rs("TASK_ID")%>&nbsp;
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

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="50px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						<%If rs("PROJ_ID") = 2 And rs("TASK_PCT_COMP") = 0 Then%>
							&nbsp;
						<%Else%>
							<%=rs("TASK_PCT_COMP")%>&nbsp;
						<%End if%>
					</font>
				</td>

			    <td class="xl30" style="border: 1 solid #666666" width="150px" align=center>
					<font face="Arial" size="1">
						<%=rs("PREDECESSORA")%>&nbsp;
					</font>
				</td>

			    <td class="xl22" style="border: 1 solid #666666" width="150px" align=center>
					<font face="Arial" size="1">
						<%=rs("SUCESSORA")%>&nbsp;
					</font>
				</td>

			  </tr>
			<%Else%>

				<tr height="17" style="height:12.75pt">
				  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
						<font face="Arial" size="1">
							<%=rs("TASK_UID")%>&nbsp;
						</font>
					</td>

				  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=30px>
						<font face="Arial" size="1">
							<%=rs("TASK_ID")%>&nbsp;
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

				  <td class="xl23" align="right" style="border: 1 solid #666666" width="50px">
						<font face="Arial" size="1">
							<%=rs("DURACAO")%>&nbsp;<%If rs("DURACAO") = 1 Then%>Dia<%Else%>Dias<%End If%>
						</font>
					</td>

				  <td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
						</font>
					</td>
				  <td class="xl30" align="right" style="border: 1 solid #666666" width="85px">
						<font face="Arial" size="1">
							<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
						</font>
					</td>

				  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
						<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
						</font>
					</td>


					<td class="xl30" style="border: 1 solid #666666" width="150px" align=center>
						<font face="Arial" size="1">
							<%=rs("PREDECESSORA")%>&nbsp;
						</font>
					</td>

				  <td class="xl22" style="border: 1 solid #666666" width="150px" align=center>
						<font face="Arial" size="1">
							<%=rs("SUCESSORA")%>&nbsp;
						</font>
					</td>

				</tr>

<%
			End If

			rs.MoveNext
			
		  Loop%>
		
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

	<%else

		response.write "<p><b><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)%>

</FORM>
</body>
</html>