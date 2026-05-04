<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Dim rs1
Dim rs2
Dim rs3

Dim Busca

'Abrindo uma conexão com o BD
set conConexao = SMR_AbrirConexaoBD()

strSql = "SP_LISTAR_FRENTES " 

set rs1 = Server.CreateObject("ADODB.RecordSet")

rs1.OPEN STRSQL, conConexao

strSql = "SP_LISTAR_Equipes_GVI"

set rs2 = Server.CreateObject("ADODB.RecordSet")

rs2.OPEN STRSQL, conConexao

strSql = "SP_LISTAR_USUARIO Null, 0"

set rs3 = Server.CreateObject("ADODB.RecordSet")

rs3.OPEN STRSQL, conConexao

%>

<html>

<head>
	<title>Projeto Sinergia </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/sinergia.css">
<body topmargin="0" leftmargin="0" bgcolor=White text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="center">
			<p><b>
			<!--<font size="3" face="Verdana" color="#000000">-->
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe Sinergia</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<form method="post" name="frmRelatorio_Medicao_Filtro" id="frmRelatorio_Medicao_Filtro">

<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0">
				<tr>
			    	<td bgcolor=White align="right" colspan="2" height="30">
						<p align="center">
				  		<!--<font face="Arial" size="2" color="#FFFFFF">-->
				  		<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				  		<b>Por favor selecione o filtro</b>
				  		</font>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Medidor:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcUsuarioCLI" style="WIDTH:300px;">
								<%'IF Trim(rs3("Nome_Usuario")) = Trim(session("UsuarioCLI")) Then selected End If%>
								<%If not rs3.EOF Then%>
										<option value="">--------------------------- Todos --------------------------</option>
								  <%While not rs3.EOF %>	
										<OPTION value="<%=Trim(rs3("Login"))%>">
										<%=rs3("Nome_Usuario")%>
										</OPTION>
										<%rs3.MoveNext
									WEnd			
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Frente:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcFrente" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
								<%If not rs1.EOF Then
									While not rs1.EOF %>	
										<OPTION value="<%=UCase(rs1("Desc_Frente"))%>">
										<%=UCase(rs1("Desc_Frente"))%>
										</OPTION>
										<%rs1.MoveNext
									WEnd			
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>
			  	<tr>
			    	<td bgcolor="#6699cc" align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Equipe:</font>
						</b>
					</td>
			    	<td align="right" bgcolor="#6699cc">
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
							<select size="1" name="slcEquipe" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
								<%If not rs2.EOF Then
									While not rs2.EOF %>	
										<OPTION value="<%=UCase(rs2("Sigla_Equipe"))%>">
										<%=UCase(rs2("Desc_Equipe"))%>
										</OPTION>
										<%rs2.MoveNext
									WEnd			
								End If%>					
    							</select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Status:</font>
						</b>
					</td>
					
			    	<td align="right" bgcolor="#6699cc">
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						<!--<font size="1" face="Arial">-->
							<select size="1" name="slcCompleto" style="WIDTH:300px;">
						        <option value="">--------------------------- Todos --------------------------</option>
						        <option value="N">Em Andamento</option>
						        <option value="S">Finalizadas</option>
   							</select>
						</font>
					</td>
				</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<!--<font color="#000000" size="1" face="Arial">-->
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data de Inicio:</font>
						</b>
					</td>
					
			    	<td align="right" bgcolor="#6699cc">
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						<!--<font size="1" face="Arial">-->
							<select size="1" name="slcDataInicio" style="WIDTH:300px;">
						        <option value="">--------------------------- Todas --------------------------</option>
						        <option value="S">Inicio Planejado até Hoje</option>
   							</select>
						</font>
					</td>
				</tr>

			  	
			</table>
		<td width="10%">&nbsp;</td>
	</tr>
</table>
<br>
<br>
<table width="100%" border="0">
	<td width="10%">&nbsp;</td>
	<td width="80%" align="center">
		<a href="javascript:Confirmar();"><img src="img/000050.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:ConfirmarExcel();"><img src="img/000047.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>

<!--		<input type="Image" name="cmdSubmit" value="Visualizar" src="img/000050.gif" onClick="Confirmar();" align="absmiddle">
		<input type="Image" name="cmdSubmitExcel" value="Excel" src="img/000047.gif" onClick="ConfirmarExcel();" align="absmiddle">
		<input type="Image" name="cmdReset" value="Redefinir" src="img/000048.gif" onClick="Redefinir();" align="absmiddle">-->

<!--		<input type="button" name="cmdSubmit" value="Visualizar" onclick="Confirmar();">
		<input type="button" name="cmdSubmitExcel" value="Excel" onclick="ConfirmarExcel();">
		<input type="reset" name="cmdReset" value="Redefinir">-->
	</td>
	<td width="10%">&nbsp;</td>
</table>
<BR>
<BR>
<!--<table cellspacing="0" cellpadding="0" align=center>
	<tr align=center>
		<td></td>
		<td align=center><a href="./GVI_selecao.asp" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Voltar ao Menu</font></a></td>
		<td></td>
	</tr>
</table>-->

</form>

<%
'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)
%>

<SCRIPT language=JavaScript>

function Confirmar()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Medicao_Filtro.action = "PMO_Relatorio_Medicao_Detalhe.asp";
		document.frmRelatorio_Medicao_Filtro.submit();
	}
}

function ConfirmarExcel()
{
	if (ValidarCombos())
	{
		document.frmRelatorio_Medicao_Filtro.action = "PMO_Relatorio_Medicao_Detalhe_Excel.asp";
		document.frmRelatorio_Medicao_Filtro.submit();
	}
}

function Redefinir()
{
	document.frmRelatorio_Medicao_Filtro.reset();
}

function ValidarCombos()
{
	if (document.frmRelatorio_Medicao_Filtro.slcCompleto.value == "" &&
		document.frmRelatorio_Medicao_Filtro.slcDataInicio.value == "" &&
		document.frmRelatorio_Medicao_Filtro.slcEquipe.value == "" &&
		document.frmRelatorio_Medicao_Filtro.slcFrente.value == "" &&
		document.frmRelatorio_Medicao_Filtro.slcUsuarioCLI.value == "")
	{
		alert("Selecionar no mínimo uma opção");
		return false;
	}
	else
	{
		return true;
	}		
}

</SCRIPT>

</body>
</html>
