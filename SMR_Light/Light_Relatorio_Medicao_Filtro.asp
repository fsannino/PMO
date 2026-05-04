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

	'Abrindo uma conexão com o BD
	set conConexao = LIGHT_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_PROJETO_MEDICAO"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()


    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_RESPONSAVEL"

    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()


    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_USUARIO"

    End With

	set rs2 = Server.CreateObject("ADODB.RecordSet")

	set rs2 = cmdResultado.Execute()

%>

<html>

<head>
	<title>Projeto BRACUSS </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/Light.css">
<body topmargin="0" leftmargin="0" bgcolor=White text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="center">
			<p><b>
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe BRACUSS</font>
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
				  		<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				  		<b>Por favor selecione o filtro</b>
				  		</font>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcProjeto" style="WIDTH:300px;">
								<OPTION value="">-----------------------------Todos-----------------------------</OPTION>
								<%Do While Not rs.EOF%>
									<OPTION value="<%=rs("PROJ_ID")%>"><%=rs("PROJ_NAME")%></OPTION>
									<%rs.MoveNext%>
								<%Loop%>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcUsuario" style="WIDTH:300px;">
								<OPTION value="">-----------------------------Todos-----------------------------</OPTION>
								<%Do While Not rs1.EOF%>
									<OPTION value="<%=rs1("TEXT_VALUE")%>"><%=rs1("TEXT_VALUE")%></OPTION>
									<%rs1.MoveNext%>
								<%Loop%>
						    </select>
						</font>
					</td>
			  	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Login:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcLogin" style="WIDTH:300px;">
								<OPTION value="">-----------------------------Todos-----------------------------</OPTION>
								<%Do While Not rs2.EOF%>
									<OPTION value="<%=rs2("Login")%>"><%=rs2("Nome_Usuario")%></OPTION>
									<%rs2.MoveNext%>
								<%Loop%>
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
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>
	</td>
	<td width="10%">&nbsp;</td>
</table>
<BR>
<BR>

</form>

<SCRIPT language=JavaScript>


function Confirmar()
{
	document.frmRelatorio_Medicao_Filtro.action = "Light_Relatorio_Medicao_Detalhe.asp";
	document.frmRelatorio_Medicao_Filtro.submit();

}


function Redefinir()
{
	document.frmRelatorio_Medicao_Filtro.reset();
}


</SCRIPT>

</body>
</html>
