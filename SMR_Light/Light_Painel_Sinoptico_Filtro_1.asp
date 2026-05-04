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
			<!--<font size="3" face="Verdana" color="#000000">-->
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Sistema da Equipe BRACUSS</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<form method="post" name="frmPainel_Sinoptico_Filtro" id="frmPainel_Sinoptico_Filtro">

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
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<!--<font size="1" face="Arial">-->
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcProjeto" style="WIDTH:300px;">
								<OPTION value=""></OPTION>
								<%Do While Not rs.EOF%>
									<OPTION value="<%=rs("PROJ_ID")%>"><%=rs("PROJ_NAME")%></OPTION>
									<%rs.MoveNext%>
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
	if (ValidarCombos())
	{
		document.frmPainel_Sinoptico_Filtro.action = "Light_Painel_Sinoptico.asp";
		document.frmPainel_Sinoptico_Filtro.submit();

	}
}


function Redefinir()
{
	document.frmPainel_Sinoptico_Filtro.reset();
}

function ValidarCombos()
{
	if (document.frmPainel_Sinoptico_Filtro.slcProjeto.value == "")
	{
		alert("Selecionar no mínimo um projeto.");
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
