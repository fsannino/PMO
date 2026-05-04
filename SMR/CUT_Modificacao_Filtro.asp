<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs1
Dim rs2
Dim cmdResultado


	If trim(session("Login")) = "" Then
		Response.Redirect ("./Erro.asp?Erro=Sua sessão expirou. Por-favor, logue-se novamente.&Voltar=true&IrPara=./CUT_LOGIN.ASP?hidOrigem=./CUT_Modificacao_Filtro.asp")
	End if

	strPlano = Request("slcPlano")

	'Abrindo uma conexão com o BD
	set conConexao = CUT_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_PLANOS"

    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()


	If trim(strPlano) <> "" Then

		Set cmdResultado = Server.CreateObject("ADODB.Command")
		    
		With cmdResultado
    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 600
		    .CommandText = "SP_LISTAR_MODULOS"

			.Parameters.Refresh
			.Parameters(1).Value = trim(strPlano)

		End With

		set rs2 = Server.CreateObject("ADODB.RecordSet")

		set rs2 = cmdResultado.Execute()

	End If

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
<form method="post" name="frm_Modificacao_Filtro" id="frm_Modificacao_Filtro">

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
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Planos:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <!--<select size="1" name="slcPlano" style="WIDTH:400px;" onchange="SelPlano();">-->
						    <select size="1" name="slcPlano" style="WIDTH:400px;">
								<OPTION value=""></OPTION>
								<%If not rs1.EOF Then
									While not rs1.EOF %>
										<OPTION value="<%=Trim(rs1("TASK_OUTLINE_NUM"))%>" <%If Trim(rs1("TASK_OUTLINE_NUM")) = Trim(strPlano) Then%>selected<%End If%>>
										<%=UCase(rs1("TASK_NAME"))%>
										</OPTION>
										<%rs1.MoveNext
									WEnd			
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>

<!--			  	<tr>
			    	<td bgcolor=#6699cc align="right">
						<b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Módulos:</font>
						</b>
					</td>
			    	<td align="right" bgcolor=#6699cc>
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
						    <select size="1" name="slcModulo" style="WIDTH:400px;">
						        <option value="">-----------------------------------------Todos-----------------------------------------</option>
								<%If trim(strPlano) <> "" Then
									 If not rs2.EOF Then
										While not rs2.EOF %>
											<OPTION value="<%=Trim(rs2("TASK_OUTLINE_NUM"))%>">
											<%=UCase(rs2("TASK_NAME"))%>
											</OPTION>
											<%rs2.MoveNext
										WEnd			
									 End If
								End If%>
						    </select>
						</font>
					</td>
			  	</tr>
			  	<tr>-->
			  	
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

function SelPlano()
{
	document.frm_Modificacao_Filtro.action = "CUT_Modificacao_Filtro.asp";
	document.frm_Modificacao_Filtro.submit();

}


function Confirmar()
{
	if (ValidarCombos())
	{
		document.frm_Modificacao_Filtro.action = "CUT_Modificacao_Detalhe.asp";
		document.frm_Modificacao_Filtro.submit();

	}
}

function Redefinir()
{
	document.frm_Modificacao_Filtro.reset();
}

function ValidarCombos()
{
	if (document.frm_Modificacao_Filtro.slcPlano.value == "")
	{
		alert("Selecionar no mínimo um Plano");
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
