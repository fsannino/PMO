<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%

'Recuperando dados da página de filtro
strEquipe = Request("strEquipe")
strData = Request("strData")
strTipo = Request("strTipo")
strCriticidade = Request("strCriticidade")
strUsuario = Request("strUsuario")
strCarteira = Request("strCarteira")

Dim rs1

'Abrindo uma conexão com o BD
set conConexao = SMR_AbrirConexaoBD()


strSql = "SP_LISTAR_COMITE "

Set cmdResultado = Server.CreateObject("ADODB.Command")
    
	With cmdResultado
    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 480
	    .CommandText = strSql
		    
		.Parameters.Refresh

	If trim(strEquipe) <> "" Then
		.Parameters(1).Value = strEquipe
	Else
		.Parameters(1).Value = null
	end if

	End With

set rs1 = Server.CreateObject("ADODB.RecordSet")

set rs1 = cmdResultado.Execute()
%>

<html>

<head>
	<title>Projeto Sinergia </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>

<link rel="stylesheet" href="estilos/sinergia.css">

<body topmargin="0" leftmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="center">
			<p><b><font size="3" face="Georgia, Times New Roman, Times, serif" color="#666666">Matriz de Criticidade</font></b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<form method="post" name="frmRelatorio_Criticidade_Detalhe" id="frmRelatorio_Criticidade_Detalhe">

<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0">
				<tr>
			    	<td align="right" colspan="2" height="30">
						<p align="center">
                        <b>
				  		<font face="Georgia, Times New Roman, Times, serif" size="2" color="#666666">Por favor selecione o Comite</font>
                        </b>
				 	</td>
			 	</tr>
			  	<tr>
			    	<td bgcolor="#639ACE" align="right">
						<b><font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">Comite:</font></b>
					</td>
			    	<!--<td width="67%" align="right" bgcolor="#639ACE">-->
			    	<td align="right" bgcolor="#639ACE">
						<font size="1" face="Arial">
						    <select size="1" name="cboComite">
						        <option value="">------------------------------------- Todos -------------------------------------</option>
								<%If not rs1.EOF Then
									While not rs1.EOF %>	
										<OPTION value="<%=rs1("COMITE")%>">
										<%=UCase(rs1("Desc_COMITE"))%>
										</OPTION>
										<%rs1.MoveNext
									WEnd			
								End If%>
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

<input type="hidden" id="strUsuario" name="strUsuario" value="<%=strUsuario%>">
<input type="hidden" id="hidComite" name="hidComite" value="">

<!--<table width="100%" border="0">
	<td width="30%">&nbsp;</td>
	<td width="30%" align="center">
		<input type="button" name="cmdSubmit" value="Enviar" onclick="Confirmar();">
		<input type="reset" name="cmdReset" value="Redefinir">
	</td>
	<td width="30%">&nbsp;</td>
</table>-->
</form>
<hr>

<%
'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)
%>

<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmRelatorio_Criticidade_Detalhe.action = "PMO_Relatorio_Issues_Detalhado.asp?strEquipe=<%=strEquipe%>&strData=<%=strData%>&strTipo=<%=strTipo%>&strCriticidade=<%=strCriticidade%>&strComite=" + document.frmRelatorio_Criticidade_Detalhe.cboComite.value + "&strCarteira=<%=strCarteira%>";
	document.frmRelatorio_Criticidade_Detalhe.submit();
}

function Redefinir()
{
	document.frmRelatorio_Criticidade_Detalhe.reset();
}

</SCRIPT>

</body>
</html>
