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
Dim vetDados
Dim vetDados_Disc
Dim strUsuario
Dim strID
Dim strData
Dim strValor
Dim strQtd
Dim strSql
Dim Cont

	'Abrindo uma conexão com o BD
	set conConexao = GRF_AbrirConexaoBD()

	If trim(session("Login")) = "" Then
		Response.Redirect ("./Erro.asp?Erro=Sua sessão expirou. Por-favor, logue-se novamente.&Voltar=true&IrPara=./GRF_LOGIN.ASP")
	End if

	strOperacao		= Request("hidOperacao")
	strUsuario      = Request("strUsuario")
	strData         = Request("slcData")


	if strOperacao <> "" Then
		
		vetDados = split(Request("hidDados"),";")

		For I = 0 to ubound(vetDados)

			vetDados_Disc = split(vetDados(I),"|")
			
			strID    = vetDados_Disc(0)
			strValor = replace(vetDados_Disc(1),",",".")
			strQtd   = vetDados_Disc(2)

			Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 1200
			    .CommandText = "SP_INCLUIR_LOG_INDICADORES"
			    
			    .Parameters.Refresh

				.Parameters(1).Value = trim(strUsuario)
				.Parameters(2).Value = strID
				.Parameters(3).Value = strValor
				.Parameters(4).Value = strQtd

			End With
			
			cmdResultado.Execute

			Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 1200
			    .CommandText = "SP_ATUALIZAR_INDICADORES"
			    
			    .Parameters.Refresh

				.Parameters(1).Value = strID
				.Parameters(2).Value = strValor
				.Parameters(3).Value = strQtd

			End With
			
			cmdResultado.Execute

'			strSql = "EXEC SP_INCLUIR_LOG_INDICADORES '" & strUsuario & "', " & strID & ", " & strValor & ", " & strQtd & " "
			
'			conConexao.execute strSql
			
'			strSql = "EXEC SP_ATUALIZAR_INDICADORES " & strID & ", " & strValor & ", " & strQtd & " "
	
'			conConexao.execute strSql
			
		Next
		
		strOperacao = ""

	End if

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_INDICADORES"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = strData


    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmMedicao_Detalhe_GRF" id="frmMedicao_Detalhe_GRF" action="GRF_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b>
		<table align=center>
			<tr>
				<td width=300px>
					&nbsp;
				</td>
				<TD width=400px  align=center>
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Indicadores de Estabilização</font></b>
				</td>
				<td width=300px>
					&nbsp;
				</td>
			</tr>
		</table>
		
		</b>
		<BR>
		
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=200px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Diretoria</font></b></td>
		    <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Unidade</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Indicador</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Valor</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Quantidade</font></b></td>
		  </tr>

		<%Cont = 0%>

		<%Do While Not rs.EOF%>

			<input type="hidden" id="hidId" name="hidId" value="<%=rs("ID")%>">

			<tr height="17" style="height:12.75pt">
				<td height="17" class="xl22" align=center style="border: 1 solid #666666" width=200px>
					<font face="Arial" size="1">
						<%=rs("Diretoria")%>&nbsp;
					</font>
				</td>

				<td class="xl28" style="border: 1 solid #666666" width="200px" align=center>
					<font face="Arial" size="1">
						<%=rs("Unidade")%>&nbsp;
					</font>
				</td>

				<td class="xl30" align=center style="border: 1 solid #666666" width="100px">
					<font face="Arial" size="1">
						<%=rs("Indicador")%>&nbsp;
					</font>
				</td>

				<td class="xl23" align=center style="border: 1 solid #666666" width="100px">
				  	<font face="Arial" size="1">
						<%=rs("Data")%>&nbsp;
				  	</font>
				  </td>

				<td class="xl23" align=center style="border: 1 solid #666666" width="100px">
				  	<font face="Arial" size="1">
				  		<input id="txtValor" name="txtValor" value="<%=rs("Valor")%>"  maxLength=12 size=12 onKeyUp="CaracteresValidos('1234567890,',this.value);">
				  		<input type="hidden" id="hidValor" name="hidValor" value="<%=rs("Valor")%>">
				  	</font>
				  </td>

				<td class="xl23" align=center style="border: 1 solid #666666" width="100px">
				  	<font face="Arial" size="1">
				  		<input id="txtQtd" name="txtQtd" value="<%=rs("Qtd")%>"  maxLength=9 size=8 onKeyUp="CaracteresValidos('1234567890',this.value);">
				  		<input type="hidden" id="hidQtd" name="hidQtd" value="<%=rs("Qtd")%>">
				  	</font>
				  </td>

			</tr>

			<%Cont = Cont + 1%>

			<%rs.MoveNext%>
			
		<%Loop%>
		
		</table>

		<p align="right">

		<table align=center>
			<tr>
				<td width=680px>
					&nbsp;
				</td>
				<td width=60px align=right>
					<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a>
				</td>
			</tr>
		</table>


		<BR>

		<img src="img/_0.gif" width="2" height="2">
		<hr>
		<input type="hidden" id="hidCont" name="hidCont" value="<%=Cont%>">
		<input type="hidden" id="hidDados" name="hidDados" value="">
		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
		<input type="hidden" id="slcData" name="slcData" value="<%=strData%>">
		<input type="hidden" id="strUsuario" name="strUsuario" value="<%=trim(session("Login"))%>">

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)

	%>
	
	
<SCRIPT language=JavaScript>

function Confirmar()
{
	var intCont	= (document.frmMedicao_Detalhe_GRF.hidCont.value - 1)
	var strAux = ""

	document.frmMedicao_Detalhe_GRF.style.cursor = "wait";
			
	document.frmMedicao_Detalhe_GRF.hidDados.value = "";
	document.frmMedicao_Detalhe_GRF.hidOperacao.value = "";
	
	if (intCont == 0)
	{
		for(var i = 0; i <= intCont;i++)
		{

			if ((document.frmMedicao_Detalhe_GRF.txtValor.value != document.frmMedicao_Detalhe_GRF.hidValor.value) || 
			    (document.frmMedicao_Detalhe_GRF.txtQtd.value != document.frmMedicao_Detalhe_GRF.hidQtd.value))
			{
				if (strAux == "")
				{
					strAux = document.frmMedicao_Detalhe_GRF.hidId.value + "|" + document.frmMedicao_Detalhe_GRF.txtValor.value + "|" + document.frmMedicao_Detalhe_GRF.txtQtd.value;
				}	
				else
				{
					strAux = strAux + ";" + document.frmMedicao_Detalhe_GRF.hidId.value + "|" + document.frmMedicao_Detalhe_GRF.txtValor.value + "|" + document.frmMedicao_Detalhe_GRF.txtQtd.value;
				}
			}
		}
	}
	else
	{
		for(var i = 0; i <= intCont;i++)
		{


			if ((document.frmMedicao_Detalhe_GRF.txtValor(i).value != document.frmMedicao_Detalhe_GRF.hidValor(i).value) || 
			    (document.frmMedicao_Detalhe_GRF.txtQtd(i).value != document.frmMedicao_Detalhe_GRF.hidQtd(i).value))
			{
				if (strAux == "")
				{
					strAux = document.frmMedicao_Detalhe_GRF.hidId(i).value + "|" + document.frmMedicao_Detalhe_GRF.txtValor(i).value + "|" + document.frmMedicao_Detalhe_GRF.txtQtd(i).value;
				}	
				else
				{
					strAux = strAux + ";" + document.frmMedicao_Detalhe_GRF.hidId(i).value + "|" + document.frmMedicao_Detalhe_GRF.txtValor(i).value + "|" + document.frmMedicao_Detalhe_GRF.txtQtd(i).value;
				}
			}
		}
	}

	
	if (strAux != "")
	{
		document.frmMedicao_Detalhe_GRF.hidDados.value = strAux;
		document.frmMedicao_Detalhe_GRF.hidOperacao.value = 'A';
		document.frmMedicao_Detalhe_GRF.style.cursor = "";
		document.frmMedicao_Detalhe_GRF.submit();
	}
	else
	{
		alert("Nenhum registro foi alterado");
		document.frmMedicao_Detalhe_GRF.style.cursor = "";
	}
}

</SCRIPT>
</FORM>
</body>
</html>
