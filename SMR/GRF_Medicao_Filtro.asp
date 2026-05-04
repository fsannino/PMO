<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->
<%
Dim intDia
Dim intMes
Dim intAno
Dim I
Dim vetDias()

intMes = cint(Month(now()))
intAno = cint(Year(now()))

If intMes = 4 or intMes = 6 or intMes = 9 or intMes = 11 Then
	intDia = cint(30)
	ReDim vetDias(29)
ElseIf intMes = 2 Then
	If (intAno/4) <> int((intAno/4)) Then
		intDia = cint(28)
		ReDim vetDias(27)
	Else
		intDia = cint(29)
		ReDim vetDias(28)
	End If
Else
	intDia = cint(31)
	ReDim vetDias(30)
End If


For I = 1 To UBound(vetDias)
	If I < 10 Then
		vetDias(I) = "0" & I & "/" & intMes & "/" & intAno
	Else
		vetDias(I) = I & "/" & intMes & "/" & intAno
	End If
Next

%>

<html>
<head>
	<title>Projeto Sinergia </title>
	<!-- #include file="includes/EstiloIndicadores.inc" -->
</head>
<link rel="stylesheet" href="estilos/sinergia.css">
<body topmargin="0" leftmargin="0" bgcolor=white text="#000000" link="#0000ff" vlink="#0000ff" alink="#0000ff">
<form method="post" name="frmMedicao_Filtro" id="frmMedicao_Filtro">

<table width="100%" border="0">
	<tr>
		<td width="30%">&nbsp;</td>
		<td width="30%" align="middle">
			<p><b>
			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Indicadores de Estabilização</font>
			</b></p>
		</td>
		<td width="30%">&nbsp;</td>
	</tr>
</table>
<table width="100%" border="0">
	<tr>
		<td width="10%">&nbsp;</td>
		<td width="80%"  align=center>
			<table border="0" >
				<tr>
			    	<td bgcolor=white align="right" colspan="2" height="30">
						<p align="center">
				  			<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">
				  				<b>Por favor selecione o filtro</b>
				  			</font>
				  		</p>
				 	</td>
			 	</tr>

			  	<tr>
			    	<td bgcolor=#6699cc align="middle">
						<b>
						<font color=white size="1" face="Georgia, Times New Roman, Times, serif">Data:</font>
						</b>
					</td>
			    	<td bgcolor=#6699cc align="middle">
						<font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">
							<SELECT size=1 name="slcData" style="WIDTH:100px;">
								<%For I = 1 To UBound(vetDias)%>
									<OPTION value="<%=vetDias(I)%>"><%=vetDias(I)%></OPTION>
								<%Next%>
							</SELECT>
						</font>
					</td>
			  	</tr>
			  	<br>
				<br>
			</table>
		</TD>
		<td width="10%">&nbsp;</td>
	</TR>
</TABLE>
<BR>
<table width="100%" border="0">
	<td width="10%">&nbsp;</td>
	<td width="80%" align="center">
		<a href="javascript:Confirmar();"><img src="img/000050.gif" width="73" height="16" border="0" align="absmiddle"></a>
		<a href="javascript:Redefinir();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>
	</td>
	<td width="10%">&nbsp;</td>
</table>

</form>

<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmMedicao_Filtro.action = "GRF_Medicao_Detalhe.asp";
	document.frmMedicao_Filtro.submit();
}

function Redefinir()
{
	document.frmMedicao_Filtro.reset();
}


</SCRIPT>

</body>
</html>
