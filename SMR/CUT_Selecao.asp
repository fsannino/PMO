<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->
<%

'	If trim(session("Login")) = "" Then
'		response.Redirect("./CUT_Login.asp")
'	End if

%>
<html>
<head>
<title>....::::::: Sinergia</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="estilos/sinergia.css">
</head>

<table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="18" colspan="4"><img src="../img/_0.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="221" valign="top">
		<div align="right"> 
			<p> 
				<img src="img/_0.gif" width="1" height="83"><br>
			</p>
        </div>
    </td>
    <td width="12">&nbsp;</td>
    <td width="467">
		<p>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				<strong> 
				</strong>
			</font>
	        <font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">
				<strong>
					<br>Entrada em Produção<br>
				</strong>
			</font>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				<strong>
					<br>Selecione a &aacute;rea desejada
				</strong>
			</font>
		</p>
		<p>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				&raquo; 
				<a href="./CUT_Login.asp" class="conf">Medição</a>
				<br>&raquo;
				<a href="CUT_Relatorio_Medicao_Detalhe.asp?strProj=1" class="conf">Cronograma</a>
				<br>&raquo;
				<a href="CUT_Painel_Sinoptico_CutOver.asp" class="conf">Painel Sinóptico Entrada em Produção</a>
				<br>&raquo;
				<a href="CUT_Login.asp?hidOrigem=./CUT_Modificacao_Filtro.asp" class="conf">Modificações</a>
			</font>
		</p>

		<p>
	        <font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">
				<strong>
					<br>Teste Carga para Entrada em Produção<br>
				</strong>
			</font>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				<strong>
					<br>Selecione a &aacute;rea desejada
				</strong>
			</font>
		</p>
		<p>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				&raquo; 
				<a href="CUT_Relatorio_Medicao_Detalhe.asp?strProj=2" class="conf">Cronograma</a>
				<br>&raquo;
				<a href="CUT_Painel_Sinoptico_Teste_Carga_CutOver.asp" class="conf">Painel Sinóptico Teste Carga para Entrada em Produção </a>
			</font>
		</p>
		<p>&nbsp;</p>


		<p>
			<font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">
				<img src="img/000025.gif" width="467" height="1">
			</font>
		</p>

		<p align="center">
			<font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				© 2003 Sinergia | A Petrobras integrada rumo ao futuro
			</font>
		</p>
		<p>&nbsp;</p>
	
	</td>
    <td width="80">&nbsp;</td>
  </tr>
</table>
</body>
</html>
