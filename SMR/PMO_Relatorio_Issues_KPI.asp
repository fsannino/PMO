<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<% 
Dim strProjeto
Dim strData
Dim strTipo
Dim strCriticidade
Dim strUsuario
Dim strCarteira
Dim strComite
Dim intConta


Dim strDescricao
DIM strResolucao
Dim intQtdeLinhas
Dim intQtdeColunas

Dim rsRelatorios
Dim cmdResultado

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,7,4) + mid(strData,4,2) + mid(strData,1,2)
End Function

'Recuperando dados da página de filtro
strProjeto = Request("strProjeto")
strData = Request("strData")
strTipo = Request("strTipo")
strCriticidade = Request("strCriticidade")
strUsuario = Request("strUsuario")
strCarteira = Request("strCarteira")
strComite = Request("strComite")
intConta = 0

'Abrindo uma conexão com o BD
set conConexao = SMR_AbrirConexaoBD()

Set cmdResultado = Server.CreateObject("ADODB.Command")
    
With cmdResultado
		    
    .ActiveConnection = conConexao
    .CommandType = 4
	.CommandTimeout = 600
    .CommandText = "SP_LISTAR_ISSUES_KPI"
				    
End With

	
set rsRelatorios = Server.CreateObject("ADODB.RecordSet")

set rsRelatorios = cmdResultado.Execute()

%>

<html>

<body topmargin="0" leftmargin="0" bgcolor="#ffffff" text="#000000" link="#0000ff" vlink="#0000ff" alink="#0000ff">

<form action="" method="post" name="frmRelatorioIssuesKPI" id="frmRelatorioIssuesKPI">

<LINK href="estilos/sinergia.css" rel=stylesheet>

	<BR>
	<div align="center">
		<center>
			<table border="0" width="90%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;PROJETO </font>
					</td>

					<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;ID </font>
					</td>

      				<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;RESOLUÇÃO</font></td>

      				<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;ANTIGUIDADE</font></td>

    	  			<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;ADIAMENTOS </font></td>

      				<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;REJEICÕES </font></td>
						
	      			<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;IDENTIFICADOR </font></td>

	      			<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;DESIGNADOR</font></td>

    	  			<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;RESPONSÁVEL</font></td>

    		  		<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;REVISOR</font></td>

		      		<td width="12%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-TOP: 1px solid; BORDER-RIGHT: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;APROVADOR </font></td>

	    		</tr>


<% 
if not rsRelatorios.EOF then
	
	do while not rsRelatorios.EOF 
	
		intConta = intConta + 1
%>
			    <tr>

					<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Desc_Projeto")%></td>

					<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("ID")%></font></td>

      				<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("TEMPO_RESOLUCAO")%></font></td>

      				<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
<%      			If rsRelatorios("TEMPO_ANTIGUIDADE") < 1 Then%>
						<font face="Arial" size="1">&nbsp;No prazo</font></td>
<%      			Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("TEMPO_ANTIGUIDADE")%></font></td>
<%      			End If%>

      				<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("ADIAMENTOS")%></font></td>

      				<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("REJEICAO")%></font></td>

	      			<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Identificador")%></font></td>

      				<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Designador")%></font></td>

    	  			<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Responsavel")%></font></td>

		      		<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Revisor")%></font></td>

		      		<td width="33%" align="right" bgcolor="#ffffff" style="BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid; BORDER-RIGHT: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Aprovador")%></font></td>
		    	</tr>
	<% 
		rsRelatorios.MoveNext
		
	loop%>
  			</table>
	  	</center>
	</div>

	<!--hr size="3" color="#000000"-->
<%
else

	response.write "<p><b><font color=#666666 size=2 face=Arial, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
	
end if

'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)

%>
</form>	
</body>
</html>