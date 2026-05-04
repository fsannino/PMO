<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim rs
Dim strSql
Dim strLogin

	strLogin = trim(Request("strLogin"))

	'Abrindo uma conexão com o BD
	set conConexao = CUT_AbrirConexaoBD()

	strSql = "SP_LISTAR_LOG_TAREFAS_ALTERADAS_MODULO '" & strLogin & "'"

	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.OPEN STRSQL, conConexao

%>
<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmListarAlterarDatas" ID="frmListarAlterarDatas" method="post">
<center>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
  </tr>
</table>

<BR>

<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">

  <%If Not rs.EOF Then%>
		
		<tr height="17" style="height:12.75pt">
			<td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
			<td class="xl27" width="340px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
			<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio Solicitado</font></b></td>
			<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim Solicitado</font></b></td>
			<td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Chave</font></b></td>
			<td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
			  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data da Atualização</font></b></td>
		</tr>
				
		<%Do While Not rs.EOF%>

			<tr height="17" style="height:12.75pt">
				
				<td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
					<font face="Arial" size="1">
						<%=rs("UID")%>
					</font>
				</td>

				<td class="xl28" style="border: 1 solid #666666" width="340px">
					<font face="Arial" size="1">
						<%=rs("TASK_NAME")%>&nbsp;
					</font>
				</td>
				
				<td class="xl28" style="border: 1 solid #666666" width="80px" align="center" >
					<font face="Arial" size="1">
						<%=FormatDateTime(rs("DataInicio"),2)%>&nbsp;
					</font>
				</td>
				<td class="xl28" style="border: 1 solid #666666" width="80px" align="center" >
					<font face="Arial" size="1">
						<%=FormatDateTime(rs("DataFim"),2)%>&nbsp;
					</font>
				</td>
				<td class="xl28" style="border: 1 solid #666666" width="50px" align="center" >
					<font face="Arial" size="1">
						<%=rs("Login")%>&nbsp;
					</font>
				</td>
				<td class="xl28" style="border: 1 solid #666666" width="150px" align=center >
					<font face="Arial" size="1">
						<%=rs("DataAtualizacao")%>&nbsp;
					</font>
				</td>
			</tr>
					
			<%rs.MoveNext%>
		<%Loop%>
		
		</table>

		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
		    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
		  </tr>
		</table>

		<BR>

	<%Else
		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
    End If%>

</center>
</form>
</body>
</html>
