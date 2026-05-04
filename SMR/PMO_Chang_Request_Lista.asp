<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs
Dim cmdResultado

	strUsuario = Request("strUsuario")

	If trim(strUsuario) = "" Then
		response.Redirect("./LOGIN.ASP?hidOrigem=./PMO_Chang_Request_Selecao.asp")
	End if
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_CHANG_REQUEST_ALT"

        .Parameters.Refresh
		.Parameters(1).Value = trim(strUsuario)

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmChang_Request_Detalhe_PMO" id="frmChang_Request_Detalhe_PMO" action="PMO_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>

	<%If Not rs.EOF Then%>
		<p>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
		  <tr height="17" style="height:12.75pt">
		    <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto</font></b></td>
		    <td height="17" class="xl27" width=100px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="350px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="100px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Ação</font></b></td>				
		    <td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Novo Inicio</font></b></td>
		    <td class="xl27" width="95px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Novo Fim</font></b></td>
		    <td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Novo %Comp</font></b></td>
		  </tr>

		<%Do While Not rs.EOF%>
			
			  <tr height="17" style="height:12.75pt">
			    <td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
					<font face="Arial" size="1">
						<%=rs("PROJ_NAME")%>&nbsp;
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="100px" align=center>
					<font face="Arial" size="1">
						<%=rs("UID")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=left style="border: 1 solid #666666" width="350px">
					<font face="Arial" size="1">
						<%=rs("Nome")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=left style="border: 1 solid #666666" width="100px">
					<font face="Arial" size="1">
						<%=rs("Acao")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("Dt_Inicio_Nov"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align=center style="border: 1 solid #666666" width="95px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("Dt_Final_Nov"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					<font face="Arial" size="1">
						<%=rs("Perc_Comp_Nov")%>&nbsp;
					</font>
				</td>
			  </tr>

			<%rs.MoveNext
			
		  Loop%>
		
		</table>

		<p align="right">
		<a href="javascript:history.go(-1);" ><img src="img/000024.gif" width="73" height="16" border="0" align="absmiddle" tabindex="4"></a>
		<BR>
		<img src="img/_0.gif" width="2" height="2">
		<hr>

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)
%>



</FORM>
</body>
</html>
