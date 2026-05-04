<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs

	strUsuario = Request("strUsuario")
	strPagina  = Request("strPagina")

	'Abrindo uma conexão com o BD
	set conConexao = TCP_AbrirConexaoBD()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_CONFIRMACAO_MEDICAO"
        
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
	<FORM name="frmConfirmacao_Medicao" id="frmConfirmacao_Medicao" action="" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<table>
			<TR>
				<TD  width=60%>&nbsp;</TD>
				<TD  align=center>
					<b><font color=Red size="4" face="Georgia, Times New Roman, Times, serif">Confirmação de Medição</font></b>
				</TD>
<!--				<TD  width=20%></TD>-->
			</TR>
		</Table>
		<BR>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Log de Medição de <%=rs("Cod_Usuario")%> - <%=rs("Nome")%>&nbsp;&nbsp;&nbsp;&nbsp;Data:&nbsp;<%=rs("DATA")%></font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Projeto</font></b></td>
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="425px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Data/Hora da Atualização</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		  </tr>

		<%Do While Not rs.EOF%>

				<tr height="17" style="height:12.75pt">

				  <td class="xl28" style="border: 1 solid #666666" width="200px">
						<font face="Arial" size="1">
							<%=rs("PROJ_NAME")%>&nbsp;
						</font>
					</td>

				  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
						<font face="Arial" size="1">
							<%=rs("TASK_UID")%>&nbsp;
						</font>
					</td>

				  <td class="xl28" style="border: 1 solid #666666" width="550px">
						<font face="Arial" size="1">
							<%=rs("TASK_NAME")%>&nbsp;
						</font>
					</td>

				  <td class="xl30" align=center style="border: 1 solid #666666" width="150px">
						<font face="Arial" size="1">
							<%=rs("Data_Atualizacao")%>&nbsp;
						</font>
				  </td>

				  <td class="xl30" align=center style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=rs("TASK_START_DATE")%>&nbsp;
						</font>
				  </td>
				  <td class="xl30" align=center style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=rs("TASK_FINISH_DATE")%>&nbsp;
						</font>
				  </td>

				  <td class="xl23" align=center style="border: 1 solid #666666" width="40px">
						<font face="Arial" size="1">
							<%=rs("TASK_PCT_COMP")%>&nbsp;
						</font>
					</td>
				</tr>

			<%rs.MoveNext%>
			
		<%Loop%>
		
		</table>

		<p align="right">
		<a href="javascript:Imprimir();">Imprimir</a>&nbsp;&nbsp;
		<a href="javascript:Voltar();"><img src="img/000024.gif" width="73" height="16" border="0"></a>
		<BR>
		<img src="img/_0.gif" width="2" height="2">
		</p>

		<table>
			<TR>
				<TD  width=63%>&nbsp;</TD>
				<TD  align=center>
					<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Obrigado por realizar a medição!!!</font></b>
				</TD>
			</TR>
		</Table>

		<hr>
		<input type="hidden" id="strPagina" name="strPagina" value="<%=trim(strPagina)%>">

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)
%>

<SCRIPT language=JavaScript>

function Imprimir()
{
	window.print();
}

function Voltar()
{

	document.frmConfirmacao_Medicao.action = document.frmConfirmacao_Medicao.strPagina.value;
	document.frmConfirmacao_Medicao.submit();
}

</SCRIPT>
</FORM>
</body>
</html>