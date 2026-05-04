<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strProjeto
Dim strUID
Dim strID
Dim strSql
Dim strOperacao
Dim strDataIni
Dim strDataFim
Dim strNome
Dim strResposavel
Dim strCenario
Dim strTransacao
Dim strSequencia
Dim strLogin
Dim strChave
Dim rs
Dim strMotivo

Function FormatarDataSQL(strData)
	FormatarDataSQL = mid(strData,7,4) + mid(strData,4,2) + mid(strData,1,2)
End Function

	strNome	   = Request("strNome")
	strProjeto = Request("strProjeto")
	strUID	   = Request("strUID")
	strID	   = Request("strID")
	strDataIni = Request("strDataIni")
	strDataFim = Request("strDataFim")
	strLogin   = Request("strLogin")
	strOperacao	= Request("hidOperacao")
	strTipo		= Request("hidTipo")
	strMotivo	= Request("txtMotivo")


	strRedirect	= "./GVI_Excluir_Tarefas.asp?strProjeto=" & strProjeto & "þstrUID=" & strUID
	
	if trim(session("Usuario"))="" then
		response.Redirect("./LOGIN.ASP?hidOrigem=" & strRedirect)
	end if


	strChave   = trim(session("Usuario"))
	strLogin   = trim(session("Usuario"))
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

%>
<html>
<head>

</head>
<body>
<LINK href="estilos/sinergia.css" rel=stylesheet>
<form name="frmExcluirTarefa_GVI" ID="frmExcluirTarefa_GVI" action="GVI_Excluir_Tarefas.asp" method="post">

<%
	if strOperacao = "" Then

		strSql = "SP_LISTAR_TAREFAS " & strProjeto & " , " & strUID & ""

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao

		strID	   = rs("TASK_ID")
		strDataIni = rs("DATA_INICIO")
		strDataFim = rs("DATA_FIM")

%>		


		<input type="hidden" id="hidOperacao" name="hidOperacao" >
		<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
		<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">
		<input type="hidden" id="strID" name="strID" value="<%=rs("TASK_ID")%>">
		<input type="hidden" id="strDataIni" name="strDataIni" value="<%=rs("DATA_INICIO")%>" >
		<input type="hidden" id="strDataFim" name="strDataFim" value="<%=rs("DATA_FIM")%>" >

		<table width="100%" border="1" cellspacing="0" cellpadding="0" bgColor=silver>
			<tr>
				<td colspan=4 bgcolor="#c0c0c0"><font color=white size="2" face="Arial, Times New Roman, Times, serif">Nome da tarefa:</font></td>
			</tr>
			<tr>
				<td colspan=4 bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=rs("TASK_NAME")%></td>
			</tr>
			<tr>
				<td bgcolor="#c0c0c0"><font color=white size="2" face="Arial, Times New Roman, Times, serif">UID</font></td>
				<td bgcolor="#c0c0c0"><font color=white size="2" face="Arial, Times New Roman, Times, serif">ID</font></td>
				<td bgcolor="#c0c0c0"><font color=white size="2" face="Arial, Times New Roman, Times, serif">Data Inicio</font></td>
				<td bgcolor="#c0c0c0"><font color=white size="2" face="Arial, Times New Roman, Times, serif">Data Fim</font></td>
			</tr>
			<tr>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=strUID%></font></td>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=rs("TASK_ID")%></font></B></td>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=rs("DATA_INICIO")%></font></B></td>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=rs("DATA_FIM")%></font></B></td>
			</tr>
		</table>
		<BR><BR><BR>
		
<%
		strSql = "SELECT TASK_NAME FROM TarefasExcluidas_GVI A, MSP_TASKS B WHERE A.proj = B.PROJ_ID AND A.UID = B.TASK_UID AND PROJ = " & strProjeto & " AND UID = " & strUID 

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao

		If not rs.eof Then

%>
			<BR><BR><BR>

			<input type="hidden" id="hidTipo" name="hidTipo" value="I">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width="1" height="1" bgcolor="#003366"><IMG height=1 src="img/_0.gif" width=1></td>
					<B><FONT face="Georgia, Times New Roman, Times, serif" color=#666666 size=2>Para cancelar a exclusão da tarefa, aperte confirmar.</FONT></B>
				</tr>
			</table>
<%
		Else

%>

			<table width="100%" border="1" cellspacing="0" cellpadding="0">
				<tr>
					<td bgcolor="#c0c0c0" width=150px>
						<font color=white size="2" face="Arial, Times New Roman, Times, serif">
							Motivo da Exclusão:&nbsp;&nbsp;
						</font>
					</td>
					<td width=400px>
						<font face="Arial, Times New Roman, Times, serif">
							<input id="txtMotivo" name="txtMotivo" value="" size=73>
						</font>
					</td>
				</tr>
			</table>
		
			<BR><BR><BR>

			<input type="hidden" id="hidTipo" name="hidTipo" value="E">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width="1" height="1" bgcolor="#003366"><IMG height=1 src="img/_0.gif" width=1></td>
					<B><FONT face="Georgia, Times New Roman, Times, serif" color=#666666 size=2>Para excluir a tarefa, aperte confirmar.</FONT></B>
				</tr>
			</table>
<%

		End If
	Else

		If strTipo = "I" Then

			strSql = "EXEC SP_CANCELA_EXCLUSAO_GVI " & strProjeto & ", " & strID & ", " &	strUID & "" 
			         
			conConexao.execute strSql

			strSql = "EXEC SP_INCLUIR_LOG_TAREFAS_EXCLUIDAS_GVI '" & strLogin & "', " & strProjeto & ", " &	strUID & ", 'C', '" & Trim(strChave) & "','" & strMotivo & "'"

			conConexao.execute strSql%>

<%		ElseIf strTipo = "E" Then

			strSql = "EXEC SP_EXCLUIR_TAREFAS_GVI " & strProjeto & ", " & strID & ", " & strUID  & ", '" & Trim(strChave) & "'"

			conConexao.execute strSql

			strSql = "EXEC SP_INCLUIR_LOG_TAREFAS_EXCLUIDAS_GVI '" & strLogin & "', " & strProjeto & ", " &	strUID & ", 'E', '" & Trim(strChave) & "','" & strMotivo & "'"

			conConexao.execute strSql%>


	  <%End If%>

		<SCRIPT language=JavaScript>
/*			top.opener.document.frmMedicao_Detalhe.submit()*/
			this.close()
		</SCRIPT>

<%	End If%>

<TABLE style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; WIDTH: 609px; BORDER-BOTTOM: 1px solid; HEIGHT: 21px" 
borderColor=white cellSpacing=0 cellPadding=0 border=0>
  
  <TR>
    <TD></TD>
    <TD align=right><A href="javascript:Confirmar();"><IMG 
      src="img/000049.gif" align=absMiddle border=0>
				<A href="javascript:this.close();" ><IMG height=16 src="img/000023.gif" width=73 align=absMiddle border=0></A> </TD></TR></TABLE></P>


	<input type="hidden" id="strLogin" name="strLogin" value="<%=strLogin%>">

</form>

<SCRIPT language=JavaScript>

function Confirmar()
{
	if (document.frmExcluirTarefa_GVI.hidTipo.value == 'E')
	{
		if (document.frmExcluirTarefa_GVI.txtMotivo.value == '')
		{
			alert("Campo motivo da exclusão obrigatório!");
		}
		else
		{
			document.frmExcluirTarefa_GVI.hidOperacao.value = 'A';
			document.frmExcluirTarefa_GVI.submit();
		}
	}
	else
	{
		document.frmExcluirTarefa_GVI.hidOperacao.value = 'A';
		document.frmExcluirTarefa_GVI.submit();
	}
}

</SCRIPT>
</body>
</html>

