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
Dim strMotivo
Dim rs

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

	strMotivo	= Request("txtMotivo")

	strPlano  = Trim(Request("strPlano"))
	strModulo = Trim(Request("strModulo"))

	strOperacao	= Request("hidOperacao")
	strTipo		= Request("hidTipo")

	
	'Abrindo uma conexão com o BD
	set conConexao = CUT_AbrirConexaoBD()

%>
<html>
<head>

</head>
<body>
<LINK href="estilos/sinergia.css" rel=stylesheet>
<form name="frmExcluirTarefa" ID="frmExcluirTarefa" action="CUT_Excluir_Tarefas.asp" method="post">
	<input type="hidden" id="hidOperacao" name="hidOperacao" >
	<input type="hidden" id="strProjeto" name="strProjeto" value="<%=strProjeto%>">
	<input type="hidden" id="strUID" name="strUID" value="<%=strUID%>">
	<input type="hidden" id="strID" name="strID" value="<%=strID%>">
	<input type="hidden" id="strDataIni" name="strDataIni" value="<%=strDataIni%>" >
	<input type="hidden" id="strDataFim" name="strDataFim" value="<%=strDataFim%>" >

<%
	if strOperacao = "" Then

		strSql = "SELECT TASK_NAME FROM MSP_TASKS WHERE PROJ_ID = " & strProjeto & " AND TASK_UID = " & strUID 

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao
%>		
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
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=strID%></font></B></td>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=strDataIni%></font></B></td>
				<td bgcolor="white"><font color=Black size="1" face="Arial, Times New Roman, Times, serif"><%=strDataFim%></font></B></td>
			</tr>
		</table>
		<BR><BR><BR>
		<BR><BR><BR>
		
<%
		strSql = "SELECT TASK_NAME FROM TarefasExcluidas A, MSP_TASKS B WHERE A.proj = B.PROJ_ID AND A.UID = B.TASK_UID AND PROJ = " & strProjeto & " AND UID = " & strUID 

		set rs = Server.CreateObject("ADODB.RecordSet")

		rs.OPEN STRSQL, conConexao

		If not rs.eof Then

%>
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
		
			<BR>

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
			strSql = "EXEC SP_CANCELA_EXCLUSAO " & strProjeto &_
			         ", " & strID & ", " &	strUID 
			         
			conConexao.execute strSql

			strSql = "EXEC SP_INCLUIR_LOG_TAREFAS_EXCLUIDAS '" & strLogin & "', " & strProjeto &_
			         ", " &	strUID & ", 'C', null" 

			conConexao.execute strSql%>

<%		ElseIf strTipo = "E" Then
			strSql = "EXEC SP_EXCLUIR_TAREFAS " & strProjeto &_
			         ", " & strID & ", " &	strUID & ", '" & strLogin & "', '" & strMotivo & "' "

			conConexao.execute strSql

			strSql = "EXEC SP_INCLUIR_LOG_TAREFAS_EXCLUIDAS '" & strLogin & "', " & strProjeto &_
			         ", " &	strUID & ", 'E', '" & strMotivo & "' "

			conConexao.execute strSql%>


<%		End If%>


		<SCRIPT language=JavaScript>
			top.opener.document.frmModificacao_Detalhe.submit()
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

	if (document.frmExcluirTarefa.hidTipo.value == 'E')
	{
		if (document.frmExcluirTarefa.txtMotivo.value == '')
		{
			alert("Campo motivo da exclusão obrigatório!");
		}
		else
		{
			document.frmExcluirTarefa.hidOperacao.value = 'A';
			document.frmExcluirTarefa.submit();
		}
	}
	else
	{
		document.frmExcluirTarefa.hidOperacao.value = 'A';
		document.frmExcluirTarefa.submit();
	}
	
/*	document.frmExcluirTarefa.hidOperacao.value = 'A';
	document.frmExcluirTarefa.submit();*/

}

</SCRIPT>
</body>
</html>

