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
Dim strFrente
Dim strEquipe
Dim vetDados
Dim vetDados_Disc
Dim strProj
Dim strID
Dim strUID
Dim strComplete
Dim strSql
Dim Cont
Dim intProj_Aux

Function VerificaHB()
Dim Dia
Dim Hora

	VerificaHB = true

	Dia = DatePart("w", Now())
	Hora = CDate(FormatDateTime(Now(), 3))

	If Dia = 3 And Hora > CDate("14:00:00") Then
	    VerificaHB = false
	End If

'	If Dia = 4 Then
'	    VerificaHB = false
'	End If

	If Dia = 5 And Hora < CDate("06:00:00") Then
	    VerificaHB = false
	End If

End Function

if VerificaHB then

	If trim(session("Login")) = "" Then
		Response.Redirect ("./Erro.asp?Erro=Sua sessão expirou. Por-favor, logue-se novamente.&Voltar=true&IrPara=./TCP_LOGIN.ASP")

	End if

	strFrente		= Request("slcFrente")
	strEquipe		= Request("slcEquipe")
	strOperacao		= Request("hidOperacao")
	strUsuario      = Request("strUsuario")
	
	'Abrindo uma conexão com o BD
	set conConexao = TCP_AbrirConexaoBD()

	if strOperacao <> "" Then
		
		vetDados = split(Request("hidDados"),";")

		For I = 0 to ubound(vetDados)
			vetDados_Disc = split(vetDados(I),"|")
			strProj = vetDados_Disc(0)
			strID  = vetDados_Disc(1)
			strUID = vetDados_Disc(2)
			strComplete = vetDados_Disc(3)

			strSql = "EXEC SP_INCLUIR_LOG_TAREFAS '" & strUsuario & "', " & strProj & ", " & strUID & ", " & strComplete & " "
			
			conConexao.execute strSql
			
			strSql = "EXEC SP_ATUALIZAR_TAREFAS " & strProj & ", " & strID & ", " & strUID & ", " & strComplete & " "
	
			conConexao.execute strSql
			
		Next
		
		strOperacao = ""

		response.Redirect("./TCP_Confirmacao_Medicao.asp?strUsuario=" & strUsuario & "&strPagina=TCP_Medicao_Detalhe.asp")

	End if

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_INCLUIR_MEDICAO_DETALHE_TB_TEMP"
        
        .Parameters.Refresh
		.Parameters(1).Value = trim(session("Login"))
		
		.Parameters(2).Value = Null
		.Parameters(3).Value = Null

    End With

	cmdResultado.Execute()


'********************************************************************************

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_SUMARIAS"

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = "M"
        
        .Parameters.Refresh

    End With

	set rs1 = Server.CreateObject("ADODB.RecordSet")

	set rs1 = cmdResultado.Execute()

    Set cmdResultado = Server.CreateObject("ADODB.Command")
    
    With cmdResultado

        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_TB_TEMP"
       
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = "M"
	
    End With

	set rs2 = Server.CreateObject("ADODB.RecordSet")

	set rs2 = cmdResultado.Execute()

	Do While Not rs2.EOF and Not rs1.EOF 

		IF rs1("PROJ_ID") = rs2("PROJ_ID") Then

			IF trim(rs1("Task_Outline_num") & ".") = Left(rs2("Task_Outline_num"),LEN(trim(rs1("Task_Outline_num")))+1) Then

				Set cmdResultado = Server.CreateObject("ADODB.Command")
				    
				With cmdResultado
    
				    .ActiveConnection = conConexao
				    .CommandType = 4
					.CommandTimeout = 600
				    .CommandText = "SP_INCLUIR_TB_TEMP"
				    
				    .Parameters.Refresh
					.Parameters(1).Value = rs1("TASK_OUTLINE_NUM")
					.Parameters(2).Value = rs1("TASK_NAME")
					.Parameters(3).Value = rs1("TASK_OUTLINE_LEVEL")
					.Parameters(4).Value = rs1("TASK_BASE_START")
					.Parameters(5).Value = rs1("TASK_BASE_FINISH")
					.Parameters(6).Value = rs1("TASK_START_DATE")
					.Parameters(7).Value = rs1("TASK_FINISH_DATE")
					.Parameters(8).Value = rs1("TASK_PCT_COMP")
					.Parameters(9).Value = rs1("TASK_IS_SUMMARY")
					.Parameters(10).Value = rs1("TASK_IS_EXTERNAL")
					.Parameters(11).Value = rs1("PROJ_ID")
					.Parameters(12).Value = rs1("PROJ_NAME")
					.Parameters(13).Value = rs1("TASK_UID")
					.Parameters(14).Value = rs1("TASK_ID")
					.Parameters(15).Value = rs1("TASK_DUR")
					.Parameters(16).Value = rs1("CLI")
					.Parameters(17).Value = Null
					.Parameters(18).Value = Null
					.Parameters(19).Value = Null
					.Parameters(20).Value = Null
					.Parameters(21).Value = Null
					.Parameters(22).Value = Null
					.Parameters(23).Value = Null
					.Parameters(24).Value = Null					
					.Parameters(25).Value = trim(session("Login"))
					.Parameters(26).Value = "M"



				End With

				cmdResultado.Execute()

				rs1.MoveNext	
		
			ElseIF trim(rs1("Task_Outline_num") & ".") < Left(rs2("Task_Outline_num"),LEN(trim(rs1("Task_Outline_num")))+1) Then

				rs1.MoveNext	
			Else 

				rs2.MoveNext	
			End If

		ElseIF rs1("PROJ_ID") > rs2("PROJ_ID") Then

			rs2.MoveNext

		Else 

			rs1.MoveNext

		End If
				
	Loop

'*********************************************************************************

    Set cmdResultado = Server.CreateObject("ADODB.Command")
        
    With cmdResultado
    
        .ActiveConnection = conConexao
        .CommandType = 4
		.CommandTimeout = 600
        .CommandText = "SP_LISTAR_MEDICAO_DETALHE"
        
        .Parameters.Refresh

		.Parameters(1).Value = trim(session("Login"))
		.Parameters(2).Value = "M"

    End With

	set rs = Server.CreateObject("ADODB.RecordSet")

	set rs = cmdResultado.Execute()
	
	%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmMedicao_Detalhe_TCP" id="frmMedicao_Detalhe_TCP" action="TCP_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="funcoes/Funcoes.js"></SCRIPT>
	<%If Not rs.EOF Then%>
		<p>
		<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema de Medição &gt;<%=rs("PROJ_NAME")%></font></b>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
				  
		  <tr height="17" style="height:12.75pt">
		    <td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
		    <td class="xl27" width="700px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
		    <td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
		    <td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
		    <td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável</font></b></td>

		  </tr>

		<%Cont = 0%>
		<%intProj_Aux = rs("PROJ_ID")%>
		<%Do While Not rs.EOF%>

			<%If rs("PROJ_ID") <> intProj_Aux Then
				
				intProj_Aux = rs("PROJ_ID")%>
				<BR>
				<b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif">Sistema de Medição &gt;<%=rs("PROJ_NAME")%></font></b>
				<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">
					<tr height="17" style="height:12.75pt">
						<td height="17" class="xl27" width=40px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">UID</font></b></td>
						<td class="xl27" width="700px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Nome</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Duração</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Inicio</font></b></td>
						<td class="xl27" width="75px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Fim</font></b></td>
						<td class="xl27" width="40px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">%Comp</font></b></td>
						<td class="xl27" width="150px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Responsável</font></b></td>

					
					</tr>
				  
			<%End If%>
			
			<%If rs("TASK_IS_SUMMARY") = True Then%>
			  <tr height="17" style="height:12.75pt" bgcolor=#ebebeb>
			    <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
					<font face="Arial" size="1">
						<%=rs("TASK_UID")%>&nbsp;
					</font>
				</td>

			    <td class="xl28" style="border: 1 solid #666666" width="700px">
					<font face="Arial" size="1">
						<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
						<%=left("                              ",rs("task_outline_level")) %>
						<%=rs("TASK_NAME")%>
						</div>
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=rs("DURACAO")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
					<font face="Arial" size="1">
						<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
					</font>
				</td>

			    <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
					<font face="Arial" size="1">
						<%=rs("TASK_PCT_COMP")%>&nbsp;
					</font>
				</td>

			    <td class="xl30" align="right" style="border: 1 solid #666666" width="150px">
					<font face="Arial" size="1">
						<%=rs("CLI")%>&nbsp;
					</font>
				</td>



			  </tr>
			  
			<%Else%>

					<input type="hidden" id="hidProj" name="hidProj" value="<%=rs("PROJ_ID")%>">
					<input type="hidden" id="hidId" name="hidId" value="<%=rs("TASK_ID")%>">
					<input type="hidden" id="hidUid" name="hidUid" value="<%=rs("TASK_UID")%>">
					<input type="hidden" id="hidTpLogin" name="hidTpLogin" value="<%=VerificaLogin(rs("LOGIN"), trim(session("Login")))%>">

					<tr height="17" style="height:12.75pt">
					  <td height="17" class="xl22" align="right" style="border: 1 solid #666666" width=40px>
							<font face="Arial" size="1">
								<%=rs("TASK_UID")%>&nbsp;
							</font>
						</td>

					  <td class="xl28" style="border: 1 solid #666666" width="700px">
							<font face="Arial" size="1">
								<div style="margin-left:<%=rs("task_outline_level") * 6 %>">
								<%=left("                              ",rs("task_outline_level")) %>
								<%=rs("TASK_NAME")%>
								</div>
							</font>
						</td>

					<td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
						<font face="Arial" size="1">
							<%=rs("DURACAO")%>&nbsp;
						</font>
					</td>

					  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
							<font face="Arial" size="1">
								<input type="hidden" id="hiddtStart" name="hiddtStart" value="<%=rs("TASK_START_DATE")%>">
								<%=FormatarDataMon(rs("TASK_START_DATE"))%>&nbsp;
							</font>
					  </td>
					  <td class="xl30" align="right" style="border: 1 solid #666666" width="75px">
							<font face="Arial" size="1">
								<input type="hidden" id="hiddtFinish" name="hiddtFinish" value="<%=rs("TASK_FINISH_DATE")%>">
								<%=FormatarDataMon(rs("TASK_FINISH_DATE"))%>&nbsp;
							</font>
					  </td>

					  <td class="xl23" align="right" style="border: 1 solid #666666" width="40px">
							<font face="Arial" size="1">
								<input id="txtComplete" name="txtComplete" value="<%=rs("TASK_PCT_COMP")%>"  maxLength=3 size=3 onblur="ValidaComplete(<%=Cont%>);" onKeyUp="CaracteresValidos('1234567890',this.value);">
								<input type="hidden" id="hidComplete" name="hidComplete" value="<%=rs("TASK_PCT_COMP")%>">
							</font>
						</td>

						<td class="xl30" align="right" style="border: 1 solid #666666" width="150px">
							<font face="Arial" size="1">
								<%=rs("CLI")%>&nbsp;
							</font>
						</td>


					</tr>

					<%Cont = Cont + 1%>

			<%End If

			rs.MoveNext
			
			If Not rs.EOF Then
				If (rs("PROJ_ID") <> intProj_Aux) Then%>
					</table>
				<%End If
			End If%>
			
		<%Loop%>
		
		</table>

		<p align="right">
		<!--<input type="Image" name="cmdConfirmar" value="Confirmar" src="img/000049.gif" align="absmiddle" onclick="Confirmar();">-->
		<!--<input type="button" name="cmdConfirmar" value="Confirmar" onclick="Confirmar();">-->
		<a href="javascript:Confirmar();"><img src="img/000049.gif" width="73" height="16" border="0"></a>
		<BR>
<!--		<table cellspacing="0" cellpadding="0" align=center>
			<tr align=center>
				<td></td>
				<td align=center><a href="./GVI_selecao.asp" align="center"><font color="#666666" size="1" face="Georgia, Times New Roman, Times, serif">Voltar ao Menu</font></a></td>
				<td></td>
			</tr>
		</table>-->
		<img src="img/_0.gif" width="2" height="2">
		<hr>
		<input type="hidden" id="hidCont" name="hidCont" value="<%=Cont%>">
		<input type="hidden" id="slcFrente" name="slcFrente" value="<%=strFrente%>">
		<input type="hidden" id="slcEquipe" name="slcEquipe" value="<%=strEquipe%>">

		<input type="hidden" id="hidDados" name="hidDados" value="">
		<input type="hidden" id="hidOperacao" name="hidOperacao" value="">
		<input type="hidden" id="strUsuario" name="strUsuario" value="<%=trim(session("Login"))%>">

	<%else

		response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
		
	end if

	'Fechando a conexão com o BD
	call FecharConexaoBD(conConexao)

Else
	response.write "<p><b><center><font color=#666666 size=2 face=Georgia, Times New Roman, Times, serif>&nbsp;&nbsp;Medição só pode ser atualizada apartir de quinta 06:00 AM</font></center></b></p>"
End If%>

<SCRIPT language=JavaScript>

function ValidaComplete(num)
{
	var intCont	= (document.frmMedicao_Detalhe_TCP.hidCont.value - 1)

	if (intCont == 0)
	{

		var conteudo = document.frmMedicao_Detalhe_TCP.txtComplete.value
		var DtFim = document.frmMedicao_Detalhe_TCP.hiddtFinish.value
		var DtInicio = document.frmMedicao_Detalhe_TCP.hiddtStart.value
		var TpLogin = document.frmMedicao_Detalhe_TCP.hidTpLogin.value

		if (jTrim(conteudo) == "")
		{
			alert("O conteúdo do campo %Comp invalido. deve ser numérico.");
			document.frmMedicao_Detalhe_TCP.txtComplete.value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
			document.frmMedicao_Detalhe_TCP.txtComplete.focus();
			return false;
		}

		if (parseInt(conteudo) < parseInt(document.frmMedicao_Detalhe_TCP.hidComplete.value))
		{
			alert("O conteúdo do campo %Comp invalido. deve ser maior que o atual.");
			document.frmMedicao_Detalhe_TCP.txtComplete.value = document.frmMedicao_Detalhe_TCP.hidComplete.value;
			document.frmMedicao_Detalhe_TCP.txtComplete.focus();
			return false;
		}

	/*	if (parseInt(conteudo) != parseInt(document.frmMedicao_Detalhe_TCP.hidComplete(num).value))
		{
			if ((TpLogin == 'P') && (conteudo > 50))
			{
				alert("O conteúdo do campo %Comp invalido para o usuário. deve ser menor ou igual a 50.");
				document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
				document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
				return false;
			}

			if ((TpLogin == 'S') && (parseInt(document.frmMedicao_Detalhe_TCP.hidComplete(num).value) < 50))
			{
				alert("O conteúdo do campo %Comp invalido para o usuário. Usuário só pode medir quando %Comp igual ou maior que 50.");
				document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
				document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
				return false;
			}
		}
	*/	
		if (parseInt(document.frmMedicao_Detalhe_TCP.hidProj.value)!= 8)
		{
			if (DtInicio == DtFim)
			{ 
				if ((conteudo != 0) && (conteudo != 100))
				{
					alert("O conteúdo do campo %Comp invalido. deve ser (0 ou 100).");
					document.frmMedicao_Detalhe_TCP.txtComplete.value = document.frmMedicao_Detalhe_TCP.hidComplete.value;
					document.frmMedicao_Detalhe_TCP.txtComplete.focus();
					return false;
				}
			}
			else
			{
				if ((conteudo != 0) && (conteudo != 25) && (conteudo != 50) && (conteudo != 75) && (conteudo != 100))
				{
					alert("O conteúdo do campo %Comp invalido. deve ser (0,25,50,75 ou 100).");
					document.frmMedicao_Detalhe_TCP.txtComplete.value = document.frmMedicao_Detalhe_TCP.hidComplete.value;
					document.frmMedicao_Detalhe_TCP.txtComplete.focus();
					return false;
				}
			}

		}
		else
		{
			if (parseInt(conteudo) < 0)
			{
				alert("O conteúdo do campo %Comp invalido. deve ser maior ou igual a 0.");
				document.frmMedicao_Detalhe.txtComplete.value = document.frmMedicao_Detalhe.hidComplete.value;
				document.frmMedicao_Detalhe.txtComplete.focus();
				return false;
			}

			if (parseInt(conteudo) > 100)
			{
				alert("O conteúdo do campo %Comp invalido. deve ser menor ou igual a 100.");
				document.frmMedicao_Detalhe.txtComplete.value = document.frmMedicao_Detalhe.hidComplete.value;
				document.frmMedicao_Detalhe.txtComplete.focus();
				return false;
			}
		}

	}
	else
	{
		var conteudo = document.frmMedicao_Detalhe_TCP.txtComplete(num).value
		var DtFim = document.frmMedicao_Detalhe_TCP.hiddtFinish(num).value
		var DtInicio = document.frmMedicao_Detalhe_TCP.hiddtStart(num).value
		var TpLogin = document.frmMedicao_Detalhe_TCP.hidTpLogin(num).value

		if (jTrim(conteudo) == "")
		{
			alert("O conteúdo do campo %Comp invalido. deve ser numérico.");
			document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
			document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
			return false;
		}

		if (parseInt(conteudo) < parseInt(document.frmMedicao_Detalhe_TCP.hidComplete(num).value))
		{
			alert("O conteúdo do campo %Comp invalido. deve ser maior que o atual.");
			document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
			document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
			return false;
		}

	/*	if (parseInt(conteudo) != parseInt(document.frmMedicao_Detalhe_TCP.hidComplete(num).value))
		{
			if ((TpLogin == 'P') && (conteudo > 50))
			{
				alert("O conteúdo do campo %Comp invalido para o usuário. deve ser menor ou igual a 50.");
				document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
				document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
				return false;
			}

			if ((TpLogin == 'S') && (parseInt(document.frmMedicao_Detalhe_TCP.hidComplete(num).value) < 50))
			{
				alert("O conteúdo do campo %Comp invalido para o usuário. Usuário só pode medir quando %Comp igual ou maior que 50.");
				document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
				document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
				return false;
			}
		}
	*/	
		if (parseInt(document.frmMedicao_Detalhe_TCP.hidProj(num).value)!= 8)
		{
			if (DtInicio == DtFim)
			{ 
				if ((conteudo != 0) && (conteudo != 100))
				{
					alert("O conteúdo do campo %Comp invalido. deve ser (0 ou 100).");
					document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
					document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
					return false;
				}
			}
			else
			{
				if ((conteudo != 0) && (conteudo != 25) && (conteudo != 50) && (conteudo != 75) && (conteudo != 100))
				{
					alert("O conteúdo do campo %Comp invalido. deve ser (0,25,50,75 ou 100).");
					document.frmMedicao_Detalhe_TCP.txtComplete(num).value = document.frmMedicao_Detalhe_TCP.hidComplete(num).value;
					document.frmMedicao_Detalhe_TCP.txtComplete(num).focus();
					return false;
				}
			}

		}
		else
		{
			if (parseInt(conteudo) < 0)
			{
				alert("O conteúdo do campo %Comp invalido. deve ser maior ou igual a 0.");
				document.frmMedicao_Detalhe.txtComplete(num).value = document.frmMedicao_Detalhe.hidComplete(num).value;
				document.frmMedicao_Detalhe.txtComplete(num).focus();
				return false;
			}

			if (parseInt(conteudo) > 100)
			{
				alert("O conteúdo do campo %Comp invalido. deve ser menor ou igual a 100.");
				document.frmMedicao_Detalhe.txtComplete(num).value = document.frmMedicao_Detalhe.hidComplete(num).value;
				document.frmMedicao_Detalhe.txtComplete(num).focus();
				return false;
			}
		}
	}
}

function Confirmar()
{
	var intCont	= (document.frmMedicao_Detalhe_TCP.hidCont.value - 1)
	var strAux = ""

	document.frmMedicao_Detalhe_TCP.style.cursor = "wait";
			
	document.frmMedicao_Detalhe_TCP.hidDados.value = "";
	document.frmMedicao_Detalhe_TCP.hidOperacao.value = "";
	
	if (intCont == 0)
	{
		for(var i = 0; i <= intCont;i++)
		{
			if (document.frmMedicao_Detalhe_TCP.txtComplete.value != document.frmMedicao_Detalhe_TCP.hidComplete.value)
			{
				if (strAux == "")
				{
					strAux = document.frmMedicao_Detalhe_TCP.hidProj.value + "|" + document.frmMedicao_Detalhe_TCP.hidId.value + "|" + document.frmMedicao_Detalhe_TCP.hidUid.value + "|" + document.frmMedicao_Detalhe_TCP.txtComplete.value;
				}	
				else
				{
					strAux = strAux + ";" + document.frmMedicao_Detalhe_TCP.hidProj.value + "|" + document.frmMedicao_Detalhe_TCP.hidId.value + "|" + document.frmMedicao_Detalhe_TCP.hidUid.value + "|" + document.frmMedicao_Detalhe_TCP.txtComplete.value;
				}
			}
		}
	}
	else
	{
		for(var i = 0; i <= intCont;i++)
		{
			if (document.frmMedicao_Detalhe_TCP.txtComplete(i).value != document.frmMedicao_Detalhe_TCP.hidComplete(i).value)
			{
				if (strAux == "")
				{
					strAux = document.frmMedicao_Detalhe_TCP.hidProj(i).value + "|" + document.frmMedicao_Detalhe_TCP.hidId(i).value+ "|" + document.frmMedicao_Detalhe_TCP.hidUid(i).value + "|" + document.frmMedicao_Detalhe_TCP.txtComplete(i).value;
				}	
				else
				{
					strAux = strAux + ";" + document.frmMedicao_Detalhe_TCP.hidProj(i).value + "|" + document.frmMedicao_Detalhe_TCP.hidId(i).value + "|" + document.frmMedicao_Detalhe_TCP.hidUid(i).value + "|" + document.frmMedicao_Detalhe_TCP.txtComplete(i).value;
				}
			}
		}
	}

	if (strAux != "")
	{
		document.frmMedicao_Detalhe_TCP.hidDados.value = strAux;
		document.frmMedicao_Detalhe_TCP.hidOperacao.value = 'A';
		document.frmMedicao_Detalhe_TCP.style.cursor = "";
		document.frmMedicao_Detalhe_TCP.submit();
	}
	else
	{
		alert("Nenhum registro foi alterado");
		document.frmMedicao_Detalhe_TCP.style.cursor = "";
	}
}

</SCRIPT>
</FORM>
</body>
</html>
