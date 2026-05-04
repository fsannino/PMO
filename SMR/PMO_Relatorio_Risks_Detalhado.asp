<!--#include file="./funcoes/Funcoes.inc"-->
<!--#include file="./head_imp.asp"-->

<% 
Dim strEquipe
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
	FormatarDataSQL =  mid(strData,4,2) & "/" & mid(strData,1,2) & "/" & mid(strData,7,4) 
End Function

'Recuperando dados da página de filtro
strEquipe = Request("strEquipe")
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
    

If Trim(strUsuario) = "" Then

	If trim(strCarteira) <> "" Then

		With cmdResultado
		    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 600
		    .CommandText = "SP_LISTAR_RISKS_DETALHE_CARTEIRA"
				    
		    .Parameters.Refresh

			'.Parameters(1).Value = FormatarDataSQL(strData)
			.Parameters(1).Value = strData
			.Parameters(2).Value = trim(strCarteira)

			If trim(strEquipe) <> "" then
				.Parameters(3).Value = trim(strEquipe)
			End if

			If trim(strComite) <> "" then
				.Parameters(4).Value = trim(strComite)
			End if

		End With

	Else

		With cmdResultado
		    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 600
		    .CommandText = "SP_LISTAR_RISKS_DETALHE"
				    
		    .Parameters.Refresh
			'.Parameters(1).Value = FormatarDataSQL(strData)
			.Parameters(1).Value = strData
			.Parameters(2).Value = trim(strTipo)

			If trim(strEquipe) <> "" then
				.Parameters(3).Value = trim(strEquipe)
			End if

			If trim(strCriticidade) <> "" then
				.Parameters(4).Value = trim(strCriticidade)
			End if
				    
		End With
	
	End If

Else

	With cmdResultado
	    
	    .ActiveConnection = conConexao
	    .CommandType = 4
		.CommandTimeout = 600
	    .CommandText = "SP_LISTAR_RISKS_DETALHE_USUARIO"

	    .Parameters.Refresh

'		.Parameters(1).Value = FormatarDataSQL(strData)
		.Parameters(1).Value = strData
		.Parameters(2).Value = trim(strTipo)
		.Parameters(3).Value = trim(strUsuario)
			    
	End With

End If
	
	
set rsRelatorios = Server.CreateObject("ADODB.RecordSet")

set rsRelatorios = cmdResultado.Execute()

%>

<html>

<body topmargin="0" leftmargin="0" bgcolor="#ffffff" text="#000000" link="#0000ff" vlink="#0000ff" alink="#0000ff">

<form action="" method="post" name="frmRelatorioIssuesDetalhado" id="frmRelatorioIssuesDetalhado">

<LINK href="estilos/sinergia.css" rel=stylesheet>

<% 
if not rsRelatorios.EOF then
	
	do while not rsRelatorios.EOF 
	
		intConta = intConta + 1
	
		If cDate(FormatDateTime(rsRelatorios("Data_Limite"),2)) >= cDate(FormatDateTime(Now,2)) Then
			strTipo = "No prazo"
		ElseIf cDate(FormatDateTime(rsRelatorios("Data_Limite"),2)) >= cDate(FormatDateTime(Now - 10,2)) Then
			strTipo = "Atrasados (menos de 10 dias)"
		Else
			strTipo = "Atrasados (mais de 10 dias)"
		End If

		If (rsRelatorios("Impacto") = "Baixo" and strTipo = "No prazo") or _
		   (rsRelatorios("Impacto") = "Baixo" and strTipo = "Atrasados (menos de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Medio" and strTipo = "No prazo") Then
			strCarteira = "Nivel 3"
		End If
		If (rsRelatorios("Impacto") = "Baixo" and strTipo = "Atrasados (mais de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Medio" and strTipo = "Atrasados (menos de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Medio" and strTipo = "Atrasados (mais de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Alto" and strTipo = "No prazo") or _
		   (rsRelatorios("Impacto") = "Alto" and strTipo = "Atrasados (menos de 10 dias)") Then
			strCarteira = "Nivel 2"
		End If
		If (rsRelatorios("Impacto") = "Alto" and strTipo = "Atrasados (mais de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Critico" and strTipo = "No prazo") or _
		   (rsRelatorios("Impacto") = "Critico" and strTipo = "Atrasados (menos de 10 dias)") or _
		   (rsRelatorios("Impacto") = "Critico" and strTipo = "Atrasados (mais de 10 dias)") Then
			strCarteira = "Nivel 1"
		End If

%>
	<BR>
	<div align="center">
		<center>
			<table border="0" width="90%" cellspacing="0" cellpadding="0" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid">
				<tr>
					<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;Equipe: </font>
					</td>

					<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Desc_Eqp")%></td>

					<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;NOME: </font>
					</td>

					<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Name")%></font></td>

	    		</tr>
			    <tr>
      				<td width="12%" align="left" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;COMITÊ:</font></td>

      				<td width="33%" align="left" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("desc_comite")%></font></td>

      				<td width="12%" align="left" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;ID:</font></td>

      				<td width="33%" align="left" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("ID")%></font></td>

    			</tr>
	    		<tr>
    	  			<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;ABERTO EM: </font>
					</td>
	      			<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;
						<%
						if (rsRelatorios("Aberto_em") <> cdate(0)) and (rsRelatorios("Aberto_em") <> cdate("01/01/1900")) then
							response.write rsRelatorios("Aberto_em")
						else
							response.write "&nbsp;"
						end if
						%>
						</td>
      				<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;DATA LIMITE: </font>
					</td>
      				<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;
						<%
						if (rsRelatorios("Data_Limite") <> cdate(0)) and (rsRelatorios("Data_Limite") <> cdate("01/01/1900")) then
							response.write rsRelatorios("Data_Limite")
						else
							response.write "&nbsp;"
						end if
						%>
						</td>
    			</tr>
    			<tr>
	      			<td width="12%" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;IMPACTO:</font></td>
      				<td width="33%" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Impacto")%></td>
      				<td width="12%" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;AREA DE IMPACTO: </font></td>
      				<td width="33%" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Area_Impacto")%></td>
    			</tr>
    			<tr>
      				<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;STAT. DE RES.: </font></td>
      				<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=strTipo%></font></td>
    	  			<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;CARTEIRA: </font></td>
    	  			<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=strCarteira%></font></td>
	    		</tr>
	    		<tr>
    	  			<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;WORKFLOW: </font></font></td>
	      			<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Workflow")%></td>
	      			<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;IDENTIFICADOR: </font></td>
	      			<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
	      			<%If rsRelatorios("Workflow") = "Identificador" Then%>
						<B><font color=red face="Arial" size="2">&nbsp;<%=rsRelatorios("Identificador")%></font></B></td>
	      			<%Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Identificador")%></font></td>
	      			<%End If%>
	    		</tr>
    			<tr>
	      			<td width="12%" align="left" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;DESIGNADOR:</font></td>
      				<td width="33%" align="left" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
	      			<%If rsRelatorios("Workflow") = "Designador" Then%>
						<B><font color=red face="Arial" size="2">&nbsp;<%=rsRelatorios("Designador")%></font></B></td>
	      			<%Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Designador")%></font></td>
	      			<%End If%>
    	  			<td width="12%" align="left" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
						<font face="Arial" size="1">&nbsp;RESPONSÁVEL:</font></td>
    	  			<td width="33%" align="left" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" height="15">
	      			<%If rsRelatorios("Workflow") = "Responsável" Then%>
						<B><font color=red face="Arial" size="2">&nbsp;<%=rsRelatorios("Responsavel")%></font></B></td>
	      			<%Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Responsavel")%></font></td>
	      			<%End If%>
	    		</tr>
		    	<tr>
    		  		<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" height="15">
						<font face="Arial" size="1">
						<font size="1">&nbsp;REVISOR:</font> </font>
					</td>
		      		<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
	      			<%If rsRelatorios("Workflow") = "Revisor" Then%>
						<B><font color=red face="Arial" size="2">&nbsp;<%=rsRelatorios("Revisor")%></font></B></td>
	      			<%Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Revisor")%></font></td>
	      			<%End If%>
		      		<td width="12%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
						<font size="1" face="Arial">&nbsp;APROVADOR: </font>
					</td>
		      		<td width="33%" align="left" bgcolor="#ffffff" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" height="15">
	      			<%If rsRelatorios("Workflow") = "Aprovador" Then%>
						<B><font color=red face="Arial" size="2">&nbsp;<%=rsRelatorios("Aprovador")%></font></B></td>
	      			<%Else%>
						<font face="Arial" size="1">&nbsp;<%=rsRelatorios("Aprovador")%></font></td>
	      			<%End If%>
		    	</tr>
    			<tr>
		      		<td width="12%" align="left" valign="top" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="60">
						<%
'						strDescricao = trim(rsRelatorios("Descricao"))
'						strDescricao_Inv = "" 
'						intInicio = 1
'						For i = 1 to len(strDescricao)
'							If Mid(strDescricao,i,1) = vbCr Then
'								If Mid(strDescricao,intInicio,i-intInicio-1) <> "" Then
'									strDescricao_Inv = strDescricao_Inv & "<BR>" & Mid(strDescricao,intInicio,i-intInicio) 
'									intInicio = i + 2
'								End If
'							End If
'						Next
'						strDescricao_Inv = strDescricao_Inv & Left("<BR>",Len(strDescricao_Inv)*4) & Mid(strDescricao,intInicio,Len(strDescricao))
						
						%>
						<font face="Arial" size="1">
                        <font size="1" face="Arial">&nbsp;DESCRIÇÃO</font>: </font>
					</td>
      				<td width="78%" align="left" valign="top" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" colspan="3" height="60">
						<!--font face="Arial"><TEXTAREA lang=pt style="BORDER-LEFT-COLOR: #ffffff; BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #ffffff; BORDER-TOP-COLOR: #ffffff; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #ffffff; SCROLLBAR-DARKSHADOW-COLOR: #ffffff; SCROLLBAR-BASE-COLOR: #ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: xx-small; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #ffffff; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR: #ffffff; SCROLLBAR-TRACK-COLOR: #ffffff; FONT-SIZE: xx-small; font-family: Arial; LEFT: 2px; WIDTH: 100%; FONT-STYLE: normal; TOP: 1px; HEIGHT: 100%" name=txtDescricao rows=6 readOnly cols=121>&nbsp;<%=strDescricao%></TEXTAREA></font></td-->
						<%If len(strDescricao) > 500 Then%>
							<font face="Arial" size=1 ><%=Left(strDescricao,500)%>...</font></td>
						<%Else%>
							<font face="Arial" size=1 ><%=strDescricao%>&nbsp;</font></td>
						<%End If%>
	    		</tr>
    			<tr>
		      		<td width="12%" align="left" valign="top" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 0px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" bgcolor="#ffffff" height="150">
						<%
						strComments = trim(rsRelatorios("Comments"))
						strComments_Inv = "" 
						intInicio = 1
						For i = 1 to len(strComments)
							If Mid(strComments,i,1) = vbCr Then
								If Mid(strComments,intInicio,i-intInicio-1) <> "" Then
									strComments_Inv = strComments_Inv & Left("<BR>",Len(strComments_Inv)*4) & Mid(strComments,intInicio,i-intInicio)
									intInicio = i + 1
								End If
							End If
						Next
						strComments_Inv = strComments_Inv & Left("<BR>",Len(strComments_Inv)*4) & Mid(strComments,intInicio,Len(strComments))
						
						%>
						<font face="Arial" size="1">&nbsp;COMENTÁRIOS: </font>
					</td>
      				<td width="78%" align="left"  valign="top" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 0px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid" bgcolor="#ffffff" colspan="3" height="150">
						<!--font face="Arial"><TEXTAREA lang=pt style="BORDER-LEFT-COLOR: #ffffff; BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #ffffff; BORDER-TOP-COLOR: #ffffff; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #ffffff; SCROLLBAR-DARKSHADOW-COLOR: #ffffff; SCROLLBAR-BASE-COLOR: #ffffff; SCROLLBAR-FACE-COLOR: #ffffff; FONT-SIZE: xx-small; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #ffffff; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR: #ffffff; SCROLLBAR-TRACK-COLOR: #ffffff; FONT-SIZE: xx-small; font-family: Arial; WIDTH: 100%; FONT-STYLE: normal; HEIGHT: 100%" tabIndex=3 name=S1 rows=6 readOnly cols=121><%=strComments_Inv%></TEXTAREA></font></td-->
						<%If len(strComments_Inv) > 1300 Then%>
							<font face="Arial" size=1 >...<%=Right(strComments_Inv,1300)%></font></td>
						<%Else%>
							<font face="Arial" size=1 ><%=strComments_Inv%>&nbsp;</font></td>
						<%End If%>
	    		</tr>
  			</table>
  			<FONT face=Arial size=1>#<%=intConta%></FONT>
	  	</center>
	</div>

	<!--hr size="3" color="#000000"-->
	<% 
		rsRelatorios.MoveNext
		
	loop

else

	response.write "<p><b><font color=#666666 size=2 face=Arial, Times New Roman, Times, serif>Não foram encontrados dados para este filtro</font></b></p>"
	
end if

'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)

%>
</form>	
</body>
</html>
