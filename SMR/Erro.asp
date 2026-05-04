<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim msgErro
Dim strData
Dim strVoltar
Dim strIrPara
Dim bolExibirBotao

	msgErro		= Trim(Request("Erro"))
	strVoltar	= Trim(Request("Voltar"))
	strIrPara	= Trim(Request("IrPara"))
	strData		= Now

	If (strVoltar = "no") Then
		bolExibirBotao = false
	Else
		bolExibirBotao = true	
	End If
	
	If (strVoltar = "true") Then
		strVoltar = True
	Else
		strVoltar = False
	End If

	If (Trim(strIrPara) <> "") Then
		strVoltar = True
	End If

	If Trim(msgErro) = "" Then
		msgErro = "Sistema temporariamente indisponível. Por favor, tente novamente mais tarde."
	End If
	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">

<SCRIPT language="javascript">

function voltarPara() 
{
<% If (strVoltar) Then %>
	<% If (Trim(strIrPara) <> "") Then %>
		document.location = "<%=strIrPara%>";
	<% Else %>
		history.go(-1);
	<% End If %>
<% End If %>
}

</SCRIPT>

	<link rel="stylesheet" href="estilos/sinergia.css">

</HEAD>
<BODY>
<FORM name="frmErro" id="frmErro" method="post" onsubmit="voltarPara(); return false;">

<TABLE >	
	<TR>			
		<TD align="left">&nbsp;</TD>
		<TD align="right">&nbsp;</TD>
	</TR>
</TABLE>
<CENTER>
<TABLE border="0" style="border: 1 solid #C0C0C0" cellspacing="0" cellpadding="0">
	<TR bgcolor=#6699cc align="left" width="300px">
		<TH><font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;Atenção,</font></TH>
	</TR>
	<TR>
		
		<TD  align=center style="color: #CC3300">
		<font color=Black size="1" face="Georgia, Times New Roman, Times, serif">
			<%=msgErro%>
		</font>
		</TD>
	</TR>
	<TR>
		<TD  align=right width="300px">
		<%If bolExibirBotao Then%>
			<input type="Image" name="cmdVoltar" value="Voltar" src="img/000024.gif" onClick="Confirmar();" align="absmiddle">
			<!--<INPUT  tabindex="0" style="Display:''" TYPE="SUBMIT" NAME="cmdVoltar" ID="cmdVoltar" VALUE="Voltar">-->
			
		<%End If%>
		</TD>
	</TR>
</TABLE>	
</CENTER>

</FORM>
<SCRIPT language=JavaScript>

function Confirmar()
{
	document.frmErro.action = "Erro.asp";
	document.frmErro.submit();
}

</SCRIPT>

</BODY>
</HTML>
