<!--#include file="./funcoes/Funcoes.inc"-->
<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0

Dim strProjeto
Dim strUID
Dim strSql

Dim rs

	strProjeto	= Request("strProjeto")
	strUID	= Request("strUID")
	
	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()

	strSql = "SP_LISTAR_ESTIMATIVAS " & strProjeto & ", " & strUID & " "

	set rs = Server.CreateObject("ADODB.RecordSet")

	rs.OPEN STRSQL, conConexao

%>
<html>
<head>

</head>
<body>
<link rel="stylesheet" href="estilos/sinergia.css">
<form name="frmEstimativas" ID="frmEstimativas" action="" method="post">
	<center>

<%IF Not rs.EOF Then%>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<BR>
	<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0">

		<tr height="17" style="height:12.75pt">
		  <td height="17" class="xl27" width=50px style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
		  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Datas</font></b></td>
		  <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
		  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">% Estimado</font></b></td>
		  <td class="xl27" width="50px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
		  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">% Realizado </font></b></td>
		  <td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
		  	<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Situação</font></b></td>
		</tr>


		<tr height="17" style="height:12.75pt">
		  <td height="17" class="xl22" align=center style="border: 1 solid #666666" width=70px>
		  	<font face="Arial" size="1">
		  		16/07/2004
		  	</font>
		  </td>

		  <td class="xl28" style="border: 1 solid #666666" width="50px">
		  	<font face="Arial" size="1">
		  		<%=rs("PREV_160704")%>&nbsp;
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="50px" align="right" >
		  	<font face="Arial" size="1">
				<%If Not isNull(rs("PREV_160704")) Then%>
			  		<%=rs("REAL_160704")%>&nbsp;
				<%Else%>
					&nbsp;
				<%End If%>
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="80px" align=center >
		  	<font face="Arial" size="1">
				<%If Not isNull(rs("PREV_160704")) Then%>
					<%If isNull(rs("REAL_160704")) Then%>
						&nbsp;
					<%Else%>
						<%If cint(rs("REAL_160704")) < cint(rs("PREV_160704")) Then%>
							<img src="icones/Vermelho.gif" alt="Atrasada" name="Atrasada" border="0" width=18>
						<%Else%>
							<img src="icones/Verde.gif" alt="No Prazo" name="Atrasada" border="0" width=18>
						<%End If%>					
					<%End If%>
				<%Else%>
					&nbsp;
				<%End If%>

		  	</font>
		  </td>
		</tr>


		<tr height="17" style="height:12.75pt">
		  <td height="17" class="xl22" align=center style="border: 1 solid #666666" width=70px>
		  	<font face="Arial" size="1">
		  		01/08/2004
		  	</font>
		  </td>

		  <td class="xl28" style="border: 1 solid #666666" width="50px">
		  	<font face="Arial" size="1">
		  		<%=rs("PREV_010804")%>&nbsp;
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="50px" align="right" >
		  	<font face="Arial" size="1">
				<%If Not isNull(rs("PREV_010804")) Then%>
			  		<%=rs("REAL_010804")%>&nbsp;
				<%Else%>
					&nbsp;
				<%End If%>
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="80px" align=center >
		  	<font face="Arial" size="1">

				<%If Not isNull(rs("PREV_010804")) Then%>
					<%If isNull(rs("REAL_010804")) Then%>
						&nbsp;
					<%Else%>
						<%If cint(rs("REAL_010804")) < cint(rs("PREV_010804")) Then%>
							<img src="icones/Vermelho.gif" alt="Atrasada" name="Atrasada" border="0" width=18>
						<%Else%>
							<img src="icones/Verde.gif" alt="No Prazo" name="Atrasada" border="0" width=18>
						<%End If%>					
					<%End If%>
				<%Else%>
					&nbsp;
				<%End If%>

		  	</font>
		  </td>
		</tr>

		<tr height="17" style="height:12.75pt">
		  <td height="17" class="xl22" align=center style="border: 1 solid #666666" width=70px>
		  	<font face="Arial" size="1">
		  		16/08/2004
		  	</font>
		  </td>

		  <td class="xl28" style="border: 1 solid #666666" width="50px">
		  	<font face="Arial" size="1">
		  		<%=rs("PREV_160804")%>&nbsp;
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="50px" align="right" >
		  	<font face="Arial" size="1">
				<%If Not isNull(rs("PREV_160804")) Then%>
			  		<%=rs("REAL_160804")%>&nbsp;
				<%Else%>
					&nbsp;
				<%End If%>
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="80px" align=center >
		  	<font face="Arial" size="1">

				<%If Not isNull(rs("PREV_160804")) Then%>
					<%If isNull(rs("REAL_160804")) Then%>
						&nbsp;
					<%Else%>
						<%If cint(rs("REAL_160804")) < cint(rs("PREV_160804")) Then%>
							<img src="icones/Vermelho.gif" alt="Atrasada" name="Atrasada" border="0" width=18>
						<%Else%>
							<img src="icones/Verde.gif" alt="No Prazo" name="Atrasada" border="0" width=18>
						<%End If%>					
					<%End If%>
				<%Else%>
					&nbsp;
				<%End If%>

		  	</font>
		  </td>
		</tr>

		<tr height="17" style="height:12.75pt">
		  <td height="17" class="xl22" align=center style="border: 1 solid #666666" width=70px>
		  	<font face="Arial" size="1">
		  		01/09/2004
		  	</font>
		  </td>

		  <td class="xl28" style="border: 1 solid #666666" width="50px">
		  	<font face="Arial" size="1">
		  		<%=rs("PREV_010904")%>&nbsp;
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="50px" align="right" >
		  	<font face="Arial" size="1">
				<%If Not isNull(rs("PREV_010904")) Then%>
			  		<%=rs("REAL_010904")%>&nbsp;
				<%Else%>
					&nbsp;
				<%End If%>
		  	</font>
		  </td>
		  <td class="xl28" style="border: 1 solid #666666" width="80px" align=center >
		  	<font face="Arial" size="1">

				<%If Not isNull(rs("PREV_010904")) Then%>
					<%If isNull(rs("REAL_010904")) Then%>
						&nbsp;
					<%Else%>
						<%If cint(rs("REAL_010904")) < cint(rs("PREV_010904")) Then%>
							<img src="icones/Vermelho.gif" alt="Atrasada" name="Atrasada" border="0" width=18>
						<%Else%>
							<img src="icones/Verde.gif" alt="No Prazo" name="Atrasada" border="0" width=18>
						<%End If%>					
					<%End If%>
				<%Else%>
					&nbsp;
				<%End If%>

		  	</font>
		  </td>
		</tr>


	</table>
	<BR>
	<table style="border-style: solid; border-width: 0" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<TD width=73%>&nbsp;</TD>
			<td align=right>
				<a href="javascript:this.close();" ><img src="img/000023.gif" width="73" height="16" border="0" align="absmiddle"></a>
			</td>
		</tr>

	</table>

	<BR>

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>

	</center>

<%Else
	response.write "<p><b><font size=3 face=Verdana color=#000000>Não foram encontrados dados para este filtro</font></b></p>"
End If%>

</form>
</body>
</html>
