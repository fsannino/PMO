<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<% 
dim intContaProjetos
dim intContaCelulasRestantes
dim intPosicaoInicial
dim intPosicaoFinal
dim intQuantFuncionario
dim intQuantEquipe
dim intQuantComite
dim intQuantPrioridade


dim strDesabilitado
dim strMarcado
dim strValorChkTodosProjetos
dim strNomeArquivoPagina
dim strFiltro
'dim strUsuario
dim strSelecionado
dim strStringOndeVouProcurar
dim strStringProcurada
dim strDetalhadoSelecionado
dim strConsolidadoSelecionado
dim strMarcadoAND
dim strMarcadoOR
dim strInformacaoPesquisada
dim strReadOnly
dim strPesquisaAvancada

dim rsFuncionarios
dim rsComites
dim rsEquipes
dim rsFiltros
dim rsObterFiltros

strValorChkTodosProjetos = ""
strDesabilitado = Request.Form("hidDesabilitado")
strPesquisaAvancada = Request.Form("hidPesquisaAvancada")

if trim(strDesabilitado) <> "" then
	strMarcado = "Checked"
	strValorChkTodosProjetos = "true"
else
	strMarcado = ""
	strValorChkTodosProjetos = ""
end if

'Abrindo uma conexão com o BD
set conConexao = LIGHT_AbrirConexaoBD()

'Recuperando o código da página na tabela Paginas
intPosicaoInicial = InStrRev(Request.ServerVariables("URL"),"/") + 1
intPosicaoFinal = Len(Request.ServerVariables("URL"))
strNomeArquivoPagina = Mid(Request.ServerVariables("URL"), intPosicaoInicial, intPosicaoFinal)

'strSQL = "EXECUTE SP_OBTER_CODIGO_PAGINA '" & strNomeArquivoPagina & "'"

'set rsPaginas = Server.CreateObject("ADODB.RecordSet")

'rsPaginas.CursorLocation = 3
'rsPaginas.Open strSQL, conConexao, 3, 1, 1

strSQL = "EXECUTE SP_LISTAR_PROJETOS Null,'R' "

set rsProjetos = Server.CreateObject("ADODB.RecordSet")

rsProjetos.CursorLocation = 3
rsProjetos.Open strSQL, conConexao, 3, 1, 1

strSQL = "EXECUTE SP_LISTAR_FUNCIONARIOS"

set rsFuncionarios = Server.CreateObject("ADODB.RecordSet")

rsFuncionarios.CursorLocation = 3
rsFuncionarios.Open strSQL, conConexao, 3, 1, 1

strSQL = "EXECUTE SP_LISTAR_COMITES"

set rsComites = Server.CreateObject("ADODB.RecordSet")

rsComites.CursorLocation = 3
rsComites.Open strSQL, conConexao, 3, 1, 1

strSQL = "EXECUTE SP_LISTAR_EQUIPES"

set rsEquipes = Server.CreateObject("ADODB.RecordSet")

rsEquipes.CursorLocation = 3
rsEquipes.Open strSQL, conConexao, 3, 1, 1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled Document</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language=JavaScript SRC="funcoes/Funcoes.js"></script>
	<script type="text/javascript" language="JavaScript">
	<!--
	function ValidarFormulario()
	{
		//Validando Projetos
		if (ValidarProjetos())
		{
			//Validando Tipos
			if (ValidarTipos())
			{
				//Validando Usuários
				if (ValidarUsuarios())
				{
					//Validando Equipes
					if (ValidarEquipes())
					{
						//Validando Comitês
						if (ValidarComites())
						{
							//Validando Prioridades
							if (ValidarPrioridades())
							{
								//Validando Status
								if (ValidarStatus())
								{
									return true;
								}
								else
								{
									alert('É necessário a escolha de pelo menos um status !');
									document.frmFiltroRelatorio.slcStatus.focus();
									return false;
								}
							}
							else
							{
								alert('É necessário a escolha de pelo menos uma prioridade !');
								document.frmFiltroRelatorio.slcPrioridade.focus();
								return false;
							}
						}
						else
						{
							alert('É necessário a escolha de pelo menos um comitê !');
							document.frmFiltroRelatorio.slcComite.focus();
							return false;
						}
					}
					else
					{
						alert('É necessário a escolha de pelo menos uma equipe !');
						document.frmFiltroRelatorio.slcEquipe.focus();
						return false;
					}
				}
				else
				{
					alert('É necessário a escolha de pelo menos um usuário !');
					document.frmFiltroRelatorio.slcUsuario.focus();
					return false;
				}
			}
			else
			{
				alert('É necessário a escolha de pelo menos um tipo !');
				document.frmFiltroRelatorio.slcTipo.focus();
				return false;
			}
		}
		else
		{
			alert('É necessário a escolha de pelo menos um projeto !');
			document.frmFiltroRelatorio.chkTodosProjetos.focus();
			return false;
		}
	}

	function ValidarProjetos()
	{
		var intCont	= (document.frmFiltroRelatorio.hidContador.value - 1);
		var blnAcheiProjetoMarcado = false;
		var i = 0;
	
		if (document.frmFiltroRelatorio.chkTodosProjetos.checked != true)
		{
			for(i; i <= intCont; i++)
			{
				if (document.frmFiltroRelatorio.chkProjeto[i].checked == true)
				{
					blnAcheiProjetoMarcado = true;
					return true;
				}
			}
			if (blnAcheiProjetoMarcado == false)
			{
				return false;		
			}
		}
		else
		{
			return true;
		}
	}

	function ValidarTipos()
	{
		if (jTrim(document.frmFiltroRelatorio.slcTipo.value) != '')
			return true
		else
			return false;				
	}

	function ValidarUsuarios()
	{
		if (jTrim(document.frmFiltroRelatorio.slcUsuario.value) != '')
			return true
		else
			return false;
	}

	function ValidarEquipes()
	{
		if (jTrim(document.frmFiltroRelatorio.slcEquipe.value) != '')
			return true
		else
			return false;
	}

	function ValidarComites()
	{
		if (jTrim(document.frmFiltroRelatorio.slcComite.value) != '')
			return true
		else
			return false;	
	}

	function ValidarPrioridades()
	{
		if (jTrim(document.frmFiltroRelatorio.slcPrioridade.value) != '')
			return true
		else
			return false;
	}

	function ValidarStatus()
	{
		if (jTrim(document.frmFiltroRelatorio.slcStatus.value) != '')
			return true
		else
			return false;
	}

	function ValidarOperadorLogico1()
	{
		if (jTrim(document.frmFiltroRelatorio.slcOpLogico1.value) != '')
			return true
		else
			return false;
	}
	
	function ValidarPesquisaAvancada()
	{
		if (jTrim(document.frmFiltroRelatorio.slcCampoAPesquisar.value) != '')
		{
			if (jTrim(document.frmFiltroRelatorio.txtInfoASerPesquisada.value) != '')
			{
				return true
			}
			else
			{
				alert('É necessário digitar o texto a ser pesquisado !');
				document.frmFiltroRelatorio.txtInfoASerPesquisada.focus();
				return false;				
			}
		}
		else
		{
			if (jTrim(document.frmFiltroRelatorio.txtInfoASerPesquisada.value) != '')
			{
				alert('É necessário a escolha de pelo menos um campo !');
				document.frmFiltroRelatorio.slcCampoAPesquisar.focus();
				return false;
			}
			else
			{
				return true
			}
		}

	}
	
	function RedirecionaPaginaRelatorio()
	{
		if (document.frmFiltroRelatorio.rdbTipo[0].checked == true)
		{
			document.frmFiltroRelatorio.hidAction.value = "Light_Relatorio_Consolidado"
		}
		else
		{
			document.frmFiltroRelatorio.hidAction.value = "Light_Relatorio_Detalhado";
		}
	}
	
	function btnRedefinir_onclick()
	{
		//Reinicializando valores de alguns objetos
		document.frmFiltroRelatorio.hidDesabilitado.value = '';
		document.frmFiltroRelatorio.chkTodosProjetos.checked = false;
		document.frmFiltroRelatorio.chkTodosProjetos.value = 'false';
		document.frmFiltroRelatorio.action = 'Light_Relatorio_Filtro.asp?pCaminho=Relatórios';

		document.frmFiltroRelatorio.submit();
	}

	function btnExcel_onclick()
	{
		if (ValidarFormulario())
		{
			if (ValidarPesquisaAvancada())
			{
				document.frmFiltroRelatorio.action = document.frmFiltroRelatorio.hidAction.value + '_Excel.asp';
				document.frmFiltroRelatorio.submit();
			}
		}
	}

	function btnHTML_onclick()
	{
		if (ValidarFormulario())
		{
			if (ValidarPesquisaAvancada())
			{
				document.frmFiltroRelatorio.action = document.frmFiltroRelatorio.hidAction.value + '.asp';
				document.frmFiltroRelatorio.submit();
			}
		}
	}

	function chkTodosProjetos_onclick()
	{
		if (document.frmFiltroRelatorio.chkTodosProjetos.checked == true)
		{
			document.frmFiltroRelatorio.hidDesabilitado.value = 'disabled'
		}
		else
		{
			document.frmFiltroRelatorio.hidDesabilitado.value = '';
		}

		document.frmFiltroRelatorio.action = "Light_Relatorio_Filtro.asp?pCaminho=Relatórios";
		document.frmFiltroRelatorio.submit();
	}

	function btnGravar_onclick()
	{
		if (ValidarFormulario())
		{
			if (ValidarPesquisaAvancada())
			{
				if (jTrim(document.frmFiltroRelatorio.txtNomePesquisa.value) != '') 
				{
					document.frmFiltroRelatorio.action = 'Grava_Filtro_Relatorio.asp';
					document.frmFiltroRelatorio.submit();
					return true;
				}
				else
				{
					alert('É necessário se digitar um nome para essa pesquisa !');
					document.frmFiltroRelatorio.chkTodosProjetos.focus();
					return false;
				}
			}
		}
	}
	
	function chkPesquisaAvancada_onclick()
	{
		if (document.frmFiltroRelatorio.chkPesquisaAvancada.checked == true)
		{
			//Reinicializando a propriedade value de alguns objetos
			document.frmFiltroRelatorio.chkPesquisaAvancada.value = 'true';
			document.frmFiltroRelatorio.hidPesquisaAvancada.value = 'checked';
			//Mostra a table de pesquisa avançada
			tblPesqAvancada1.style.display = '';
//			tblPesqAvancada2.style.display = '';
			window.scrollTo(0,200);
		}
		else
		{
			//Reinicializando a propriedade value de alguns objetos
			document.frmFiltroRelatorio.chkPesquisaAvancada.value = '';
//			document.frmFiltroRelatorio.slcOpLogico1.value = '';
			document.frmFiltroRelatorio.slcCampoAPesquisar.value = '';
			document.frmFiltroRelatorio.txtInfoASerPesquisada.value = '';
			document.frmFiltroRelatorio.hidPesquisaAvancada.value = '';
			
			//Esconde a table de pesquisa avançada
			tblPesqAvancada1.style.display = 'none';
//			tblPesqAvancada2.style.display = 'none';
			window.scrollTo(0,0);
		}
	}

	function slcFiltro_onchange() 
	{
		document.frmFiltroRelatorio.hidFiltro.value = document.frmFiltroRelatorio.slcFiltro.value;
		document.frmFiltroRelatorio.action = 'Light_Relatorio_Filtro.asp?pCaminho=Relatórios';
		document.frmFiltroRelatorio.submit();
		return true;
	}


	function chkTodosTipos_onclick()
	{

		if (document.frmFiltroRelatorio.chkTodosTipo.checked == true)
		{

			for(var i = 0; i <= 3;i++)
			{
				document.frmFiltroRelatorio.slcTipo(i).selected = true;
			}
			window.scrollTo(0,0);
		}
		else
		{
			for(var i = 0; i <= 3;i++)
			{
				document.frmFiltroRelatorio.slcTipo(i).selected = false;
			}
			window.scrollTo(0,0);
		}
	}


	function chkTodosUsuarios_onclick()
	{

		if (document.frmFiltroRelatorio.chkTodosUsuario.checked == true)
		{

			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantFuncionario.value;i++)
			{
				document.frmFiltroRelatorio.slcUsuario(i).selected = true;
			}
			window.scrollTo(0,0);
		}
		else
		{
			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantFuncionario.value;i++)
			{
				document.frmFiltroRelatorio.slcUsuario(i).selected = false;
			}
			window.scrollTo(0,0);
		}
	}



	function chkTodosEquipes_onclick()
	{

		if (document.frmFiltroRelatorio.chkTodosEquipe.checked == true)
		{

			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantEquipe.value;i++)
			{
				document.frmFiltroRelatorio.slcEquipe(i).selected = true;
			}
			window.scrollTo(0,0);
		}
		else
		{
			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantEquipe.value;i++)
			{
				document.frmFiltroRelatorio.slcEquipe(i).selected = false;
			}
			window.scrollTo(0,0);
		}
	}


	function chkTodosComites_onclick()
	{

		if (document.frmFiltroRelatorio.chkTodosComite.checked == true)
		{

			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantComite.value;i++)
			{
				document.frmFiltroRelatorio.slcComite(i).selected = true;
			}
			window.scrollTo(0,0);
		}
		else
		{
			for(var i = 0; i <= document.frmFiltroRelatorio.hidQuantComite.value;i++)
			{
				document.frmFiltroRelatorio.slcComite(i).selected = false;
			}
			window.scrollTo(0,0);
		}
	}


	function chkTodosPrioridades_onclick()
	{

		if (document.frmFiltroRelatorio.chkTodosPrioridade.checked == true)
		{

			for(var i = 0; i <= 3;i++)
			{
				document.frmFiltroRelatorio.slcPrioridade(i).selected = true;
			}
			window.scrollTo(0,0);
		}
		else
		{
			for(var i = 0; i <= 3;i++)
			{
				document.frmFiltroRelatorio.slcPrioridade(i).selected = false;
			}
			window.scrollTo(0,0);
		}
	}



-->
	</script>
</head>
<body onload="chkPesquisaAvancada_onclick()">
<form name="frmFiltroRelatorio" id="frmFiltroRelatorio" method="post">
<link rel="stylesheet" href="estilos/sinergia.css">
<table width="100%" border="0" cellpadding="0">
	<tr> 
	    <td width="50%">
			<font face="Georgia, Times New Roman, Times, serif" size="2" color="#666666">
			<strong>Selecione os projetos de interesse:</strong>
			</font>
		</td>
	    <td width="11%">
		</td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="0">
	<tr align="left" valign="top"> 
    	<td colspan="5"> 
      		<table width="100%" border="0" bordercolor="#000000" cellpadding="0">
        		<tr bgcolor="#639ACE"> 
          			<td bgcolor="#639ACE">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Projetos:</strong>
						</font>
					</td>
          			<td bgcolor=Gainsboro>
						<font color=Black size="1" face="Georgia, Times New Roman, Times, serif">
            			<input type="checkbox" name="chkTodosProjetos" id="chkTodosProjetos" onClick="chkTodosProjetos_onclick()" <%=strMarcado%> value="<%=strValorChkTodosProjetos%>">Todos
						</font>
					</td>
          			<td colspan="4" bgcolor=Gainsboro>
						<font face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;</font> 
          			</td>
        		</tr>
				<%
								
				intContaProjetos = 6
				
				do while not rsProjetos.EOF
					
					if intContaProjetos >= 6 then
						'Número de projetos impressos por linha foi ultrapassado
						Response.Write "<tr bgcolor=Gainsboro>"
						intContaProjetos = 0
					end if

		          	Response.Write "	<td bgcolor=Gainsboro width=""16.66%"">"
					Response.Write "		<font color=Black size=""1"" face=""Georgia, Times New Roman, Times, serif"">"
        		    Response.Write "		<input type=""checkbox"" name=""chkProjeto"" id=""chkProjeto"" value=" & rsProjetos("Project_Code") & " " & strDesabilitado & " " & strMarcado & ">" & ucase(trim(rsProjetos("Desc_Projeto")))
					Response.Write "		</font>"
					Response.Write "	</td>"
					
					intContaProjetos = intContaProjetos + 1
					
					if intContaProjetos >= 6 then
						Response.Write "</tr>"
					end if
					
					rsProjetos.MoveNext
					
				loop

				'Somando 1 a qtde. de células que contêm um checkbox de projeto, para que 
				'possamos verificar quantas células que não possuem um checkbox de projeto irão restar.
				intContaProjetos = intContaProjetos + 1
				
				'Colocando apenas a cor de fundo nas células não foram preenchidas com um checkbox de projeto.
				if intContaProjetos <= 6 then
				
					for intContaCelulasRestantes = intContaProjetos to 6
						Response.Write "	<td bgcolor=""Gainsboro"" width=""16.66%"">&nbsp;</tr>"
					next
					
				end if
				%>
        		<tr bgcolor="Gainsboro"> 
          			<td colspan="6">&nbsp;</td>
        		</tr>
        		<tr bgcolor="#639ACE"> 
          			<td bgcolor="#639ACE">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Tipo de Relatório:</strong>
						</font>
					</td>
          			<td bgcolor="Gainsboro">
						<font size="1" face="Arial, Helvetica, sans-serif">
						<strong>
            			<input type="radio" name="rdbTipo" id="rdbTipo" onClick="RedirecionaPaginaRelatorio();" checked>Simplificado
						</strong>
						</font>
					</td>
          			<td bgcolor="Gainsboro">
						<div align="left">
							<font size="1" face="Arial, Helvetica, sans-serif">
							<strong>
              				<input type="radio" name="rdbTipo" id="rdbTipo" onClick="RedirecionaPaginaRelatorio();">Detalhado
							</strong>
							</font>
						</div>
					</td>
          			<td bgcolor="Gainsboro">&nbsp;</td>
          			<td bgcolor="#639ACE"> 
						<div align="left">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
								<strong>Apenas com RDO:</strong>
							</font>
						</div>
					</td>
          			<td bgcolor="Gainsboro">
						<strong>
						<font size="1" face="Arial, Helvetica, sans-serif"> 
            			<input type="checkbox" name="chkRDO" id="chkRDO" value="SIM">
            			</font>
						</strong>
					</td>
        		</tr>
        		<tr bgcolor=White> 
					<td>&nbsp;</td>
          			<td>&nbsp;</td>
			        <td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
        		<tr bgcolor=White> 
          			<td colspan="6">
						<font face="Georgia, Times New Roman, Times, serif" size="2" color="#666666">
							<strong>Defina as informações de interesse para os filtros dos relatórios:</strong>
						</font>
					</td>
<!--          			<td>&nbsp;</td>
          			<td>&nbsp;</td>-->
        		</tr>
      		</table>
			<table width="100%" border="0">
				<tr> 
					<!--<td align="left" valign="middle" bgcolor="#000000" height="20">-->
					<td align="left" valign="middle" bgcolor="#639ACE" height="20"  width=50%>
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Tipo:</strong>
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="#639ACE" height="20">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Usuário:</strong>
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="#639ACE" height="20">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Equipe:</strong>
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="#639ACE" height="20">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Comitê:</strong>
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="#639ACE" height="20">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Prioridade:</strong>
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="#639ACE" height="20">
						<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
							<strong>Status:</strong>
						</font>
					</td>
				</tr>

				<tr> 
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20"  width=50%>
						<font color="Black" size="1" face="Georgia, Times New Roman, Times, serif">
							<input type="checkbox" name="chkTodosTipo" id="chkTodosTipo" onClick="chkTodosTipos_onclick();" value="">Todos
						</font>
					</td>	
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20">
						<font color="Black" size="1" face="Georgia, Times New Roman, Times, serif">
							<input type="checkbox" name="chkTodosUsuario" id="chkTodosUsuario" onClick="chkTodosUsuarios_onclick();" value="">Todos
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20">
						<font color="Black" size="1" face="Georgia, Times New Roman, Times, serif">
							<input type="checkbox" name="chkTodosEquipe" id="chkTodosEquipe" onClick="chkTodosEquipes_onclick();" value="">Todos
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20">
						<font color="Black" size="1" face="Georgia, Times New Roman, Times, serif">
							<input type="checkbox" name="chkTodosComite" id="chkTodosComite" onClick="chkTodosComites_onclick();" value="">Todos
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20">
						<font color="Black" size="1" face="Georgia, Times New Roman, Times, serif">
							<input type="checkbox" name="chkTodosPrioridade" id="chkTodosPrioridade" onClick="chkTodosPrioridades_onclick();" value="">Todos
						</font>
					</td>
					<td align="left" valign="middle" bgcolor="Gainsboro" height="20">
					</td>
				</tr>

				<tr>
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcTipo" id="slcTipo" size="7" multiple style="WIDTH: 110px; font-family:Arial; font-size:10 px">
		    	    		<option value="Issues" <%=strSelecionado%>>ISSUES</option>
		        			<option value="Risks" <%=strSelecionado%>>RISCOS</option>
		        			<option value="Actions" <%=strSelecionado%>>AÇÕES</option>
		        			<option value="Chang_Requests" <%=strSelecionado%>>CHANGE REQUEST</option>
		      			</select>
					</td>				
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcUsuario" id="slcUsuario" size="7" style="WIDTH: 170px; font-family:Arial; font-size:10 px" multiple>
						<%
						
						intQuantFuncionario = 0
						
						do while not rsFuncionarios.EOF
							%>						
							<option value="<%=rsFuncionarios("Cod_Funcionario")%>" <%=strSelecionado%>>
							<%=trim(ucase(rsFuncionarios("Nome_Funcionario")))%>
							</option>
							
							<%rsFuncionarios.MoveNext
							intQuantFuncionario = intQuantFuncionario + 1
						loop
						%>
		      			</select>
					</td>
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcEquipe" id="slcEquipe" size="7" style="WIDTH: 190px; font-family:Arial; font-size:10 px" multiple>
						<%
						
						intQuantEquipe = 0
						
						do while not rsEquipes.EOF%>	

							<option value="<%=rsEquipes("Cod_Eqp")%>" <%=strSelecionado%>>
							<%=trim(ucase(rsEquipes("Desc_Eqp")))%>
							</option>
							
							<%rsEquipes.MoveNext
							intQuantEquipe = intQuantEquipe + 1								
						loop
						%>
		      			</select>
					</td>
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcComite" id="slcComite" size="7" style="WIDTH: 170px; font-family:Arial; font-size:10 px" multiple>
						<%

						intQuantComite = 0
						
						do while not rsComites.EOF%>
						
							<option value="<%=rsComites("Cod_Comite")%>" <%=strSelecionado%>>
							<%=trim(ucase(rsComites("Desc_Comite")))%>
							</option>
							
							<%rsComites.MoveNext

							intQuantComite = intQuantComite + 1								
							
						loop
						%>
		      			</select>						
					</td>
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcPrioridade" id="slcPrioridade" size="7" style="WIDTH: 130px; font-family:Arial; font-size:10 px" multiple>
		        			<option value="Critical" <%=strSelecionado%>>CRÍTICA</option>
		        			<option value="High" <%=strSelecionado%>>ALTA</option>
		        			<option value="Medium" <%=strSelecionado%>>MÉDIA</option>
		        			<option value="Low" <%=strSelecionado%>>BAIXA</option>
		      			</select>						
					</td>
					<td align="center" valign="top" bgcolor="Gainsboro">
						<select name="slcStatus" id="slcStatus" size="7" style="WIDTH:195px; font-family:Arial; font-size:10 px">
							<option value="SA" <%=strSelecionado%>>ABERTO SEMANA ATUAL</option>
							<option value="PS" <%=strSelecionado%>>ABERTO PROXIMA SEMANA</option>
							<option value="TA" <%=strSelecionado%>>TODOS ABERTOS</option>
							<option value="NP" <%=strSelecionado%>>NO PRAZO</option>
							<option value="PD-10" <%=strSelecionado%>>ATRASADO ATÉ 10 DIAS</option>
							<option value="PD+10" <%=strSelecionado%>>ATRASADO MAIS DE 10 DIAS</option>
							<option value="PD" <%=strSelecionado%>>TODOS ATRASADOS</option>
							<option value="Closed" <%=strSelecionado%>>FECHADO</option>
						</select>
					</td>
				</tr>
<!--      			<table width="100%" border="0" bordercolor="#000000">
	        		<tr bgcolor="#639ACE"> 
    	      			<td width="100%">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
								<strong>
									<input type="checkbox" name="chkPesquisaAvancada" id="chkPesquisaAvancada" <%=strPesquisaAvancada%> onClick="chkPesquisaAvancada_onclick()" value="true">Pesquisa Avançada
								</strong>
							</font>
						</td>
	        		</tr>
    	  		</table>-->
				<table name="tblPesqAvancada1" id="tblPesqAvancada1" width="100%" border="0">
					<tr bgcolor="#639ACE"> 
						<td colspan="6" align="left">
							<font color="#FFFFFF" size="1" face="Georgia, Times New Roman, Times, serif">
								<strong>
								PESQUISA POR CONTEÚDO > Selecione o campo para pesquisa e digite uma ou mais palavras desse campo
								</strong>
							</font>
						</td>
					</tr>
					<tr bgcolor="Gainsboro"> 
						<td width="14%" align="center" valign="top">
							<select name="slcCampoAPesquisar" id="slcCampoAPesquisar" style="WIDTH:140px; font-family:Arial; font-size:10 px">
								<option value="">----------------------------</option>
								<option value="NAME" <%=strSelecionadoNAME%>>Nome</option>
								<option value="COMMENTS" <%=strSelecionadoCOMMENTS%>>Comentário</option>
								<option value="DESCRICAO" <%=strSelecionadoDESCRICAO%>>Descrição</option>
							</select>
						</td>
						<td colspan="2" align="center" valign="top">
							<font face="Arial, Helvetica, sans-serif">
		    		        <textarea name="txtInfoASerPesquisada" id="txtInfoASerPesquisada" cols="200" rows="2" style="WIDTH:755px; font-family:Arial; font-size:10 px"><%=strInformacaoPesquisada%></textarea>
		        		    </font>
						</td>
					</tr>
				</table>
			</table>
		</td>
	  </tr>
</table>
<table width="100%" border="0">
	<tr> 
		<td width="33%">&nbsp;</td>
	    <td width="33%">

			<a href="javascript:btnHTML_onclick();"><img src="img/000050.gif" width="73" height="16" border="0" align="absmiddle"></a>
			<a href="javascript:btnExcel_onclick();"><img src="img/000047.gif" width="73" height="16" border="0" align="absmiddle"></a>
			<a href="javascript:btnRedefinir_onclick();"><img src="img/000048.gif" width="73" height="16" border="0" align="absmiddle"></a>

<!--   	    <input type="button" name="btnHTML" id="btnHTML" value="  HTML  " onClick="btnHTML_onclick()">
        	<input type="button" name="btnExcel" id="btnExcel" value="MS-Excel" onClick="btnExcel_onclick()">
        	<input type="button" name="btnRedefinir" id="btnRedefinir" value="Redefinir" onClick="btnRedefinir_onclick()">
-->
		</td>
		<td width="33%">&nbsp;</td>
	  </tr>
</table>
<input type="hidden" name="hidDesabilitado" id="hidDesabilitado" value="">
<input type="hidden" name="hidPesquisaAvancada" id="hidPesquisaAvancada" value="">
<%
if trim(strFiltro) <> "" then
	if trim(strConsolidadoSelecionado) <> "" then%>
		<input type="hidden" name="hidAction" id="hidAction" value="Light_Relatorio_Consolidado">
	<%else%>
		<input type="hidden" name="hidAction" id="hidAction" value="Light_Relatorio_Detalhado">
	<%end if
else%>
	<input type="hidden" name="hidAction" id="hidAction" value="Light_Relatorio_Consolidado">
<%end if%>
<input type="hidden" name="hidContador" id="hidContador" value="<%=rsProjetos.RecordCount%>">
<input type="hidden" name="hidFiltro" id="hidFiltro" value="">
<input type="hidden" name="hidQuantFuncionario" id="hidQuantFuncionario" value="<%=(intQuantFuncionario - 1)%>">
<input type="hidden" name="hidQuantEquipe" id="hidQuantEquipe" value="<%=(intQuantEquipe - 1)%>">
<input type="hidden" name="hidQuantComite" id="hidQuantComite" value="<%=(intQuantComite - 1)%>">

<script language="JavaScript">
/*if (document.frmFiltroRelatorio.chkPesquisaAvancada.checked == true)
{
	tblPesqAvancada1.style.display = '';
}
else
{
	tblPesqAvancada1.style.display = 'none';
}*/
</script>
</form>
</body>
</html>
<%

'Fechando os Recordsets
rsFuncionarios.Close 
rsComites.Close
rsEquipes.Close

'Destruindo os objetos Recordset
set rsFuncionarios = nothing
set rsComites = nothing
set rsEquipes = nothing
set rsObterFiltros = nothing

'Fechando a conexão com o BD
call FecharConexaoBD(conConexao)
%>