// ----------------------------------------------------------------------------------------------------

function ValidaEmail(Email)
{	

//Contador que possibilita o for na String recebida
var intPos;
 
//Quantidade de caracteres da String recebida
var intTam;
 
//Código ASCII do caracter que está sendo processado
var intASCII;
 
//Armazena a posição do @ na string
var intPosArroba;
 
//Armazena a posição do ponto após o @ na String
var intPosPonto;
 
//Constantes contendo os caracteres válidos de um email
var strCharAaZ = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
var strChar0a9 = '0123456789';
var strCharOutrosValidos = '_-.@';
var strAux;
      
    
   //Verifica se a string recebida foi preenchida com pelo menos 3 caracteres  'X@X'
   if (Email.length < 3){
		return(false);
   }
   
       
   // === REGRAS DO ARROBA NO E-MAIL
   
   //Recupera a quantidade de caracteres recebidos na String
   strAux = jTrim(Email);
   intTam = strAux.length;
   
   //Armazena a posição do arroba na String
   intPosArroba = Email.indexOf('@');
 
   //Se não existir @ não é um email válido.
   if (intPosArroba == -1) { 
      return(false);
   }
      
   //Armazena a posição do ponto posterior ao @ na String
   intPosPonto = Email.indexOf('.', intPosArroba + 1);
   
   if (Email.indexOf('..') != -1) { 
      return(false);
   }

 
   //Precisa existir um e somente um @
   if (intPosArroba >= 0) {
      //Existe pelo menos um @ na string
      if (Email.indexOf('@', intPosArroba + 1) != -1) {
         // Existe pelo mais de um @ na String
         return(false);
      }
      else {
         //Só existe um @ na String
         //Verifica se existe pelo menos um caracter antes do @ na string   --- "x@"
         if (intPosArroba == 0) {
            return(false);
         } 
         
         //e pelo menos 4 após o @ na string    --- "@x.xx"
         if ((intPosArroba + 4) >= (intTam)) {
            return(false);
         }
         
         //e deverá existir pelo menos um caracter entre o @ e o ponto
         if ((intPosPonto - intPosArroba) < 2) {
			return(false);
		 }
      }
   }
   else {
      return(false);
   }  		
   
   
   
   // =================================================================
   //  Garantimos até aqui que temos uma String do tipo   --- "x@x.xx"
   // =================================================================
   
   var strCharValidos
   strCharValidos = strCharAaZ + strChar0a9;
   
   //Verifica se o primeiro e o último caracter são alfanumericos
   if ((strCharValidos.indexOf(Email.charAt(0)) == -1) || (strCharValidos.indexOf(Email.charAt(intTam - 1)) == -1))
   {
	  return(false);
   } 
   
   //Percorre todos os caracteres da String recebida com o objetivo de validá-las
    
   var intI=0;
   var strCaracterCorrente;
   var bolAchou;
   
   strCharValidos = strCharAaZ + strChar0a9 + strCharOutrosValidos;
   
   for (intI = 0; intI <= (intTam - 1); intI++)
   {
	  strCaracterCorrente = Email.substring(intI, intI + 1);
			  
	  bolAchou = strCharValidos.indexOf(strCaracterCorrente);
		
	  //se não encontrou no conjunto de caracteres válidos, sai
	  if (bolAchou == -1)
	  {
	     return(false);
	  }
	}
   
   
   return(true);

}

// ----------------------------------------------------------------------------------------------------

function LimparSelecaoSelect(Objeto)
{
	document.all.item(Objeto).selectedIndex = -1;
}

// ----------------------------------------------------------------------------------------------------

function jTrim(strMsg)
{   
	strNm = '';
	tamstr = strMsg.length;
	tot = tamstr;
	ileft = 0;
	iright= 0;
	
	while(strMsg.charAt(ileft)== " ")
	{
		ileft = ileft + 1;
	}

	if (ileft == tot)
	{
		return "";
	}
				
	while(strMsg.charAt(tamstr-1)== " ")
	{
		tamstr = tamstr - 1;
		iright = iright + 1;
	}

	strNm = strMsg.substring(ileft,tot-iright);
	return strNm;	
}

// ----------------------------------------------------------------------------------------------------

function VerificaBrancos(Valor){
var strValor = Valor;
var intI = 0;

	for (intI = 0; intI < strValor.length; intI++)
	{	
		if (strValor.substring(intI,intI + 1) == " ")
		{
			return false;
		}	
	}
	
	return false;
}

// ----------------------------------------------------------------------------------------------------

function ValidaNumerico(Valor, Vazio, Decimais, Sinal, VlrMin, VlrMax)
{
	
	//Verifica se existem espaços entre os números
	if (VerificaBrancos(Valor)){
		return false;
	}
	
	//se o campo não puder estar vazio testa se está
	//se puder e estiver retorna OK 
	if (!Vazio)
	{
		if (jTrim(Valor) == '')
		{
			return false;
		}
	}
	else
	{
		if (jTrim(Valor) == '')
		{
			return true;
		}
	}

var bolAchou;
var bolJaTemVirgula;
var strCaracterCorrente;
var intI = 0;
var intTamanhoValor = Valor.length - 1;
var strNumeros = '0123456789';

	//Se for indicado na chamada da função que
	//que o número pode conter decimais,
	//acrescenta a vírgula aos caracteres permitidos
	if (Decimais > 0)
	{
		strNumeros = strNumeros + ',';
	}
	
	//Se for indicado na chamada da função que
	//que o número pode ter sinal,
	//acrescenta o sinal negativo aos caracteres permitidos
	if (Sinal)
	{	
		strNumeros = '-' + strNumeros;
	}
	
	//verifica se cada caracter é válido
	for (intI = 0; intI <= intTamanhoValor; intI++)
	{
		strCaracterCorrente = Valor.substring(intI, intI + 1);
			  
		bolAchou = strNumeros.indexOf(strCaracterCorrente);

		//se não for, sai
		if (bolAchou == -1)
		{
			return(false);
		}
		else
		{	
			//se for válido e estiver previsto sinal negativo,
			// valida a posição do sinal
			if (strCaracterCorrente == '-') 
			{
				if (!(intI == 0))
				{
					return(false);
				}
			}
			
			//se o caractere corrente vor a vírgula
			if (strCaracterCorrente == ',')
			{
					//testa o número de casas decimais
					// e a posição da vírgula
					if (intI < intTamanhoValor - Decimais)
					{
						return(false);
					}
			}
		}		  
	}

		
	// Verifica se o numérico está no intervalo definido,
	// se o mesmo for definido
	if (VlrMin != 0 || VlrMax != 0)
	{
		if ((TrocaSeparadorDecimal(Valor,"US") < VlrMin) || 
		    (TrocaSeparadorDecimal(Valor,"US") >= VlrMax))
		{
			return(false);
		}
	}

	return true;
}

// ----------------------------------------------------------------------------------------------------

function ValidaTexto(Valor, Vazio, CaracteresInvalidos)
{	

	//se o campo não puder estar vazio testa se está
	//se puder e estiver retorna OK 
	if (!Vazio)
	{
		if (jTrim(Valor) == '')
		{
			return false;
		}
	}
	else
	{
		if (jTrim(Valor) == '')
		{
			return true;
		}
	}
	


var bolAchou;
var strCaracterCorrente;
var intI = 0;
var intTamanhoValor = Valor.length - 1;


	// Se o parâmetro de caracteres inválidos não estiver preenchido,
	// preenche com valores defaults + aspas simples 
	if (CaracteresInvalidos == '')
	{
		CaracteresInvalidos = '!#$%¨&*()_-+=*/<>?{}[]´`^~;:' + "'" ;
	}

	
	//verifica se cada caracter é válido
	for (intI = 0; intI <= intTamanhoValor; intI++)
	{
		strCaracterCorrente = Valor.substring(intI, intI + 1);
			  
		bolAchou = CaracteresInvalidos.indexOf(strCaracterCorrente);
		
		//se encontrou algum caracter inválido, sai
		if (bolAchou != -1)
		{
			return(false);
		}
	}

	return true;
}

// ----------------------------------------------------------------------------------------------------

function DesabilitaVetorCampos(NomeCampo)
{
var intContador=0;
var intNumControles=document.all.item(NomeCampo).length - 1;
				
	for(intContador==0;intContador<=intNumControles;intContador++)
	{
		document.all.item(NomeCampo)(intContador).disabled=true;
	}
}

// ----------------------------------------------------------------------------------------------------

function DesabilitaCampos(ListaDeCampos)
{
var intContador=0;
var arrCampos = ListaDeCampos.split("|");
var intNumControles = arrCampos.length - 1;
	for(intContador==0;intContador<=intNumControles;intContador++)
	{
		document.all.item(arrCampos[intContador]).disabled=true;
	}
}

// ----------------------------------------------------------------------------------------------------

function VerificaSelecionados(NomeVetorCampos)
{
var intContador=0;
var intNumControles=document.all.item(NomeVetorCampos).length - 1;
var intAcumulador=0;
				
	for(intContador==0;intContador<=intNumControles;intContador++)
	{
		if (document.all.item(NomeVetorCampos)(intContador).checked)
		{
			intAcumulador++;
		}
	}
	return(intAcumulador);
}

// ----------------------------------------------------------------------------------------------------

function TrocaSeparadorDecimal(Valor,BR_US)
{

var intPosicao;
var De;
var Por;
var ValorFinal;
var Valor1 = new String(Valor);

	//Valor1 = Valor.String();

	//alert(Valor + ' ' + BR_US);
	if (BR_US == 'BR')
	{
		De = '.';
		Por = ',';
	}
	else
	{
		De = ',';
		Por = '.';
	}

	intPosicao = Valor1.indexOf(De);
	
	// Só trocará se o Caracter ("." ou ",") for encontrado
	if (intPosicao != -1)
	{
		ValorFinal = Valor1.substring(0,intPosicao) + Por;
		ValorFinal += Valor1.substring(intPosicao+1,Valor1.length);
	}
	else	
	{
		ValorFinal = Valor1;
	}
			
		
	return(ValorFinal);
}

// ----------------------------------------------------------------------------------------------------

function LocalizaValorVetorCampos(NomeVetorCampos, Valor, Selecionar)
{
 
var intContador=0;
var intNumControles=document.all.item(NomeVetorCampos).length - 1;
    
	for(intContador==0;intContador<=intNumControles;intContador++)
	{
		if (document.all.item(NomeVetorCampos)(intContador).value == Valor)
		{
			if (document.all.item(NomeVetorCampos).tagName == 'SELECT')
			{
				document.all.item(NomeVetorCampos).value = Valor;
			}
			else
			{
				document.all.item(NomeVetorCampos)(intContador).checked = true;
			}
			
		return(intContador);
		
		}
	}
 
}

// ----------------------------------------------------------------------------------------------------

function LimpaVetorCampos(NomeCampo)
{
var intContador = 0;
var intNumControles = document.all.item(NomeCampo).length - 1;
    
	for(intContador == 0; intContador <= intNumControles; intContador++)
	{
		document.all.item(NomeCampo)(intContador).checked = false;
	}
}

// ----------------------------------------------------------------------------------------------------

function ValidarData(Data) 
{
 // Valida se a data entrada eh valida e se esta no formato dd/mm/aaaa
 // ou se é nula
 
var dtData;
var intI;
var strAux;
var intResultado;	
var strCaracteresValidos = "0123456789"
var strDia = Data.substring(0, 2)
var strBarra1 = Data.substring(2, 3)
var strMes = Data.substring(3, 5)
var strBarra2 = Data.substring(5, 6)
var strAno = Data.substring(6, 10)
 
	// Se a data esta em branco
	if (Data.length == 0) return(true);
	 
	// Se tamanho <> 10 dd/mm/aaaa
	if (Data.length != 10) return(false);
	 
	strDia = Data.substring(0, 2)
	strBarra1 = Data.substring(2, 3)
	strMes = Data.substring(3, 5)
	strBarra2 = Data.substring(5, 6)
	strAno = Data.substring(6, 10)
	 
	dtData = strDia + strMes + strAno
	 
	for (intI = 0; intI <= 7; intI++) {
	 strAux = dtData.substring(intI, ++intI);
	 intResultado = strCaracteresValidos.indexOf(strAux);
	 if (intResultado == -1) return(false);
	}
	 
	if (strMes < 1 || strMes > 12) return(false);

	if (strBarra1 != '/')  return(false);

	if (strDia < 1 || strDia > 31) return(false);

	if (strBarra2 != '/')  return(false);

	if (strAno<1)  return(false);

	if (strAno < 1753) return(false);

	if (strMes == 4 || strMes == 6 || strMes == 9 || strMes == 11) 
	{
		if (strDia == 31)  return(false);
		
	}

	if (strMes == 2)
	{
		strAux = parseInt(strAno/4);
		if (isNaN(strAux)) return(false);
		if (strDia > 29) return(false);
		if (strDia == 29 && ((strAno/4) != parseInt(strAno/4))) return(false);
	}
	 
	return(true);

}

// ----------------------------------------------------------------------------------------------------

//Converte a data passada num dos três formatos do swicth
//e transforma num objeto Date do Java Script
function ConverteDataHora(Data,Formato)
{
var dtData
var strDia = '';
var strMes = '';
var strAno = '';

	switch (Formato)
	{
		case 'BR' :
			strDia = Data.substring(0,2);
			strMes = Data.substring(3,5);
			strAno = Data.substring(6,16);
			break;

		case 'Universal' :
			strDia = Data.substring(6,8);
			strMes = Data.substring(4,6);
			strAno = Data.substring(0,4);
			strAno = Data.substring(9,16);
			break;
			
		case 'US':
			strAno = Data;
			break;
	}

	strAno = strMes + "/" + strDia + "/" + strAno;
	dtData = new Date(strAno);
	return(dtData);

}

// ----------------------------------------------------------------------------------------------------

//Converte a data passada num dos três formatos do swicth
//e transforma num objeto Date do Java Script
function ConverteData(Data,Formato)
{
var dtData
var strDia = '';
var strMes = '';
var strAno = '';

	switch (Formato)
	{
		case 'BR' :
			strDia = Data.substring(0,2);
			strMes = Data.substring(3,5);
			strAno = Data.substring(6,10);
			break;

		case 'Universal' :
			strDia = Data.substring(6,8);
			strMes = Data.substring(4,6);
			strAno = Data.substring(0,4);
			break;
			
		case 'US':
			strAno = Data;
			break;
	}

	strAno = strMes + "/" + strDia + "/" + strAno;
	dtData = new Date(strAno);
	return(dtData);

}


// ----------------------------------------------------------------------------------------------------

//Retorna a data corrente no formaro brasileiro
function DataCorrente()
{
var Data = new Date();

	return(Data.getDate() + '/' + (Data.getMonth()+1) + '/' + Data.getFullYear())

}

// ----------------------------------------------------------------------------------------------------

//Passe duas datas. Se a primeira for maior que a segunda, retorna True.
//Indique também em qual formato você está passando as duas datas.
//Esta função depende da ConverteData
function TestarDifData(DataMenor,DataMaior,Formato)
{
var dtMenor;
var dtMaior;

	if (DataMenor != '')
	{
		dtMenor = ConverteData(DataMenor,Formato);
	}

	if (DataMaior != '')
	{
		dtMaior = ConverteData(DataMaior,Formato);
	}

	if (dtMenor <= dtMaior)
	{
		return (true);
	}
	else
	{
		return (false);
	}

}

// ----------------------------------------------------------------------------------------------------

// Formata uma data no padrão dd/mm/yyyy, com dois dígitos no dia e dois no mês
function FormataDataDDMM(Data)
{

var intPosDia;
var intPosMes;
var intPosAno;
var barra;
var dia='';
var mes='';
var ano='';
var DataFormatada;

	barra = '/';
	
	intPosDia = Data.indexOf(barra);
	
	// busca o dia
	if (intPosDia != -1)
	{
		dia = Data.substring(0,intPosDia);
		if (dia.length < 2)
		{
			dia = '0' + dia;			
		}
	}

	intPosMes = Data.indexOf(barra,intPosDia+1);
	
	// busca o mês
	if (intPosMes != -1)
	{
		mes = Data.substring(intPosDia+1,intPosMes);
		if (mes.length < 2)
		{
			mes = '0' + mes;			
		}	
			
		ano = Data.substring(intPosMes+1,Data.length);
	}

	
	DataFormatada = dia + '/' + mes + '/' + ano;
	return (DataFormatada);
	
}

// ----------------------------------------------------------------------------------------------------

function Redirecionar(Destino){
var Formulario = document.createElement("FORM")

	Formulario.name = 'Redirecionar';
	Formulario.id = 'Redirecionar';
	Formulario.method = 'POST';
	Formulario.action = Destino;
	Formulario.submit();
}

var iSelecionado =0

// ----------------------------------------------------------------------------------------------------

function PopulaComboDerivada(Objeto1, Objeto2, Objeto3,ObjetoSelecionado, Txt)
{
	//Ex.: Objeto1 - Cliente	//Objeto2 -	ClienteProjeto	//Objeto3 - Projeto

    var i;
    var Posicao;
    var objOPTION;
   
    var vlrSelecionador;
    var vlrSelecionado;
    var texto;
    var EncontrouSelecionado =  false;
    
    iSelecionado = 0
    
	//Limpa a Combo a ser Preenchida
    for (i = Objeto3.options.length  ; i >= 0; i-- )
	{
		Objeto3.options.remove(0);    
	}

	objOPTION = document.createElement("OPTION");
	objOPTION.value = ""
	objOPTION.text = Txt // -- Selecione bla bla bla --
	Objeto3.add(objOPTION);
	
	// Varre a Lista com os dados
    for (i = 0; i < Objeto2.length; i++ )
	{	
	
		Posicao =  Objeto2[i].value.indexOf("|")
		
		vlrSelecionador = Objeto2[i].value.substring(0,Posicao)
		vlrSelecionado = Objeto2[i].value.substring(Posicao +1 ,Objeto2[i].value.length)
		texto = Objeto2[i].innerText
		
		//Se o valor 
		if (Objeto1.value == vlrSelecionador)
		{
								
				objOPTION = document.createElement("OPTION");
				objOPTION.value = vlrSelecionado
				objOPTION.text = texto

				Objeto3.add(objOPTION);
				
				if (!EncontrouSelecionado)
				{
					iSelecionado ++;
				}
				
				if (ObjetoSelecionado.value == vlrSelecionado)
				{
					EncontrouSelecionado = true;
				}
				
		}		

	}
}

// ----------------------------------------------------------------------------------------------------
	
function CampoObrigatorio(Objeto){
	if (jTrim(Objeto.value) == '')
	{	
		alert('Campo Obrigatório.');
		Objeto.focus();
		return (true);
	}
	else
	{
		return (false);
	}
}

// ----------------------------------------------------------------------------------------------------

function desabilitar(objeto){
	objeto.readOnly = true;
	objeto.style.backgroundColor = "#DCDCDC";
}

// ----------------------------------------------------------------------------------------------------

function habilitar(objeto){
	objeto.readOnly = false;
	objeto.style.backgroundColor = "white";
}

// ----------------------------------------------------------------------------------------------------

function obterIndicePorValor(combo, valor) {
	for (i = 0; i < combo.options.length; i++) {
		if (combo.options[i].value == valor) {
			return i;
		}
	}
	return null;
}

// ----------------------------------------------------------------------------------------------------

function obterIndicePorTexto(combo, texto) {
	for (i = 0; i < combo.options.length; i++) {
		if (combo.options[i].innerText == texto) {
			return i;
		}
	}
	return null;
}

// ----------------------------------------------------------------------------------------------------

function ValidaHora(hora) {
	// Valida se a hora entrada eh valida e se esta no formato hh:mm
	// ou se é nula
	var hr;
	var fc_i;
	var fc_a1;
	var fc_str;
	var fc_CharsValidos = "0123456789"
	var fc_hora = hora.substring(0, 2)
	var fc_p1 = hora.substring(2, 3)
	var fc_min = hora.substring(3, 5)

	if (hora.length == 0) return(true);
	
	if (hora.length != 5) return(false);
	
	fc_hora = hora.substring(0, 2)
	fc_p1 = hora.substring(2, 3)
	fc_min = hora.substring(3, 5)

	hr = fc_hora + fc_min

	for (fc_i = 0; fc_i <= 3; fc_i++) {
		fc_a1 = hr.substring(fc_i, ++fc_i);
		fc_str = fc_CharsValidos.indexOf(fc_a1);
		if (fc_str == -1) return(false);
	}
	
	if (fc_hora > 23) return(false);
	if (fc_p1 != ':')  return(false);
	if (fc_min > 59) return(false);
		
	return(true);
}

// ----------------------------------------------------------------------------------------------------

/*
* Garante que a página em questão seja carregada em tela cheia
* Ex.: tela de login.
**/
function garantirAberturaEmTelaCheia() {
	if (top.location.href != document.location.href) {
		top.location.href = document.location.href;
	}
}

// ----------------------------------------------------------------------------------------------------

/*
* Se a página em questão não estiver dentro da estrutura de frames da aplicação
* desvia a chamada para a página  principal
**/
function garantirAberturaEmFrame() {
	if (top.location.href == document.location.href) {
		top.location.href = "SAFPW000.asp";
	}
}

// ----------------------------------------------------------------------------------------------------
/*
* Esta função abre a URL recebida em uma nova janela do browser pré-configurada
**/
function abrirRelatorio(sURL) {

var sName;
var sFeatures;
var bReplace;
var novaJanela;

	sName = null;

	sFeatures = "";
	sFeatures = sFeatures + "channelmode=no,"
	sFeatures = sFeatures + "directories=no,"
	sFeatures = sFeatures + "fullscreen=no,"
	sFeatures = sFeatures + "height=520,"
	sFeatures = sFeatures + "left=0,"
	sFeatures = sFeatures + "location=no,"
	sFeatures = sFeatures + "menubar=yes,"
	sFeatures = sFeatures + "resizable=yes,"
	sFeatures = sFeatures + "scrollbars=yes,"
	sFeatures = sFeatures + "status=no,"
	sFeatures = sFeatures + "titlebar=yes,"
	sFeatures = sFeatures + "toolbar=no,"
	sFeatures = sFeatures + "top=0,"
	sFeatures = sFeatures + "width=790"
	
	bReplace = false;
	
	novaJanela = window.open(sURL, sName, sFeatures, bReplace);
	novaJanela.focus();

}

// ----------------------------------------------------------------------------------------------------

function TestarDifDataHora(Obj1, Obj2)
{
//RECEBE DOIS OBJETOS COM DATA E HORA NO FORMAT BRASIL
//EX.: 22/11/2001 12:10

dtData  = ConverteDataHora(Obj1,'BR');
dtData1 = ConverteDataHora(Obj2,'BR');

	if (dtData < dtData1)
	{
		return (true);
	}
	else
	{
		return (false);
	}
}

// ----------------------------------------------------------------------------------------------------

function Inserir(Todos, objLista1, objLista2) {
//Função utilizada para mover os itens de uma lista de origem para uma lista de destino
//Todos: passe true se desejar que todos os itens sejam movidos
//objLista1: Objeto "SELECT" origem
//objLista2: Objeto "SELECT" destino

var arrSelecionados = new Array();		//Arranjo com os índices dos itens selecionados
var intCont			= 0;				//Contador da lista origem
var intCont2		= 0;				//Contador da lista destino
var objOPTION;							//Variável de manipulação do elemento OPTION
var Tamanho			= objLista1.size;	//Pegando o tamanho original da lista
var LimiteDeSelecao = 50				//Limite de itens que o usuário pode escolher ao mesmo tempo
	

	//Aumentando as listas para evitar o erro interno do navegador
	if (objLista1.options.length >= objLista2.options.length) {
		objLista1.size = objLista1.options.length;
		objLista2.size = objLista1.options.length;	
	}
	else {
		objLista1.size = objLista2.options.length;
		objLista2.size = objLista2.options.length;	
	}
	
	//Pegando os índices dos itens selecionados
	//No caso de todos, pega todos os índices
	if (Todos) {
		for (intCont=0; intCont < objLista1.options.length;intCont++){
			arrSelecionados[arrSelecionados.length] = intCont;
			objLista1.options(intCont).selected = false;
		}
	}
	else {
		for (intCont=0; intCont < objLista1.options.length; intCont++){
			if(objLista1.options(intCont).selected) {
				arrSelecionados[arrSelecionados.length] = intCont;
				objLista1.options(intCont).selected = false;
			}
		}
	}

	//Testando se o usuário escolheu mais que o limite
	if (arrSelecionados.length > LimiteDeSelecao) {
		//Redimensionando as listas para o tamanho normal
		objLista1.size = Tamanho;
		objLista2.size = Tamanho;
		
		alert('Selecione no máximo ' + LimiteDeSelecao + ' itens por vez.');
		objLista1.selectedIndex = -1;
		return (false);
	}

	if (arrSelecionados.length == 0) {
		//Redimensionando as listas para o tamanho normal
		objLista1.size = Tamanho;
		objLista2.size = Tamanho;
		return (false);
	}
		
	objLista1.selectedIndex = -1;
	objLista2.selectedIndex = -1;
	
	for (intCont = 0; intCont < arrSelecionados.length; intCont++) {
		
		//Pedindo encarecidamente que o mecanismo JScript remova os objetos inúteis
		CollectGarbage();
		
		//Criando os elementos a serem inseridos na lista destino
		objOPTION = document.createElement("OPTION");
		
		//Copiando as propriedades originais
		objOPTION.value = objLista1.options(arrSelecionados[intCont]).value;
		objOPTION.text  = objLista1.options(arrSelecionados[intCont]).innerText;
			
		//Se já houver item na lista destino, efetua a inserção ordenada
		if (objLista2.options.length != 0) {
			
			//Procurando a posição correta do fim para o início
			//pois é a maior probabilidade
			for (intCont2 = objLista2.options.length - 1; intCont2 >= 0; intCont2--) {
				if(RemoveAcentuacao(objLista2.options(intCont2).text) < 
					RemoveAcentuacao(objLista1.options(arrSelecionados[intCont]).innerText)) {
					
					//Encontrando, move e sai deste laço interno				
					objLista2.add(objOPTION,intCont2 + 1); 
					break;
				} 
			} 
			
			//Se não encontrou no laço anterior, chegará até aqui
			//Isso quer dizer que o item a ser movido é menor do que o
			//primeiro da lista de destino
			//Sendo assim, movo para antes do primeiro
			if (intCont2 < 0) {
				objLista2.add(objOPTION, 0); 
			}
		}
		else {
			objLista2.add(objOPTION);
		}
		
		//Liberando o objeto
		objOPTION = null;
	}

	//Apagando os itens movidos da lista origem		
	if (Todos) {
		objLista1.length = 0;
	}
	else {
		for (intCont = arrSelecionados.length - 1; intCont >= 0; intCont--) {
			objLista1.remove(arrSelecionados[intCont]);
		}
	}
		
	//Redimensionando as listas para o tamanho normal
	objLista1.size = Tamanho;
	objLista2.size = Tamanho;
	
	return (true);
}

// ----------------------------------------------------------------------------------------------------

function Remover(Todos, objLista1,objLista2) {
	//Foi feita essa passagem para se manter a compatibilidade das páginas já implementadas
	return Inserir(Todos, objLista2, objLista1);
}

// ----------------------------------------------------------------------------------------------------

function RemoveAcentuacao(TextoAcentuado) {
//Utilizado para remover a acentuação para testar a ordem alfabética
var intCont = 0;
var strTextoNaoAcentuado = "";
var chrCaractereNaoAcentuado = "";

	for (intCont = 0; intCont < TextoAcentuado.length; intCont++){

		chrCaractereNaoAcentuado = TextoAcentuado.substring(intCont,intCont + 1); //Obtem letra a letra
		chrCaractereNaoAcentuado = TextoAcentuado.charCodeAt(intCont); //Obtem codigo ASCII
		
		//Verifica ocorrencias de A
		if ((chrCaractereNaoAcentuado == 225) 
			|| (chrCaractereNaoAcentuado == 193) 
			|| (chrCaractereNaoAcentuado == 227)
			|| (chrCaractereNaoAcentuado == 195)
			|| (chrCaractereNaoAcentuado == 227)
			|| (chrCaractereNaoAcentuado == 228)
			|| (chrCaractereNaoAcentuado == 194)
			|| (chrCaractereNaoAcentuado == 196))
			
				chrCaractereNaoAcentuado = 65;
				
		else //Verifica ocorrencias de E
			if ((chrCaractereNaoAcentuado == 233)
				|| (chrCaractereNaoAcentuado == 201)
				|| (chrCaractereNaoAcentuado == 203)
				|| (chrCaractereNaoAcentuado == 235))
					
					chrCaractereNaoAcentuado = 69;
					
				else //Verifica ocorrencias de I
					if ((chrCaractereNaoAcentuado == 237)
						|| (chrCaractereNaoAcentuado == 205)
						|| (chrCaractereNaoAcentuado == 207)
						|| (chrCaractereNaoAcentuado == 239))
							
							chrCaractereNaoAcentuado=74;
							
					else //Verifica ocorrencias de O
						if ((chrCaractereNaoAcentuado == 243)
							|| (chrCaractereNaoAcentuado == 211)
							|| (chrCaractereNaoAcentuado == 245)
							|| (chrCaractereNaoAcentuado == 213)
							|| (chrCaractereNaoAcentuado == 246)
							|| (chrCaractereNaoAcentuado == 214))
								chrCaractereNaoAcentuado = 73;
						else //Verifica ocorrencias de U
							if ((chrCaractereNaoAcentuado == 250)
								|| (chrCaractereNaoAcentuado == 218)
								|| (chrCaractereNaoAcentuado == 252)
								|| (chrCaractereNaoAcentuado == 220))
									chrCaractereNaoAcentuado = 85;
							else //Verifica ocorrencias de Ç
								if ((chrCaractereNaoAcentuado == 231)
									||(chrCaractereNaoAcentuado == 199)
									||(chrCaractereNaoAcentuado == 231))
										chrCaractereNaoAcentuado = 67;
										
		chrCaractereNaoAcentuado = String.fromCharCode(chrCaractereNaoAcentuado);//Reconverte o caractere ASCII para letra
				
		strTextoNaoAcentuado = strTextoNaoAcentuado + chrCaractereNaoAcentuado; //Monta a nova String
	
	}
	
	strTextoNaoAcentuado = strTextoNaoAcentuado.toUpperCase();//Passa para maiusculas

	return (strTextoNaoAcentuado);
}