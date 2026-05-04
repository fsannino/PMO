<%

'******************************************
'Biblioteca de funções usadas pelas páginas
'******************************************

'*************************************
'Esta função abre uma conexão com o BD
'*************************************
function LIGHT_AbrirConexaoBD

	dim strConexao
	dim conConexaoAux

	'Path do arquivo UDL
	strConexao = "File Name=C:\Fontes\SMR_Light\Light.udl;"

	'Abre a conexão
	set conConexaoAux = Server.CreateObject("ADODB.Connection")
	conConexaoAux.Open strConexao

	set LIGHT_AbrirConexaoBD = conConexaoAux

end function

'**************************************
'Esta função fecha uma conexão com o BD
'**************************************
function FecharConexaoBD(pConexao)

	pConexao.close
	set pConexao = nothing

end function


'**********************************************
'Esta função extrai uma substring de uma string
'**********************************************
function ExtraiSubstring(pStringOndeVouProcurar,pStringProcurada)

	dim strSubString
	dim intPosicao

	if (trim(pStringOndeVouProcurar) <> "") and (trim(pStringProcurada) <> "") then

		intPosicao = instr(1,pStringOndeVouProcurar,pStringProcurada,1)
		
		if intPosicao > 0 then
			
			strSubString = mid(pStringOndeVouProcurar,(intPosicao+len(pStringProcurada)+1),len(pStringOndeVouProcurar))
			
			ExtraiSubstring = strSubString

		end if

	end if

end function


'****************************************************
'Esta função monta a descrição de um determinado tipo
'****************************************************
function Descricao_Tipo (pCodigoTipo)

	select case ucase(trim(pCodigoTipo))
		case "ISS"
			Descricao_Tipo = "Issues"
		case "ACT"
			Descricao_Tipo = "Ações"
		case "RSK"
			Descricao_Tipo = "Riscos"
		case "CHR"
			Descricao_Tipo = "Change Request"
	end select
	
end function


'***************************************************************
'Esta função retorna a quantidade de linhas necessárias para
'que um TextArea abra automaticamente, sem ter que usar o scroll
'***************************************************************
function RetornaQtdeLinhasTextArea(pString, pQtdeColunas)

	dim I
	
	for I = 1 to len(pString)
		if (mid(pString,I,1) = chr(10)) or (mid(pString,I,1) = chr(13)) then
			intContador = intContador + 1
		end if
	next
	
	intContador = (intContador + (cInt((len(pString)/pQtdeColunas)))) + 2
	
	RetornaQtdeLinhasTextArea = intContador
	
end function


'***************************************************************
'Esta função cria um vetor de acordo com as informações passadas
'pValores   -> Informações a serem tratadas
'pSeparador -> Caracter que vai separar as informações
'pTipoDado  -> S = String / N = Numérico
'***************************************************************
function TrataStringSP(pValores, pSeparador, pTipoDado)

	dim vetVetor
	dim intI

	'Criando um vetor com base nas informações passadas
	vetVetor = split(pValores, pSeparador)

	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	for intI = 0 to ubound(vetVetor)
		if ucase(trim(pTipoDado)) = "S" then
			TrataStringSP = TrataStringSP & """" & trim(vetVetor(intI)) & ""","
		else
			TrataStringSP = TrataStringSP & trim(vetVetor(intI)) & ","
		end if
	next

	'Retornando uma string tratada que vai ser passada para uma stored procedure
	TrataStringSP = mid(TrataStringSP,1,(len(TrataStringSP)-1))

end function


'***************************************************************
'Esta função cria um vetor de acordo com as informações passadas
'pValores   -> Informações a serem tratadas
'pSeparador -> Caracter que vai separar as informações
'***************************************************************
function TrataStringFiltro(pValores, pSeparador)

	dim vetVetor
	dim intI

	'Criando um vetor com base nas informações passadas
	vetVetor = split(pValores, pSeparador)

	'Formatando a informação para que ela possa ser enviada a Stored Procedure
	for intI = 0 to ubound(vetVetor)
'		if ucase(trim(pTipoDado)) = "S" then
			TrataStringFiltro = TrataStringFiltro & trim(vetVetor(intI)) & "|"
'		else
'			TrataStringFiltro = TrataStringFiltro & trim(vetVetor(intI)) & "|"
'		end if
	next

	'Retornando uma string tratada que vai ser passada para uma stored procedure
	TrataStringFiltro = mid(TrataStringFiltro,1,(len(TrataStringFiltro)-1))

end function


'***************************************************************
'Esta função verifica se o usuario tem autorizacao para acessar a pagina
'strCodigo   -> Nome da pagina
'blnRedirecionar -> Redireciona para pagina de erros
'***************************************************************

Function VerificarAutorizacaoUsuario(strCodigo, blnRedirecionar)

Dim strAut
Dim strBusca
Dim msgErro

	strBusca = ";" & strCodigo & ";"

	strAut = Session("Autorizacao")

	If blnRedirecionar Then
		If (InStr(strAut, strBusca) = 0) Then
			msgErro = "Acesso negado. Certifique-se de suas permissões no sistema"
			Response.Redirect "../Erro/Erro.asp?Erro=" & msgErro
		End If
	Else
		VerificarAutorizacaoUsuario = (InStr(strAut, strBusca) > 0)
	End If

End Function


'***************************************************************
'Esta função transforma data no formato dd/mm/yyyy para dd mon yyyy 
'strData   -> Data no formato dd/mm/yyyy
'***************************************************************

Function FormatarDataMon(strData)

Dim strDia
Dim strMes
Dim strAno
Dim strDescMes

	strDia = Mid(strData,1,2)
	strMes = Mid(strData,4,2)
	strAno = Mid(strData,7,4)

	Select Case strMes
		Case "01"
			strDescMes = "Jan"
		Case "02"
			strDescMes = "Fev"
		Case "03"
			strDescMes = "Mar"
		Case "04"
			strDescMes = "Abr"
		Case "05"
			strDescMes = "Mai"
		Case "06"
			strDescMes = "Jun"
		Case "07"
			strDescMes = "Jul"
		Case "08"
			strDescMes = "Ago"
		Case "09"
			strDescMes = "Set"
		Case "10"
			strDescMes = "Out"
		Case "11"
			strDescMes = "Nov"
		Case "12"
			strDescMes = "Dez"		
	End Select		

	FormatarDataMon = strDia & " " & strDescMes & " " & strAno
	
End Function

Function ReplicaPlics(vsArgumento)

   ' Descrição: Replica os Plics para argumentos tipo Texto

   Dim mnIncremento
   Dim msArgumentoFinal
   
   If IsNull(vsArgumento) Then
      ReplicaPlics = ""
      Exit Function
   End If
   
   For mnIncremento = 1 To Len(vsArgumento)
       If Mid(vsArgumento, mnIncremento, 1) = "'" Then
          msArgumentoFinal = msArgumentoFinal & "''"
       Else
          msArgumentoFinal = msArgumentoFinal & Mid(vsArgumento, mnIncremento, 1)
       End If
   Next
   
   ReplicaPlics = msArgumentoFinal
    
End Function

'***************************************************************

Function VerificaLogin(strLogin, strUsuario)
Dim vetDados
Dim I

    If InStr(strLogin, ";") <> 0 Then
        vetDados = Split(strLogin, ";")
        If UBound(vetDados) > 0 Then
            For I = 0 To UBound(vetDados)
                If ucase(trim(vetDados(I))) = ucase(trim(strUsuario)) Then
                    If I = 0 Then
                        VerificaLogin = "P"
                        Exit Function
                    Else
                        VerificaLogin = "S"
                        Exit Function
                    End If
                End If
            Next
        Else
            VerificaLogin = "T"
            Exit Function
        End If
    Else
        VerificaLogin = "T"
        Exit Function
    End If
    
End Function


%>