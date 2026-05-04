<!--#include file="./MenuPrincipal.asp"-->
<!--#include file="./head.asp"-->

<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = 0
Server.ScriptTimeout = 10000

Dim rs
Dim cmdResultado
Dim strDiretoria
Dim strAtividade
Dim strData
Dim vetCS(6,14)
Dim vetAB(12,14)
Dim vetEP(9,14)
Dim vetFI(5,14)
Dim vetGE(4,14)

Dim vetIN(3,14)
Dim vetSE(8,14)
Dim vetPR(9,14)
Dim I
Dim Cont


	'Abrindo uma conexão com o BD
	set conConexao = SMR_AbrirConexaoBD()


Function RetornaValorPainelSinoptico(strTipo, strUID)
Dim cmdResultado
Dim rs
Dim strSql


	If trim(strUID) <> "" Then

		strSql = ""

		Select Case strTipo
			Case "TP1"
				strSql = "SP_LISTAR_PAINEL_SINOPTICO_TIPO1"
			Case "TP2"
				strSql = "SP_LISTAR_PAINEL_SINOPTICO_TIPO2"
		End Select

		If 	strSql <> "" Then 

			Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 600
			    .CommandText = strSql

			    .Parameters.Refresh

				.Parameters(1).Value = strUID

			End With

			set rs = Server.CreateObject("ADODB.RecordSet")

			set rs = cmdResultado.Execute()

			If not rs.EOF Then
				If not isnull(rs("VALOR")) Then
					RetornaValorPainelSinoptico = rs("VALOR")
				Else
					RetornaValorPainelSinoptico = "0"
				End If
			Else
				RetornaValorPainelSinoptico = "0"
			End If
		Else
			RetornaValorPainelSinoptico = "0"
		End If
	Else
		RetornaValorPainelSinoptico = "0"
	End If
	
End Function


Function RetornaDataPainelSinoptico(strUID)
Dim cmdResultado
Dim rs
Dim strSql

	If trim(strUID) <> "" Then

		strSql = "SP_LISTAR_PAINEL_SINOPTICO_DATAS"

		Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
		With cmdResultado
    
		    .ActiveConnection = conConexao
		    .CommandType = 4
			.CommandTimeout = 600
		    .CommandText = strSql
			    
		    .Parameters.Refresh

			.Parameters(1).Value = strUID

		End With

		set rs = Server.CreateObject("ADODB.RecordSet")

		set rs = cmdResultado.Execute()

		If not rs.EOF Then
			RetornaDataPainelSinoptico = rs("TASK_START_DATE")
		Else
			RetornaDataPainelSinoptico = "01/01/1900"
		End If
	Else
		RetornaDataPainelSinoptico = "01/01/1900"
	End If
	
End Function


Function RetornaPrevRealPainelSinoptico(strTipo, strUID)
Dim cmdResultado
Dim rs
Dim strSql


	If trim(strUID) <> "" Then

		strSql = ""

		Select Case strTipo
			Case "TP1"
				strSql = "SP_LISTAR_PAINEL_SINOPTICO_PREVISTO"
			Case "TP2"
				strSql = "SP_LISTAR_PAINEL_SINOPTICO_REALIZADO"
		End Select

		If 	strSql <> "" Then 

			Set cmdResultado = Server.CreateObject("ADODB.Command")
			    
			With cmdResultado
    
			    .ActiveConnection = conConexao
			    .CommandType = 4
				.CommandTimeout = 600
			    .CommandText = strSql

			    .Parameters.Refresh

				.Parameters(1).Value = strUID

			End With

			set rs = Server.CreateObject("ADODB.RecordSet")

			set rs = cmdResultado.Execute()

			If not rs.EOF Then
				If not isnull(rs("VALOR")) Then
					RetornaPrevRealPainelSinoptico = rs("VALOR")
				Else
					RetornaPrevRealPainelSinoptico = " "
				End If
			Else
				RetornaPrevRealPainelSinoptico = " "
			End If
		Else
			RetornaPrevRealPainelSinoptico = " "
		End If
	Else
		RetornaPrevRealPainelSinoptico = " "
	End If
	
End Function



	strDiretoria	= trim(Request("slcDiretoria"))
	strAtividade	= trim(Request("slcAtividade"))


	'Consolidado - Nome
	vetCS(0,1) = "ABAST"
	vetCS(1,1) = "E&P"
	vetCS(2,1) = "FINANCEIRA"
	vetCS(3,1) = "GAS/ENERGIA"
	vetCS(4,1) = "INTERNACIONAL"
	vetCS(5,1) = "SERVIÇOS"
	vetCS(6,1) = "PRESIDENCIA"
	

	'Consolidado - Teste de Campo - Preparação
	vetCS(0,2) = "2965"
	vetCS(1,2) = "9"
	vetCS(2,2) = ""
	vetCS(3,2) = ""
	vetCS(4,2) = "6610"
	vetCS(5,2) = ""
	vetCS(6,2) = ""

	'Consolidado - Teste de Campo - Execução
	vetCS(0,3) = "2984"
	vetCS(1,3) = "5689"
	vetCS(2,3) = ""
	vetCS(3,3) = ""
	vetCS(4,3) = "6630"
	vetCS(5,3) = ""
	vetCS(6,3) = ""

	'Consolidado - Saneamento
	vetCS(0,4) = "7054"
	vetCS(1,4) = "7051"
	vetCS(2,4) = "7057"
	vetCS(3,4) = "7068"
	vetCS(4,4) = "7060"
	vetCS(5,4) = "7073"
	vetCS(6,4) = ""

	'Consolidado - Treinamento
	vetCS(0,5) = "6269"
	vetCS(1,5) = "34"
	vetCS(2,5) = "7079"
	vetCS(3,5) = "7084"
	vetCS(4,5) = "7078"
	vetCS(5,5) = "5312"
	vetCS(6,5) = "4427"

	'Consolidado - Comunicação - Teste de Campo
	vetCS(0,6) = "5727"
	vetCS(1,6) = "5790"
	vetCS(2,6) = "6442"
	vetCS(3,6) = "6919"
	vetCS(4,6) = "6650"
	vetCS(5,6) = "5773"
	vetCS(6,6) = "5745"

	'Consolidado - Comunicação - Entrada em Produção
	vetCS(0,7) = "5917"
	vetCS(1,7) = "6357"
	vetCS(2,7) = "6437"
	vetCS(3,7) = "6973"
	vetCS(4,7) = "6865"
	vetCS(5,7) = "5889"
	vetCS(6,7) = "5845"

	'Consolidado - Mobilização Prontidão
	vetCS(0,8) = "5804"
	vetCS(1,8) = "6356"
	vetCS(2,8) = "6436"
	vetCS(3,8) = "6968"
	vetCS(4,8) = "6845"
	vetCS(5,8) = "5860"
	vetCS(6,8) = "5827"

	'Consolidado - Impacto
	vetCS(0,9) = "7055"
	vetCS(1,9) = "7053"
	vetCS(2,9) = ""
	vetCS(3,9) = "7188"
	vetCS(4,9) = ""
	vetCS(5,9) = "7072"
	vetCS(6,9) = ""

	'Consolidado - Suporte - Cap. ALI's Mapeamento Perfil
	vetCS(0,10) = "3090"
	vetCS(1,10) = "18"
	vetCS(2,10) = "6450"
	vetCS(3,10) = "6951"
	vetCS(4,10) = "6773"
	vetCS(5,10) = "5152"
	vetCS(6,10) = "3509"

	'Consolidado - Suporte - Cap. ALI's Suporte a Usuario
	vetCS(0,11) = "6298"
	vetCS(1,11) = "6286"
	vetCS(2,11) = "6448"
	vetCS(3,11) = "6945"
	vetCS(4,11) = "6751"
	vetCS(5,11) = "6344"
	vetCS(6,11) = "6319"

	'Consolidado - Suporte - Sala de Controle
	vetCS(0,12) = "7032"
	vetCS(1,12) = "7030"
	vetCS(2,12) = "7034"
	vetCS(3,12) = "7038"
	vetCS(4,12) = "7036"
	vetCS(5,12) = "7040"
	vetCS(6,12) = "7042"

	'Consolidado - INFRA - Geral
	vetCS(0,13) = "7056"
	vetCS(1,13) = "7052"
	vetCS(2,13) = "7058"
	vetCS(3,13) = "7070"
	vetCS(4,13) = "7067"
	vetCS(5,13) = "7071"
	vetCS(6,13) = "7076"

	'Consolidado - INFRA - NF
	vetCS(0,14) = "7269"
	vetCS(1,14) = "7258"
	vetCS(2,14) = "7283"
	vetCS(3,14) = "7295"
	vetCS(4,14) = "7290"
	vetCS(5,14) = "7315"
	vetCS(6,14) = "7298"


	'ABAST - Nome
	vetAB(0,1) = "AB"
	vetAB(1,1) = "FAFEN"
	vetAB(2,1) = "LUBNOR"
	vetAB(3,1) = "RECAP"
	vetAB(4,1) = "REDUC"
	vetAB(5,1) = "REGAP"
	vetAB(6,1) = "REMAN"
	vetAB(7,1) = "REPAR"
	vetAB(8,1) = "REPLAN"
	vetAB(9,1) = "REVAP"
	vetAB(10,1) = "RLAM"
	vetAB(11,1) = "RPBC"
	vetAB(12,1) = "SIX"

	'ABAST - Teste de Campo - Preparação
	vetAB(0,2) = ""
	vetAB(1,2) = "2970"
	vetAB(2,2) = ""
	vetAB(3,2) = ""
	vetAB(4,2) = ""
	vetAB(5,2) = ""
	vetAB(6,2) = ""
	vetAB(7,2) = ""
	vetAB(8,2) = "2977"
	vetAB(9,2) = ""
	vetAB(10,2) = "2979"
	vetAB(11,2) = ""
	vetAB(12,2) = ""

	'ABAST - Teste de Campo - Execução
	vetAB(0,3) = ""
	vetAB(1,3) = "2990"
	vetAB(2,3) = ""
	vetAB(3,3) = ""
	vetAB(4,3) = ""
	vetAB(5,3) = ""
	vetAB(6,3) = ""
	vetAB(7,3) = ""
	vetAB(8,3) = "2997"
	vetAB(9,3) = ""
	vetAB(10,3) = "2999"
	vetAB(11,3) = ""
	vetAB(12,3) = ""

	'ABAST - Saneamento
	vetAB(0,4) = "7117"
	vetAB(1,4) = "7118"
	vetAB(2,4) = "7119"
	vetAB(3,4) = "7120"
	vetAB(4,4) = "7121"
	vetAB(5,4) = "7122"
	vetAB(6,4) = "7123"
	vetAB(7,4) = "7124"
	vetAB(8,4) = "7125"
	vetAB(9,4) = "7126"
	vetAB(10,4) = "7127"
	vetAB(11,4) = "7128"
	vetAB(12,4) = "7129"

	'ABAST - Treinamento
	vetAB(0,5) = "6272"
	vetAB(1,5) = "6274"
	vetAB(2,5) = "6275"
	vetAB(3,5) = "6276"
	vetAB(4,5) = "6277"
	vetAB(5,5) = "6278"
	vetAB(6,5) = "6279"
	vetAB(7,5) = "6280"
	vetAB(8,5) = "6281"
	vetAB(9,5) = "6282"
	vetAB(10,5) = "6283"
	vetAB(11,5) = "6284"
	vetAB(12,5) = "6285"

	'ABAST - Comunicação - Teste de Campo
	vetAB(0,6) = "5739"
	vetAB(1,6) = "5741"
	vetAB(2,6) = "5742"
	vetAB(3,6) = "5743"
	vetAB(4,6) = "5744"
	vetAB(5,6) = "5733"
	vetAB(6,6) = "5734"
	vetAB(7,6) = "5735"
	vetAB(8,6) = "5736"
	vetAB(9,6) = "5731"
	vetAB(10,6) = "5732"
	vetAB(11,6) = "5730"
	vetAB(12,6) = "5729"

	'ABAST - Comunicação - Entrada em Produção
	vetAB(0,7) = "5910"
	vetAB(1,7) = "5912"
	vetAB(2,7) = "5913"
	vetAB(3,7) = "5914"
	vetAB(4,7) = "5915"
	vetAB(5,7) = "5916"
	vetAB(6,7) = "5907"
	vetAB(7,7) = "5908"
	vetAB(8,7) = "5909"
	vetAB(9,7) = "5904"
	vetAB(10,7) = "5905"
	vetAB(11,7) = "5903"
	vetAB(12,7) = "5902"

	'ABAST - Mobilização Prontidão
	vetAB(0,8) = "5821"
	vetAB(1,8) = "5823"
	vetAB(2,8) = "5805"
	vetAB(3,8) = "5806"
	vetAB(4,8) = "5807"
	vetAB(5,8) = "5808"
	vetAB(6,8) = "5809"
	vetAB(7,8) = "5810"
	vetAB(8,8) = "5811"
	vetAB(9,8) = "5812"
	vetAB(10,8) = "5813"
	vetAB(11,8) = "5814"
	vetAB(12,8) = "5815"


	'ABAST - Impacto
	vetAB(0,9) = "7130"
	vetAB(1,9) = "7131"
	vetAB(2,9) = "7132"
	vetAB(3,9) = "7133"
	vetAB(4,9) = "7134"
	vetAB(5,9) = "7135"
	vetAB(6,9) = "7136"
	vetAB(7,9) = "7137"
	vetAB(8,9) = "7138"
	vetAB(9,9) = "7139"
	vetAB(10,9) = "7140"
	vetAB(11,9) = "7141"
	vetAB(12,9) = "7142"

	'ABAST - Suporte - Cp. ALI's Mapeamento Perfil
	vetAB(0,10) = "3075"
	vetAB(1,10) = "3077"
	vetAB(2,10) = "3078"
	vetAB(3,10) = "3079"
	vetAB(4,10) = "3080"
	vetAB(5,10) = "3081"
	vetAB(6,10) = "3082"
	vetAB(7,10) = "3083"
	vetAB(8,10) = "3084"
	vetAB(9,10) = "3085"
	vetAB(10,10) = "3086"
	vetAB(11,10) = "3087"
	vetAB(12,10) = "3088"


	'ABAST - Suporte - Cp. ALI's Suporte a Usuario
	vetAB(0,11) = "6301"
	vetAB(1,11) = "6303"
	vetAB(2,11) = "6304"
	vetAB(3,11) = "6305"
	vetAB(4,11) = "6306"
	vetAB(5,11) = "6307"
	vetAB(6,11) = "6308"
	vetAB(7,11) = "6309"
	vetAB(8,11) = "6310"
	vetAB(9,11) = "6311"
	vetAB(10,11) = "6312"
	vetAB(11,11) = "6313"
	vetAB(12,11) = "6314"

	'ABAST - Suporte - Sala de Controle
	vetAB(0,12) = "7340"
	vetAB(1,12) = "7341"
	vetAB(2,12) = "7342"
	vetAB(3,12) = "7343"
	vetAB(4,12) = "7344"
	vetAB(5,12) = "7345"
	vetAB(6,12) = "7346"
	vetAB(7,12) = "7347"
	vetAB(8,12) = "7348"
	vetAB(9,12) = "7349"
	vetAB(10,12) = "7350"
	vetAB(11,12) = "7351"
	vetAB(12,12) = "7352"


	'ABAST - INFRA - GERAL
	vetAB(0,13) = "7143"
	vetAB(1,13) = "7144"
	vetAB(2,13) = "7145"
	vetAB(3,13) = "7146"
	vetAB(4,13) = "7147"
	vetAB(5,13) = "7148"
	vetAB(6,13) = "7149"
	vetAB(7,13) = "7150"
	vetAB(8,13) = "7151"
	vetAB(9,13) = "7152"
	vetAB(10,13) = "7153"
	vetAB(11,13) = "7154"
	vetAB(12,13) = "7155"

	'ABAST - INFRA - NF
	vetAB(0,14) = "7270"
	vetAB(1,14) = "7271"
	vetAB(2,14) = "7272"
	vetAB(3,14) = "7273"
	vetAB(4,14) = "7274"
	vetAB(5,14) = "7275"
	vetAB(6,14) = "7276"
	vetAB(7,14) = "7277"
	vetAB(8,14) = "7278"
	vetAB(9,14) = "7279"
	vetAB(10,14) = "7280"
	vetAB(11,14) = "7281"
	vetAB(12,14) = "7282"



	'E&P - Nomes
	vetEP(0,1) = "E&P-CORP"
	vetEP(1,1) = "E&P-SERV"
	vetEP(2,1) = "UN-BA"
	vetEP(3,1) = "UN-BC"
	vetEP(4,1) = "UN-BSOL"
	vetEP(5,1) = "UN-ES"
	vetEP(6,1) = "UN-EXP"
	vetEP(7,1) = "UN-RIO"
	vetEP(8,1) = "UN-RNCE"
	vetEP(9,1) = "UN-SEAL"

	'E&P - Teste de Campo - Preparação
	vetEP(0,2) = ""
	vetEP(1,2) = "1886"
	vetEP(2,2) = ""
	vetEP(3,2) = "1888"
	vetEP(4,2) = ""
	vetEP(5,2) = ""
	vetEP(6,2) = ""
	vetEP(7,2) = ""
	vetEP(8,2) = "1893"
	vetEP(9,2) = ""

	'E&P  - Teste de Campo - Execução
	vetEP(0,3) = ""
	vetEP(1,3) = "5694"
	vetEP(2,3) = ""
	vetEP(3,3) = "5696"
	vetEP(4,3) = ""
	vetEP(5,3) = ""
	vetEP(6,3) = ""
	vetEP(7,3) = ""
	vetEP(8,3) = "5701"
	vetEP(9,3) = ""

	'E&P  - Saneamento
	vetEP(0,4) = "7085"
	vetEP(1,4) = "7086"
	vetEP(2,4) = "7087"
	vetEP(3,4) = "7088"
	vetEP(4,4) = "7089"
	vetEP(5,4) = "7090"
	vetEP(6,4) = "7091"
	vetEP(7,4) = "7092"
	vetEP(8,4) = "7093"
	vetEP(9,4) = "7094"

	'E&P - Treinamento
	vetEP(0,5) = "2393"
	vetEP(1,5) = "478"
	vetEP(2,5) = "479"
	vetEP(3,5) = "480"
	vetEP(4,5) = "481"
	vetEP(5,5) = "482"
	vetEP(6,5) = "483"
	vetEP(7,5) = "484"
	vetEP(8,5) = "485"
	vetEP(9,5) = "486"

	'E&P  - Comunicação - Teste de Campo
	vetEP(0,6) = "5724"
	vetEP(1,6) = "5725"
	vetEP(2,6) = "5719"
	vetEP(3,6) = "5720"
	vetEP(4,6) = "5721"
	vetEP(5,6) = "5722"
	vetEP(6,6) = "5717"
	vetEP(7,6) = "5718"
	vetEP(8,6) = "5716"
	vetEP(9,6) = "5715"

	'E&P  - Comunicação - Entrada em Produção
	vetEP(0,7) = "5901"
	vetEP(1,7) = "5896"
	vetEP(2,7) = "5897"
	vetEP(3,7) = "5898"
	vetEP(4,7) = "5893"
	vetEP(5,7) = "5894"
	vetEP(6,7) = "5895"
	vetEP(7,7) = "5713"
	vetEP(8,7) = "5892"
	vetEP(9,7) = "5891"


	'E&P  - Mobilização Prontidão
	vetEP(0,8) = "5800"
	vetEP(1,8) = "5801"
	vetEP(2,8) = "5797"
	vetEP(3,8) = "5798"
	vetEP(4,8) = "5795"
	vetEP(5,8) = "5796"
	vetEP(6,8) = "5793"
	vetEP(7,8) = "5794"
	vetEP(8,8) = "5792"
	vetEP(9,8) = "5791"



	'E&P  - Impacto
	vetEP(0,9) = "7097"
	vetEP(1,9) = "7098"
	vetEP(2,9) = "7099"
	vetEP(3,9) = "7100"
	vetEP(4,9) = "7101"
	vetEP(5,9) = "7102"
	vetEP(6,9) = "7103"
	vetEP(7,9) = "7104"
	vetEP(8,9) = "7105"
	vetEP(9,9) = "7106"



	'E&P  - Suporte - Cap. ALI's Mapeamento Perfil
	vetEP(0,10) = "2326"
	vetEP(1,10) = "1438"
	vetEP(2,10) = "1439"
	vetEP(3,10) = "1440"
	vetEP(4,10) = "1441"
	vetEP(5,10) = "1442"
	vetEP(6,10) = "1443"
	vetEP(7,10) = "1444"
	vetEP(8,10) = "1445"
	vetEP(9,10) = "1446"


	'E&P  - Suporte - Cap. ALI's Suporte a Usuario
	vetEP(0,11) = "6288"
	vetEP(1,11) = "6289"
	vetEP(2,11) = "6290"
	vetEP(3,11) = "6291"
	vetEP(4,11) = "6292"
	vetEP(5,11) = "6293"
	vetEP(6,11) = "6294"
	vetEP(7,11) = "6295"
	vetEP(8,11) = "6296"
	vetEP(9,11) = "6297"


	'E&P  - Suporte - Sala de Controle
	vetEP(0,12) = "7330"
	vetEP(1,12) = "7331"
	vetEP(2,12) = "7332"
	vetEP(3,12) = "7333"
	vetEP(4,12) = "7334"
	vetEP(5,12) = "7335"
	vetEP(6,12) = "7336"
	vetEP(7,12) = "7337"
	vetEP(8,12) = "7338"
	vetEP(9,12) = "7339"



	'E&P - INFRA - Geral
	vetEP(0,13) = "7107"
	vetEP(1,13) = "7108"
	vetEP(2,13) = "7109"
	vetEP(3,13) = "7110"
	vetEP(4,13) = "7111"
	vetEP(5,13) = "7112"
	vetEP(6,13) = "7113"
	vetEP(7,13) = "7114"
	vetEP(8,13) = "7115"
	vetEP(9,13) = "7116"



	'E&P - INFRA - NF
	vetEP(0,14) = "7259"
	vetEP(1,14) = "7260"
	vetEP(2,14) = "7261"
	vetEP(3,14) = "7262"
	vetEP(4,14) = "7263"
	vetEP(5,14) = "7264"
	vetEP(6,14) = "7265"
	vetEP(7,14) = "7266"
	vetEP(8,14) = "7267"
	vetEP(9,14) = "7268"



	'FINANC - Nomes
	vetFI(0,1) = "CONTABILIDADE"
	vetFI(1,1) = "FINANÇAS"
	vetFI(2,1) = "FINPROJ"
	vetFI(3,1) = "INVESTIDORES"
	vetFI(4,1) = "PLAFIN"
	vetFI(5,1) = "TRIBUTARIO"

	'FINANC - Teste de Campo - Preparação
	vetFI(0,2) = ""
	vetFI(1,2) = ""
	vetFI(2,2) = ""
	vetFI(3,2) = ""
	vetFI(4,2) = ""
	vetFI(5,2) = ""

	'FINANC - Teste de Campo - Execução
	vetFI(0,3) = ""
	vetFI(1,3) = ""
	vetFI(2,3) = ""
	vetFI(3,3) = ""
	vetFI(4,3) = ""
	vetFI(5,3) = ""

	'FINANC - Saneamento
	vetFI(0,4) = "7156"
	vetFI(1,4) = ""
	vetFI(2,4) = ""
	vetFI(3,4) = ""
	vetFI(4,4) = "7160"
	vetFI(5,4) = "7161"

	'FINANC - Treinamento
	vetFI(0,5) = "4436"
	vetFI(1,5) = "4441"
	vetFI(2,5) = "4443"
	vetFI(3,5) = "4446"
	vetFI(4,5) = "4449"
	vetFI(5,5) = "4453"

	'FINANC - Comunicação - Teste de Campo
	vetFI(0,6) = "5768"
	vetFI(1,6) = "5757"
	vetFI(2,6) = "5759"
	vetFI(3,6) = "5762"
	vetFI(4,6) = "5753"
	vetFI(5,6) = "5748"

	'FINANC - Comunicação - Entrada em Produção
	vetFI(0,7) = "5939"
	vetFI(1,7) = "5928"
	vetFI(2,7) = "5930"
	vetFI(3,7) = "5933"
	vetFI(4,7) = "5924"
	vetFI(5,7) = "5919"

	'FINANC - Mobilização Prontidão
	vetFI(0,8) = "5854"
	vetFI(1,8) = "5832"
	vetFI(2,8) = "5834"
	vetFI(3,8) = "5837"
	vetFI(4,8) = "5840"
	vetFI(5,8) = "5842"

	'FINANC - Impacto
	vetFI(0,9) = ""
	vetFI(1,9) = ""
	vetFI(2,9) = ""
	vetFI(3,9) = ""
	vetFI(4,9) = ""
	vetFI(5,9) = ""

	'FINANC - Suporte - Cap. ALI's Mapeamento Perfil
	vetFI(0,10) = "3518"
	vetFI(1,10) = "3523"
	vetFI(2,10) = "3525"
	vetFI(3,10) = "3528"
	vetFI(4,10) = "3531"
	vetFI(5,10) = "3535"

	'FINANC - Suporte - Cap. ALI's Suporte a Usuario
	vetFI(0,11) = "6326"
	vetFI(1,11) = "6331"
	vetFI(2,11) = "6332"
	vetFI(3,11) = "6335"
	vetFI(4,11) = "6338"
	vetFI(5,11) = "6342"

	'FINANC - Suporte - Sala de Controle
	vetFI(0,12) = "7353"
	vetFI(1,12) = "7354"
	vetFI(2,12) = "7355"
	vetFI(3,12) = "7356"
	vetFI(4,12) = "7357"
	vetFI(5,12) = "7358"

	'FINANC - INFRA - Geral
	vetFI(0,13) = "7168"
	vetFI(1,13) = "7169"
	vetFI(2,13) = "7170"
	vetFI(3,13) = "7171"
	vetFI(4,13) = "7172"
	vetFI(5,13) = "7173"

	'FINANC - INFRA - NF
	vetFI(0,14) = "7284"
	vetFI(1,14) = "7285"
	vetFI(2,14) = "7286"
	vetFI(3,14) = "7287"
	vetFI(4,14) = "7288"
	vetFI(5,14) = "7289"


	'GAS/ENER - Nomes
	vetGE(0,1) = "ENERGIA"
	vetGE(1,1) = "GAS-NATURAL"
	vetGE(2,1) = "ASSESSORIA"
	vetGE(3,1) = "CONSERVAÇÃO"
	vetGE(4,1) = "TCOM"

	'GAS/ENER - Teste de Campo - Preparação
	vetGE(0,2) = ""
	vetGE(1,2) = ""
	vetGE(2,2) = ""
	vetGE(3,2) = ""
	vetGE(4,2) = ""

	'GAS/ENER - Teste de Campo - Execução
	vetGE(0,3) = ""
	vetGE(1,3) = ""
	vetGE(2,3) = ""
	vetGE(3,3) = ""
	vetGE(4,3) = ""

	'GAS/ENER - Saneamento
	vetGE(0,4) = "7186"
	vetGE(1,4) = "7187"
	vetGE(2,4) = ""
	vetGE(3,4) = ""
	vetGE(4,4) = ""

	'GAS/ENER - Treinamento
	vetGE(0,5) = "4439"
	vetGE(1,5) = "4445"
	vetGE(2,5) = "7395"
	vetGE(3,5) = "7396"
	vetGE(4,5) = "7397"

	'GAS/ENER - Comunicação - Teste de Campo
	vetGE(0,6) = "5761"
	vetGE(1,6) = "5755"
	vetGE(2,6) = ""
	vetGE(3,6) = ""
	vetGE(4,6) = ""

	'GAS/ENER - Comunicação - Entrada em Produção
	vetGE(0,7) = "5926"
	vetGE(1,7) = "5932"
	vetGE(2,7) = ""
	vetGE(3,7) = ""
	vetGE(4,7) = ""

	'GAS/ENER - Mobilização Prontidão
	vetGE(0,8) = "5830"
	vetGE(1,8) = "5836"
	vetGE(2,8) = ""
	vetGE(3,8) = ""
	vetGE(4,8) = ""

	'GAS/ENER - Impacto
	vetGE(0,9) = ""
	vetGE(1,9) = ""
	vetGE(2,9) = ""
	vetGE(3,9) = ""
	vetGE(4,9) = ""

	'GAS/ENER - Suporte - Cap. ALI's Mapeamento Perfil
	vetGE(0,10) = "3521"
	vetGE(1,10) = "3527"
	vetGE(2,10) = ""
	vetGE(3,10) = ""
	vetGE(4,10) = ""

	'GAS/ENER - Suporte - Cap. ALI's Suporte a Usuario
	vetGE(0,11) = "6329"
	vetGE(1,11) = "6334"
	vetGE(2,11) = "7398"
	vetGE(3,11) = "7399"
	vetGE(4,11) = "7400"

	'GAS/ENER - Suporte - Sala de Controle
	vetGE(0,12) = "7363"
	vetGE(1,12) = "7364"
	vetGE(2,12) = ""
	vetGE(3,12) = ""
	vetGE(4,12) = ""

	'GAS/ENER - INFRA - Geral
	vetGE(0,13) = "7190"
	vetGE(1,13) = "7191"
	vetGE(2,13) = ""
	vetGE(3,13) = ""
	vetGE(4,13) = ""

	'GAS/ENER - INFRA - NF
	vetGE(0,14) = "7296"
	vetGE(1,14) = "7297"
	vetGE(2,14) = ""
	vetGE(3,14) = ""
	vetGE(4,14) = ""


	'INTERNAC - Nomes
	vetIN(0,1) = "INTER-PLAN"
	vetIN(1,1) = "INTER-EP"
	vetIN(2,1) = "INTER-GEE"
	vetIN(3,1) = "INTER-ABAST"

	'INTERNAC - Teste de Campo - Preparação
	vetIN(0,2) = ""
	vetIN(1,2) = "3343"
	vetIN(2,2) = ""
	vetIN(3,2) = ""

	'INTERNAC - Teste de Campo - Execução
	vetIN(0,3) = ""
	vetIN(1,3) = "3375"
	vetIN(2,3) = ""
	vetIN(3,3) = ""

	'INTERNAC - Saneamento
	vetIN(0,4) = "7174"
	vetIN(1,4) = ""
	vetIN(2,4) = ""
	vetIN(3,4) = ""

	'INTERNAC - Treinamento
	vetIN(0,5) = "4429"
	vetIN(1,5) = "4430"
	vetIN(2,5) = "4431"
	vetIN(3,5) = "4432"

	'INTERNAC - Comunicação - Teste de Campo
	vetIN(0,6) = "5771"
	vetIN(1,6) = "5772"
	vetIN(2,6) = "5763"
	vetIN(3,6) = "5764"

	'INTERNAC - Comunicação - Entrada em Produção
	vetIN(0,7) = "5943"
	vetIN(1,7) = "5944"
	vetIN(2,7) = "5934"
	vetIN(3,7) = "5935"

	'INTERNAC - Mobilização Prontidão
	vetIN(0,8) = "5847"
	vetIN(1,8) = "5848"
	vetIN(2,8) = "5849"
	vetIN(3,8) = "5850"

	'INTERNAC - Impacto
	vetIN(0,9) = ""
	vetIN(1,9) = ""
	vetIN(2,9) = ""
	vetIN(3,9) = ""

	'INTERNAC - Suporte - Cap. ALI's Mapeamento Perfil
	vetIN(0,10) = "3511"
	vetIN(1,10) = "3512"
	vetIN(2,10) = "3513"
	vetIN(3,10) = "3514"

	'INTERNAC - Suporte - Cap. ALI's Suporte a Usuario
	vetIN(0,11) = "6316"
	vetIN(1,11) = "6320"
	vetIN(2,11) = ""
	vetIN(3,11) = "6322"

	'INTERNAC - Suporte - Sala de Controle
	vetIN(0,12) = "7359"
	vetIN(1,12) = "7360"
	vetIN(2,12) = "7361"
	vetIN(3,12) = "7362"

	'INTERNAC - INFRA - Geral
	vetIN(0,13) = "7182"
	vetIN(1,13) = "7183"
	vetIN(2,13) = "7184"
	vetIN(3,13) = "7185"

	'INTERNAC - INFRA - NF
	vetIN(0,14) = "7291"
	vetIN(1,14) = "7292"
	vetIN(2,14) = "7293"
	vetIN(3,14) = "7294"


	'SERV - Nomes
	vetSE(0,1) = "CENPES"
	vetSE(1,1) = "COMP-NSM"
	vetSE(2,1) = "COMP-RNNE"
	vetSE(3,1) = "COMP-RSPS"
	vetSE(4,1) = "COMP-RSUD"
	vetSE(5,1) = "ENGENHARIA e EMPREENDIMENTOS"
	vetSE(6,1) = "MATERIAIS"
	vetSE(7,1) = "SMS"
	vetSE(8,1) = "TI"

	'SERV - Teste de Campo - Preparação
	vetSE(0,2) = ""
	vetSE(1,2) = ""
	vetSE(2,2) = ""
	vetSE(3,2) = ""
	vetSE(4,2) = ""
	vetSE(5,2) = ""
	vetSE(6,2) = ""
	vetSE(7,2) = ""
	vetSE(8,2) = ""

	'SERV - Teste de Campo - Execução
	vetSE(0,3) = ""
	vetSE(1,3) = ""
	vetSE(2,3) = ""
	vetSE(3,3) = ""
	vetSE(4,3) = ""
	vetSE(5,3) = ""
	vetSE(6,3) = ""
	vetSE(7,3) = ""
	vetSE(8,3) = ""

	'SERV - Saneamento
	vetSE(0,4) = "7192"
	vetSE(1,4) = "7193"
	vetSE(2,4) = "7194"
	vetSE(3,4) = "7195"
	vetSE(4,4) = "7196"
	vetSE(5,4) = "7198"
	vetSE(6,4) = "7199"
	vetSE(7,4) = "7200"
	vetSE(8,4) = "7257"


	'SERV - Treinamento
	vetSE(0,5) = "5455"
	vetSE(1,5) = "5456"
	vetSE(2,5) = "5457"
	vetSE(3,5) = "5458"
	vetSE(4,5) = "5459"
	vetSE(5,5) = "5461"
	vetSE(6,5) = "5462"
	vetSE(7,5) = "4452"
	vetSE(8,5) = "5465"

	'SERV - Comunicação - Teste de Campo
	vetSE(0,6) = "5784"
	vetSE(1,6) = "5785"
	vetSE(2,6) = "5786"
	vetSE(3,6) = "5780"
	vetSE(4,6) = "5781"
	vetSE(5,6) = "5783"
	vetSE(6,6) = "5778"
	vetSE(7,6) = "5750"
	vetSE(8,6) = "5776"


	'SERV - Comunicação - Entrada em Produção
	vetSE(0,7) = "5953"
	vetSE(1,7) = "5954"
	vetSE(2,7) = "5955"
	vetSE(3,7) = "5957"
	vetSE(4,7) = "5958"
	vetSE(5,7) = "5956"
	vetSE(6,7) = "5951"
	vetSE(7,7) = "5921"
	vetSE(8,7) = "5949"


	'SERV - Mobilização Prontidão
	vetSE(0,8) = "5861"
	vetSE(1,8) = "5862"
	vetSE(2,8) = "5863"
	vetSE(3,8) = "5864"
	vetSE(4,8) = "5865"
	vetSE(5,8) = "5866"
	vetSE(6,8) = "5867"
	vetSE(7,8) = "5856"
	vetSE(8,8) = "5870"



	'SERV - Impacto
	vetSE(0,9) = "7204"
	vetSE(1,9) = "7389"
	vetSE(2,9) = "7390"
	vetSE(3,9) = "7391"
	vetSE(4,9) = "7392"
	vetSE(5,9) = "7209"
	vetSE(6,9) = "7210"
	vetSE(7,9) = ""
	vetSE(8,9) = ""


	'SERV - Suporte - Cap. ALI's Mapeamento Perfil
	vetSE(0,10) = "5214"
	vetSE(1,10) = "5215"
	vetSE(2,10) = "5216"
	vetSE(3,10) = "5217"
	vetSE(4,10) = "5218"
	vetSE(5,10) = "5220"
	vetSE(6,10) = "5221"
	vetSE(7,10) = "3534"
	vetSE(8,10) = "5224"


	'SERV - Suporte - Cap. ALI's Suporte a Usuario
	vetSE(0,11) = "6345"
	vetSE(1,11) = "6346"
	vetSE(2,11) = "6347"
	vetSE(3,11) = "6348"
	vetSE(4,11) = "6349"
	vetSE(5,11) = "6350"
	vetSE(6,11) = "6351"
	vetSE(7,11) = "6341"
	vetSE(8,11) = "6354"


	'SERV - Suporte - Sala de Controle
	vetSE(0,12) = "7365"
	vetSE(1,12) = "7366"
	vetSE(2,12) = "7367"
	vetSE(3,12) = "7368"
	vetSE(4,12) = "7369"
	vetSE(5,12) = "7370"
	vetSE(6,12) = "7371"
	vetSE(7,12) = "7372"
	vetSE(8,12) = "7375"

	'SERV - INFRA - Geral
	vetSE(0,13) = "7216"
	vetSE(1,13) = "7217"
	vetSE(2,13) = "7218"
	vetSE(3,13) = "7219"
	vetSE(4,13) = "7220"
	vetSE(5,13) = "7221"
	vetSE(6,13) = "7222"
	vetSE(7,13) = "7224"
	vetSE(8,13) = "7227"


	'SERV - INFRA - NF
	vetSE(0,14) = "7316"
	vetSE(1,14) = "7317"
	vetSE(2,14) = "7318"
	vetSE(3,14) = "7319"
	vetSE(4,14) = "7320"
	vetSE(5,14) = "7321"
	vetSE(6,14) = "7322"
	vetSE(7,14) = "7324"
	vetSE(8,14) = "7327"



	'PRESIDENCIA - Nomes
	vetPR(0,1) = "AUDITORIA"
	vetPR(1,1) = "COM INST"
	vetPR(2,1) = "DESEMPENHO"
	vetPR(3,1) = "DSG"
	vetPR(4,1) = "ESTRATEGIA"
	vetPR(5,1) = "GAPRE e ASS.PRESIDENCIA"
	vetPR(6,1) = "JURIDICO"
	vetPR(7,1) = "NOVOS NEGOCIOS"
	vetPR(8,1) = "SEGEPE"
	vetPR(9,1) = "RH"

	'PRESIDENCIA - Teste de Campo - Preparação
	vetPR(0,2) = ""
	vetPR(1,2) = ""
	vetPR(2,2) = ""
	vetPR(3,2) = ""
	vetPR(4,2) = ""
	vetPR(5,2) = ""
	vetPR(6,2) = ""
	vetPR(7,2) = ""
	vetPR(8,2) = ""
	vetPR(9,2) = ""

	'PRESIDENCIA - Teste de Campo - Execução
	vetPR(0,3) = ""
	vetPR(1,3) = ""
	vetPR(2,3) = ""
	vetPR(3,3) = ""
	vetPR(4,3) = ""
	vetPR(5,3) = ""
	vetPR(6,3) = ""
	vetPR(7,3) = ""
	vetPR(8,3) = ""
	vetPR(9,3) = ""

	'PRESIDENCIA - Saneamento
	vetPR(0,4) = ""
	vetPR(1,4) = ""
	vetPR(2,4) = ""
	vetPR(3,4) = ""
	vetPR(4,4) = ""
	vetPR(5,4) = ""
	vetPR(6,4) = ""
	vetPR(7,4) = ""
	vetPR(8,4) = ""
	vetPR(9,4) = ""

	'PRESIDENCIA - Treinamento
	vetPR(0,5) = "4433"
	vetPR(1,5) = "4434"
	vetPR(2,5) = "4437"
	vetPR(3,5) = "4438"
	vetPR(4,5) = "4440"
	vetPR(5,5) = "4444"
	vetPR(6,5) = "4447"
	vetPR(7,5) = "4448"
	vetPR(8,5) = "4451"
	vetPR(9,5) = "4450"


	'PRESIDENCIA - Comunicação - Teste de Campo
	vetPR(0,6) = "5765"
	vetPR(1,6) = "5766"
	vetPR(2,6) = "5769"
	vetPR(3,6) = "5770"
	vetPR(4,6) = "5756"
	vetPR(5,6) = "5760"
	vetPR(6,6) = "5751"
	vetPR(7,6) = "5752"
	vetPR(8,6) = "5749"
	vetPR(9,6) = "5754"


	'PRESIDENCIA - Comunicação - Entrada em Produção
	vetPR(0,7) = ""
	vetPR(0,7) = "5936"
	vetPR(1,7) = "5937"
	vetPR(2,7) = "5940"
	vetPR(3,7) = "5941"
	vetPR(4,7) = "5927"
	vetPR(5,7) = "5931"
	vetPR(6,7) = "5922"
	vetPR(7,7) = "5923"
	vetPR(8,7) = "5920"
	vetPR(9,7) = "5925"


	'PRESIDENCIA - Mobilização Prontidão
	vetPR(0,8) = "5851"
	vetPR(1,8) = "5852"
	vetPR(2,8) = "5828"
	vetPR(3,8) = "5829"
	vetPR(4,8) = "5831"
	vetPR(5,8) = "5835"
	vetPR(6,8) = "5838"
	vetPR(7,8) = "5839"
	vetPR(8,8) = "5855"
	vetPR(9,8) = "5841"


	'PRESIDENCIA - Impacto
	vetPR(0,9) = ""
	vetPR(1,9) = ""
	vetPR(2,9) = ""
	vetPR(3,9) = ""
	vetPR(4,9) = ""
	vetPR(5,9) = ""
	vetPR(6,9) = ""
	vetPR(7,9) = ""
	vetPR(8,9) = ""
	vetPR(9,9) = ""


	'PRESIDENCIA - Suporte - Cap. ALI's Mapeamento Perfil
	vetPR(0,10) = "3515"
	vetPR(1,10) = "3516"
	vetPR(2,10) = "3519"
	vetPR(3,10) = "3520"
	vetPR(4,10) = "3522"
	vetPR(5,10) = "3526"
	vetPR(6,10) = "3529"
	vetPR(7,10) = "3530"
	vetPR(8,10) = "3533"
	vetPR(9,10) = "3532"

	'PRESIDENCIA - Suporte - Cap. ALI's Suporte a Usuario
	vetPR(0,11) = "6323"
	vetPR(1,11) = "6324"
	vetPR(2,11) = "6327"
	vetPR(3,11) = "6328"
	vetPR(4,11) = ""
	vetPR(5,11) = "6333"
	vetPR(6,11) = "6336"
	vetPR(7,11) = "6337"
	vetPR(8,11) = "6340"
	vetPR(9,11) = "6339"


	'PRESIDENCIA - Suporte - Sala de Controle
	vetPR(0,12) = "7377"
	vetPR(1,12) = "7378"
	vetPR(2,12) = "7379"
	vetPR(3,12) = "7380"
	vetPR(4,12) = "7381"
	vetPR(5,12) = "7382"
	vetPR(6,12) = "7383"
	vetPR(7,12) = "7384"
	vetPR(8,12) = "7385"
	vetPR(9,12) = "7386"

	'PRESIDENCIA - INFRA - Geral
	vetPR(0,13) = "7246"
	vetPR(1,13) = "7247"
	vetPR(2,13) = "7248"
	vetPR(3,13) = "7249"
	vetPR(4,13) = "7250"
	vetPR(5,13) = "7251"
	vetPR(6,13) = "7252"
	vetPR(7,13) = "7253"
	vetPR(8,13) = "7254"
	vetPR(9,13) = "7223"


	'PRESIDENCIA - INFRA - NF
	vetPR(0,14) = "7299"
	vetPR(1,14) = "7300"
	vetPR(2,14) = "7301"
	vetPR(3,14) = "7302"
	vetPR(4,14) = "7303"
	vetPR(5,14) = "7304"
	vetPR(6,14) = "7305"
	vetPR(7,14) = "7306"
	vetPR(8,14) = "7307"
	vetPR(9,14) = "7323"


%>

	<html>

	<head>
	<title>Projeto Sinergia </title>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
	<FORM name="frmRelatorio_Medicao_Detalhe" id="frmRelatorio_Medicao_Detalhe" action="GVI_Relatorio_Medicao_Detalhe.asp" method="post">
	<link rel="stylesheet" href="estilos/sinergia.css">
	<SCRIPT language=JavaScript SRC="scripts/valida.js"></SCRIPT>
		<p>
		
		<table width="100%" border="0">
			<tr>
				<td width="30%">&nbsp;</td>
				<td width="30%" align="middle">
					<p><b>
					<font color="#666666" size="3" face="Georgia, Times New Roman, Times, serif">Painel Sinóptico</font>
					</b></p>
				</td>
				<td width="30%">&nbsp;</td>
			</tr>
		</table>
		<BR>
		<table style="border-style: solid; border-width: 1" border="0" cellspacing="0" cellpadding="0" align=center>
				  
			<tr height="17" style="height:12.75pt">

				<td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Diretoria/Unidade</font></b></td>

				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
				<%If  trim(strAtividade) = "TC" Then%>

					<td class="xl27" width="120px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc colSpan=2><b>
				        <a href="downloads/Criterio_TCampo_Comun_Sup_Imp.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Teste de Campo</font></a></b></td>

					<%If trim(strAtividade) = "TC" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>

				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				        <a href="downloads/Criterio_Saneamento.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Saneamento</font></a></b></td>

					<%If trim(strAtividade) = "SA" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>

						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>

					<%End If%>
-->
				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				        <a href="downloads/Criterio_Treinamento.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Treinamento</font></a></b></td>

					<%If trim(strAtividade) = "TR" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>

				<%End If%>


				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
				<%If trim(strAtividade) = "CO" Then%>
					<td class="xl27" width="120px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc colSpan=2><b>
				        <a href="downloads/Criterio_TCampo_Comun_Sup_Imp.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Comunicação</font></a></b></td>

					<%If trim(strAtividade) = "CO" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>

				<%End If%>

				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
				<%If trim(strAtividade) = "MO" Then%>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Mobilização Prontidão</font></b></td>

					<%If trim(strAtividade) = "MO" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>

				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				        <a href="downloads/Criterio_TCampo_Comun_Sup_Imp.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Impactos</font></a></b></td>

					<%If trim(strAtividade) = "IM" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>-->

				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" Then%>
					<td class="xl27" width="180px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc colSpan=2><b>
				        <a href="downloads/Criterio_TCampo_Comun_Sup_Imp.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Suporte</font></a></b></td>
				<%End If%>


				<%If trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
				        <a href="downloads/Criterio_TCampo_Comun_Sup_Imp.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Suporte</font></a></b></td>
					<%If trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
					<%End If%>
				<%End If%>


				<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" Then%>
					<td class="xl27" width="120px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc colSpan=2><b>
				        <a href="downloads/Criterio_Infra.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Infra</font></a></b></td>
				<%End If%>

				<%If trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc ><b>
					    <a href="downloads/Criterio_Infra.ppt" target="_blank" class="conf">
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Infra</font></a></b></td>

					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Previsto</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Realizado</font></b></td>
				<%End If%>


			</tr>
				

			<tr height="17" style="height:12.75pt">

				<td class="xl27" width="200px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
					<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
				<%If  trim(strAtividade) = "TC" Then%>
				
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc>
						<b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
								Preparação
							</font>
						</b>
					</td>

					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc>
						<b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">
								Execução
							</font>
						</b>
					</td>

					<%If trim(strAtividade) = "TC" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%End If%>

				<%End If%>


				<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%If trim(strAtividade) = "SA" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

					<%End If%>-->
				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

					<%If trim(strAtividade) = "TR" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%End If%>


				<%End If%>

				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
				<%If trim(strAtividade) = "CO" Then%>
				
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Teste de Campo</font></b></td>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Entrada em Produção</font></b></td>

					<%If trim(strAtividade) = "CO" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%End If%>


				<%End If%>

				<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
				<%If trim(strAtividade) = "MO" Then%>
			
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

					<%If trim(strAtividade) = "MO" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%End If%>

				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>

					<%If trim(strAtividade) = "IM" Then%>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
						<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
							<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<%End If%>-->

				<%End If%>

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" Then%>
					<td class="xl27" width="90px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cap. ALI's Perfil</font></b></td>
					<td class="xl27" width="90px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cap. ALI's Suporte Usuário</font></b></td>
<!--					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Suporte e Controle</font></b></td>-->
				<%End If%>
				
				<%If trim(strAtividade) = "SP-CAP-PER" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cap. ALI's Perfil</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
				<%End If%>

				<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Cap. ALI's Suporte Usuário</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
				<%End If%>

				<%If trim(strAtividade) = "SP-SL-CTR" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Suporte e Controle</font></b></td>
<!--					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>-->
				<%End If%>
		

				<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Geral</font></b></td>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">NF</font></b></td>
				<%End If%>

				<%If trim(strAtividade) = "IN-GERAL" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">Geral</font></b></td>

					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
				<%End If%>


				<%If trim(strAtividade) = "IN-NF" Then%>
					<td class="xl27" width="60px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">NF</font></b></td>

					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
					<td class="xl27" width="80px" style="border: 1 solid #666666" align="center" bgcolor=#6699cc><b>
						<font color=White size="1" face="Georgia, Times New Roman, Times, serif">&nbsp;</font></b></td>
				<%End If%>




			</tr>


		<%'ABAST%>
		
			<%If strDiretoria = "" or strDiretoria = "AB" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(0,1),"/","_")%>.ppt" target="_blank" class="conf">
					  		<font face="Arial" size="1" color=Black>
					  			<%=vetCS(0,1)%>
							</font>
						</a>
					  </td>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If trim(strAtividade) = "TC" Then%>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,2))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,4))%>&nbsp;
								</td>
						<%End If%>-->

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,5))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,8))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(0,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,9))%>&nbsp;
								</td>
						<%End If%>-->

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(0,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(0,10))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(0,11))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,11))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

<!--						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(0,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>-->


						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
<!--							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,12))%>&nbsp;
								</td>-->
						<%End If%>

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(0,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,13))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(0,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(0,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>

						<%End If%>

						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(0,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(0,14))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "AB" Then%>

				<%For I = 0  to ubound(vetAB)%>
					<tr height="17" style="height:12.75pt">
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
							<a href="Apres_Unidades/<%=REPLACE(vetAB(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1"  color=Black>
						  			&nbsp;&nbsp;<%=vetAB(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If trim(strAtividade) = "TC" Then%>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,2))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetAB(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP"  OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP"  OR trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetAB(I,10))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP2",vetAB(I,10))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP"  OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetAB(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

<!--							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetAB(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->

							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,12))%>&nbsp;
									</td>-->
							<%End If%>

						<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetAB(I,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,13))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>

							  		</font>
							  	</td>
							<%End If%>


							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
						  					<%strData = RetornaDataPainelSinoptico(vetAB(I,14))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
						  						<%strResultado = RetornaValorPainelSinoptico("TP1",vetAB(I,14))
						  						Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>

										</font>
									</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,13))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetAB(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetAB(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'E&P%>
			<%If strDiretoria = "" or strDiretoria = "EP" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(1,1),"/","_")%>.ppt" target="_blank" class="conf">
					  		<font face="Arial" size="1" color=Black>
					  			<%=vetCS(1,1)%>
						  	</font>
						</a>
					  </td>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>



					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>

					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>

					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">

					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>

					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">

					  			<%strData = RetornaDataPainelSinoptico(vetCS(1,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>

					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP"  OR trim(strAtividade) = "SP-CAP-PER" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

						  			<%strData = RetornaDataPainelSinoptico(vetCS(1,10))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
						  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(1,10))
						  				Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP"  OR trim(strAtividade) = "SP-CAP-SUP" Then%>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

						  			<%strData = RetornaDataPainelSinoptico(vetCS(1,11))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
						  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,11))
						  				Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						<%End If%>

						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

						  			<%strData = RetornaDataPainelSinoptico(vetCS(1,12))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
						  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,12))
						  				Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						<%End If%>-->

						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,12))%>&nbsp;
								</td>-->
						<%End If%>



					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(1,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,13))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>


									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(1,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(1,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>

									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(1,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(1,14))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "EP" Then%>

				<%For I = 0  to ubound(vetEP)%>
					<tr height="17" style="height:12.75pt"  >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
							<a href="Apres_Unidades/<%=REPLACE(vetEP(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1"  color=Black>
						  			&nbsp;&nbsp;<%=vetEP(I,1)%>
						  		</font>
							</a>
						  </td>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,2))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
	  		
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>



						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetEP(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
					  					<%strData = RetornaDataPainelSinoptico(vetEP(I,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
					  						<%strResultado = RetornaValorPainelSinoptico("TP2",vetEP(I,10))
					  						Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetEP(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetEP(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->

							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,12))%>&nbsp;
									</td>-->
							<%End If%>

						<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetEP(I,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,13))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>

							  		</font>
							  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
								  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
								  		<font face="Arial" size="1">
							  				<%strData = RetornaDataPainelSinoptico(vetEP(I,14))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
							  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetEP(I,14))
							  					Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>

								  		</font>
								  	</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,13))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetEP(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetEP(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'FINANC%>
		
			<%If strDiretoria = "" or strDiretoria = "FI" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(2,1),"/","_")%>.ppt" target="_blank" class="conf">
						  	<font face="Arial" size="1"  color=Black>
						  		<%=vetCS(2,1)%>
						  	</font>
					  	</a>
					  </td>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>



					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(2,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>

					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER"Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

						  			<%strData = RetornaDataPainelSinoptico(vetCS(2,10))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
						  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(2,10))
						  				Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

						  			<%strData = RetornaDataPainelSinoptico(vetCS(2,11))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
						  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,11))
						  				Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						<%End If%>

						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">

										<%strData = RetornaDataPainelSinoptico(vetCS(2,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>

									</font>
								</td>
						<%End If%>-->

						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,12))%>&nbsp;
								</td>-->
						<%End If%>

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  	<font face="Arial" size="1">
							  		<%strData = RetornaDataPainelSinoptico(vetCS(2,13))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
							  			<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,13))
							  			Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
					  			</font>
					  		</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(2,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(2,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>

									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(2,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(2,14))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "FI" Then%>

				<%For I = 0  to ubound(vetFI)%>
					<tr height="17" style="height:12.75pt" >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
							<a href="Apres_Unidades/<%=REPLACE(vetFI(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1"  color=Black>
						  			&nbsp;&nbsp;<%=vetFI(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,2))%>&nbsp;
									</td>
							<%End If%>

						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">

					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">

					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">

					  				<%strData = RetornaDataPainelSinoptico(vetFI(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>

						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
					  					<%strData = RetornaDataPainelSinoptico(vetFI(I,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
					  						<%strResultado = RetornaValorPainelSinoptico("TP2",vetFI(I,10))
					  						Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetFI(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>


							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetFI(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->


							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,12))%>&nbsp;
									</td>-->
							<%End If%>


						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
								  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
								  		<font face="Arial" size="1">
							  				<%strData = RetornaDataPainelSinoptico(vetFI(I,13))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
							  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,13))
							  					Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
								  		</font>
								  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
								  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
								  		<font face="Arial" size="1">
							  				<%strData = RetornaDataPainelSinoptico(vetFI(I,14))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
							  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetFI(I,14))
							  					Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
								  		</font>
								  	</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,13))%>&nbsp;
									</td>
							<%End If%>


							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetFI(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetFI(I,14))%>&nbsp;
									</td>
							<%End If%>
						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'GAS/ENER%>
		
			<%If strDiretoria = "" or strDiretoria = "GE" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(3,1),"/","_")%>.ppt" target="_blank" class="conf">
						  	<font face="Arial" size="1"  color=Black>
						  		<%=vetCS(3,1)%>
						  	</font>
					  	</a>
					  </td>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(3,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(3,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(3,10))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(3,11))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,11))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(3,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>-->

						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,12))%>&nbsp;
								</td>-->
						<%End If%>



					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(3,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,13))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(3,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(3,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(3,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(3,14))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "GE" Then%>

				<%For I = 0  to ubound(vetGE)%>
					<tr height="17" style="height:12.75pt" >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
							<a href="Apres_Unidades/<%=REPLACE(vetGE(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1"  color=Black>
						  			&nbsp;&nbsp;<%=vetGE(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,2))%>&nbsp;
									</td>
							<%End If%>


						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetGE(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetGE(I,10))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP2",vetGE(I,10))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetGE(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetGE(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->

							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,12))%>&nbsp;
									</td>-->
							<%End If%>

					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetGE(I,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,13))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetGE(I,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetGE(I,14))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>


							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,13))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetGE(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetGE(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'INTERNAC%>
		
			<%If strDiretoria = "" or strDiretoria = "IN" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(4,1),"/","_")%>.ppt" target="_blank" class="conf">
						  	<font face="Arial" size="1" color=Black>
					  			<%=vetCS(4,1)%>
						  	</font>
				  		</a>
					  </td>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(4,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(4,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(4,10))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(4,11))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,11))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(4,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>-->


						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,12))%>&nbsp;
								</td>-->
						<%End If%>



					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(4,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,13))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(4,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(4,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(4,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(4,14))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "IN" Then%>

				<%For I = 0  to ubound(vetIN)%>
					<tr height="17" style="height:12.75pt" >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
							<a href="Apres_Unidades/<%=REPLACE(vetIN(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1" color=Black>
						  			&nbsp;&nbsp;<%=vetIN(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,2))%>&nbsp;
									</td>
							<%End If%>


						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>
						
						
						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>



						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetIN(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetIN(I,10))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP2",vetIN(I,10))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetIN(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetIN(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->


							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,12))%>&nbsp;
									</td>-->
							<%End If%>


						<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetIN(I,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,13))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
							  		<font face="Arial" size="1">
						  				<%strData = RetornaDataPainelSinoptico(vetIN(I,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
						  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetIN(I,14))
						  					Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
							  		</font>
							  	</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,13))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetIN(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetIN(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'SERV%>
		
			<%If strDiretoria = "" or strDiretoria = "SE" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(5,1),"/","_")%>.ppt" target="_blank" class="conf">
						  	<font face="Arial" size="1" color=Black>
						  		<%=vetCS(5,1)%>
						  	</font>
					  	</a>
					  </td>



					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(5,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(5,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(5,10))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(5,11))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,11))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>
					  	
						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(5,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>-->


						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,12))%>&nbsp;
								</td>-->
						<%End If%>



					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(5,13))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,13))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(5,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(5,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(5,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(5,14))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "SE" Then%>

				<%For I = 0  to ubound(vetSE)%>
					<tr height="17" style="height:12.75pt" >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
							<a href="Apres_Unidades/<%=REPLACE(vetSE(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1" color=Black>
						  			&nbsp;&nbsp;<%=vetSE(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,2))%>&nbsp;
									</td>
							<%End If%>


						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "SA" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetSE(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetSE(I,10))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP2",vetSE(I,10))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>


							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetSE(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>


							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetSE(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->


							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,12))%>&nbsp;
									</td>-->
							<%End If%>



						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
								  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
								  		<font face="Arial" size="1">
							  				<%strData = RetornaDataPainelSinoptico(vetSE(I,13))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
							  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,13))
							  					Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
								  		</font>
								  	</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
								  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
								  		<font face="Arial" size="1">
							  				<%strData = RetornaDataPainelSinoptico(vetSE(I,14))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
							  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetSE(I,14))
							  					Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
								  		</font>
								  	</td>
							<%End If%>

							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,13))%>&nbsp;
									</td>
							<%End If%>


							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetSE(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetSE(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>


		<%'PRESIDENCIA%>
		
			<%If strDiretoria = "" or strDiretoria = "PR" or strDiretoria = "CO" Then%>
				<tr height="17" style="height:12.75pt" bgcolor=LightGrey>
					<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=200px>
						<a href="Apres_Unidades/<%=REPLACE(vetCS(6,1),"/","_")%>.ppt" target="_blank" class="conf">
						  	<font face="Arial" size="1"  color=Black>
						  		<%=vetCS(6,1)%>
						  	</font>
					  	</a>
					  </td>



					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
					<%If  trim(strAtividade) = "TC" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,2))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,2))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,3))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,3))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "TC" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,2))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,2))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>


					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--	  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,4))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,4))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,4))%>&nbsp;
								</td>

							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,4))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,5))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,5))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "TR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,5))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,5))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>


					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
					<%If trim(strAtividade) = "CO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,6))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,6))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

					  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,7))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,7))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "CO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,6))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,6))%>&nbsp;
								</td>
						<%End If%>

					<%End If%>

					<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
					<%If trim(strAtividade) = "MO" Then%>
					
					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,8))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,8))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>
					  	
						<%If trim(strAtividade) = "MO" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,8))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,8))%>&nbsp;
								</td>
						<%End If%>
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--					  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
					  		<font face="Arial" size="1">
					  			<%strData = RetornaDataPainelSinoptico(vetCS(6,9))%>
								<%If cdate(strData) > date() Then%>
									<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
								<%Else%>
					  				<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,9))
					  				Select Case strResultado%>
									<%  Case "1"%>
											<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
									<%  Case "2"%>
											<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
									<%  Case "3"%>
											<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
									<%  Case Else%>
											<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
									<%End Select%>
								<%End If%>
					  		</font>
					  	</td>

						<%If trim(strAtividade) = "IM" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,9))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,9))%>&nbsp;
								</td>
						<%End If%>-->
					  	
					<%End If%>

					<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(6,10))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP2",vetCS(6,10))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(6,11))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,11))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>
					  	
						<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(6,12))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,12))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>-->


						<%If trim(strAtividade) = "SP-CAP-PER" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,10))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,10))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,11))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,11))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "SP-SL-CTR" Then%>
							<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,12))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,12))%>&nbsp;
								</td>-->
						<%End If%>



					<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
											<%strData = RetornaDataPainelSinoptico(vetCS(6,13))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
												<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,13))
												Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
									<font face="Arial" size="1">
										<%strData = RetornaDataPainelSinoptico(vetCS(6,14))%>
										<%If cdate(strData) > date() Then%>
											<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
										<%Else%>
											<%strResultado = RetornaValorPainelSinoptico("TP1",vetCS(6,14))
											Select Case strResultado%>
											<%  Case "1"%>
													<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
											<%  Case "2"%>
													<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
											<%  Case "3"%>
													<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
											<%  Case Else%>
													<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
											<%End Select%>
										<%End If%>
									</font>
								</td>
						<%End If%>


						<%If trim(strAtividade) = "IN-GERAL" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,13))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,13))%>&nbsp;
								</td>
						<%End If%>

						<%If trim(strAtividade) = "IN-NF" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP1",vetCS(6,14))%>&nbsp;
								</td>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
									<%=RetornaPrevRealPainelSinoptico("TP2",vetCS(6,14))%>&nbsp;
								</td>
						<%End If%>


					<%End If%>

				</tr>
			<%End If%>

			<%If strDiretoria = "" or strDiretoria = "PR" Then%>

				<%For I = 0  to ubound(vetPR)%>
					<tr height="17" style="height:12.75pt" >
						<td height="17" class="xl22" align=left style="border: 1 solid #666666" width=250px>
							<a href="Apres_Unidades/<%=REPLACE(vetPR(I,1),"/","_")%>.ppt" target="_blank" class="conf">
						  		<font face="Arial" size="1" color=Black>
						  			&nbsp;&nbsp;<%=vetPR(I,1)%>
						  		</font>
							</a>
						  </td>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "TC" Then%>
						<%If  trim(strAtividade) = "TC" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,2))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,2))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,3))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,3))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "TC" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,2))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,2))%>&nbsp;
									</td>
							<%End If%>


						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SA" Then%>
<!--		  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,4))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,4))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "SA" Then%>
							<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,4))%>&nbsp;
									</td>

								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,4))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>


						<%If trim(strAtividade) = "" OR trim(strAtividade) = "TR" Then%>
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,5))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,5))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "TR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,5))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,5))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>


						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "CO" Then%>
						<%If trim(strAtividade) = "CO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,6))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,6))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

						  <td class="xl23" align=center style="border: 1 solid #666666" width="60px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,7))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,7))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>

							<%If trim(strAtividade) = "CO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,6))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,6))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

						<%'If trim(strAtividade) = "" OR trim(strAtividade) = "MO" Then%>
						<%If trim(strAtividade) = "MO" Then%>
						
						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,8))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,8))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "MO" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,8))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,8))%>&nbsp;
									</td>
							<%End If%>
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IM" Then%>
<!--						  <td class="xl23" align=center style="border: 1 solid #666666" width="80px">
						  		<font face="Arial" size="1">
					  				<%strData = RetornaDataPainelSinoptico(vetPR(I,9))%>
									<%If cdate(strData) > date() Then%>
										<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
									<%Else%>
					  					<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,9))
					  					Select Case strResultado%>
										<%  Case "1"%>
												<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
										<%  Case "2"%>
												<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
										<%  Case "3"%>
												<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
										<%  Case Else%>
												<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
										<%End Select%>
									<%End If%>
						  		</font>
						  	</td>
						  	
							<%If trim(strAtividade) = "IM" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,9))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,9))%>&nbsp;
									</td>
							<%End If%>-->
						  	
						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" OR trim(strAtividade) = "SP-CAP-SUP" OR trim(strAtividade) = "SP-SL-CTR" Then%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetPR(I,10))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP2",vetPR(I,10))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetPR(I,11))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,11))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>

							<!--<%If trim(strAtividade) = "" OR trim(strAtividade) = "SP" OR trim(strAtividade) = "SP-SL-CTR" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetPR(I,12))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,12))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>-->


							<%If trim(strAtividade) = "SP-CAP-PER" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,10))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,10))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-CAP-SUP" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,11))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,11))%>&nbsp;
									</td>
							<%End If%>

							<%If trim(strAtividade) = "SP-SL-CTR" Then%>
								<!--<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,12))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,12))%>&nbsp;
									</td>-->
							<%End If%>


						<%End If%>

						<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" OR trim(strAtividade) = "IN-NF" Then%>
							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetPR(I,13))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,13))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>


							<%If trim(strAtividade) = "" OR trim(strAtividade) = "IN" OR trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="60px">
										<font face="Arial" size="1">
					  						<%strData = RetornaDataPainelSinoptico(vetPR(I,14))%>
											<%If cdate(strData) > date() Then%>
												<img src="icones/cinza_pv.gif" alt="No Stress!" name="Bola1" border="0" width=18>
											<%Else%>
					  							<%strResultado = RetornaValorPainelSinoptico("TP1",vetPR(I,14))
					  							Select Case strResultado%>
												<%  Case "1"%>
														<img src="icones/verde_pv.gif" alt="Muito Bom!" name="Bola1" border="0" width=18>
												<%  Case "2"%>
														<img src="icones/amarelo_pv.gif" alt="Perigo!" name="Bola2" border="0" width=18>
												<%  Case "3"%>
														<img src="icones/vermelho_pv.gif" alt="OOOPS!" name="Bola3" border="0" width=18>
												<%  Case Else%>
														<img src="icones/linea.gif" alt="Não se Aplica!" name="Bola3" border="0" width=18>
												<%End Select%>
											<%End If%>
										</font>
									</td>
							<%End If%>


							<%If trim(strAtividade) = "IN-GERAL" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,13))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,13))%>&nbsp;
									</td>
							<%End If%>


							<%If trim(strAtividade) = "IN-NF" Then%>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP1",vetPR(I,14))%>&nbsp;
									</td>
								<td class="xl23" align=center style="border: 1 solid #666666" width="80px">
										<%=RetornaPrevRealPainelSinoptico("TP2",vetPR(I,14))%>&nbsp;
									</td>
							<%End If%>

						<%End If%>

					</tr>
				<%Next%>
			<%End If%>

		</table>

		<p align=center>
		<BR>
		<hr>

</FORM>
</body>
</html>