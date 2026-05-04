Attribute VB_Name = "mdlGobal"
Option Explicit

Public dbConexaoSMR     As ADODB.Connection
Public dbConexaoSMRTRAB As ADODB.Connection
Public dbConexaoTIN     As ADODB.Connection
Public dbConexaoEP      As ADODB.Connection
Public dbConexaoTCP     As ADODB.Connection
Public dbConexaoCUT     As ADODB.Connection
Public dbConexaoSC      As ADODB.Connection


Public Sub gsFecharConexaoBDSMR()

    If Not (dbConexaoSMR Is Nothing) Then
        If dbConexaoSMR.State = adStateOpen Then
            dbConexaoSMR.Close
            Set dbConexaoSMR = Nothing
        End If
    End If

End Sub


Public Sub gsFecharConexaoBDSMRTRAB()

'    If Not (dbConexaoSMRTRAB Is Nothing) Then
'        If dbConexaoSMRTRAB.State = adStateOpen Then
'            dbConexaoSMRTRAB.Close
'            Set dbConexaoSMRTRAB = Nothing
'        End If
'    End If

End Sub


Public Sub gsFecharConexaoBDTIN()

    If Not (dbConexaoTIN Is Nothing) Then
        If dbConexaoTIN.State = adStateOpen Then
            dbConexaoTIN.Close
            Set dbConexaoTIN = Nothing
        End If
    End If

End Sub

Public Sub gsFecharConexaoBDEP()

    If Not (dbConexaoEP Is Nothing) Then
        If dbConexaoEP.State = adStateOpen Then
            dbConexaoEP.Close
            Set dbConexaoEP = Nothing
        End If
    End If

End Sub


Public Sub gsFecharConexaoBDTCP()

'    If Not (dbConexaoTCP Is Nothing) Then
'        If dbConexaoTCP.State = adStateOpen Then
'            dbConexaoTCP.Close
'            Set dbConexaoTCP = Nothing
'        End If
'    End If

End Sub

Public Sub gsFecharConexaoBDCUT()

    If Not (dbConexaoCUT Is Nothing) Then
        If dbConexaoCUT.State = adStateOpen Then
            dbConexaoCUT.Close
            Set dbConexaoCUT = Nothing
        End If
    End If

End Sub

Public Sub gsFecharConexaoBDSC()

    If Not (dbConexaoSC Is Nothing) Then
        If dbConexaoSC.State = adStateOpen Then
            dbConexaoSC.Close
            Set dbConexaoSC = Nothing
        End If
    End If

End Sub

Public Sub gsAbrirConexaoBDSMR()

    Dim strConexao As String
    
    Set dbConexaoSMR = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_SMR.udl;"

    'Abre a conexão
    dbConexaoSMR.Open strConexao
    dbConexaoSMR.CommandTimeout = 1800


End Sub


Public Sub gsAbrirConexaoBDSMRTRAB()

'    Dim strConexao As String
'
'    Set dbConexaoSMRTRAB = New ADODB.Connection
'
'    'Constrói a ConnectionString para abertura da conexão
'    strConexao = "File Name=" & App.Path & "\Carga_SMRTRAB.udl;"
'
'    'Abre a conexão
'    dbConexaoSMRTRAB.Open strConexao
'    dbConexaoSMRTRAB.CommandTimeout = 1800


End Sub

Public Sub gsAbrirConexaoBDTIN()

    Dim strConexao As String
    
    Set dbConexaoTIN = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_TIN.udl;"

    'Abre a conexão
    dbConexaoTIN.Open strConexao
    dbConexaoTIN.CommandTimeout = 1800


End Sub

Public Sub gsAbrirConexaoBDEP()

    Dim strConexao As String
    
    Set dbConexaoEP = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_EP.udl;"

    'Abre a conexão
    dbConexaoEP.Open strConexao
    dbConexaoEP.CommandTimeout = 1800


End Sub

Public Sub gsAbrirConexaoBDTCP()

'    Dim strConexao As String
'
'    Set dbConexaoTCP = New ADODB.Connection
'
'    'Constrói a ConnectionString para abertura da conexão
'    strConexao = "File Name=" & App.Path & "\Carga_TCP.udl;"
'
'    'Abre a conexão
'    dbConexaoTCP.Open strConexao
'    dbConexaoTCP.CommandTimeout = 1800

End Sub


Public Sub gsAbrirConexaoBDCUT()

    Dim strConexao As String
    
    Set dbConexaoCUT = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_CUT.udl;"

    'Abre a conexão
    dbConexaoCUT.Open strConexao
    dbConexaoCUT.CommandTimeout = 1800

End Sub


Public Sub gsAbrirConexaoBDSC()

    Dim strConexao As String
    
    Set dbConexaoSC = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_SC.udl;"

    'Abre a conexão
    dbConexaoSC.Open strConexao
    dbConexaoSC.CommandTimeout = 1800

End Sub


'*************************************************************************************************************************

Public Function gsReplicaPlics(ByVal vsArgumento As Variant) As String

   ' Descrição: Replica os Plics para argumentos tipo Texto

   Dim mnIncremento     As Integer
   Dim msArgumentoFinal As String
   
   If IsNull(vsArgumento) Then
      gsReplicaPlics = ""
      Exit Function
   End If
   
   For mnIncremento = 1 To Len(vsArgumento)
       If Mid(vsArgumento, mnIncremento, 1) = "'" Then
          msArgumentoFinal = msArgumentoFinal & "''"
       Else
          msArgumentoFinal = msArgumentoFinal & Mid(vsArgumento, mnIncremento, 1)
       End If
   Next mnIncremento
   
   gsReplicaPlics = msArgumentoFinal
    
End Function

Public Function gsTrataPath(ByVal vsPath As String) As String

   'Rotina de tratamento de path
   If Right(vsPath, 1) = "\" Then
      gsTrataPath = vsPath
   Else
      gsTrataPath = vsPath & "\"
   End If

End Function


Function FormatarDataSQL(strData)
    FormatarDataSQL = Mid(strData, 7, 4) + Mid(strData, 4, 2) + Mid(strData, 1, 2)
End Function

