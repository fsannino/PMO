Attribute VB_Name = "mdlGobal"
Option Explicit

Public dbConexaoLight  As ADODB.Connection


Public Sub gsFecharConexaoBDLight()

    If Not (dbConexaoLight Is Nothing) Then
        If dbConexaoLight.State = adStateOpen Then
            dbConexaoLight.Close
            Set dbConexaoLight = Nothing
        End If
    End If

End Sub


Public Sub gsAbrirConexaoBDLight()

    Dim strConexao As String
    
    Set dbConexaoLight = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_Light.udl;"

    'Abre a conexão
    dbConexaoLight.Open strConexao
    dbConexaoLight.CommandTimeout = 1800

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

