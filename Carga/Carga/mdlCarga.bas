Attribute VB_Name = "mdlCarga"
Option Explicit

'DLL para Get (Pegar) 'em arquivos .Ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public dbConexao As ADODB.Connection

'Public dbConexao As ADODB.Connection
Public strMensagem As String

Type Log_TP
    Campo_1 As String * 1
    Campo_B As String * 1
    Campo_2 As String * 100
    Campo_3 As String * 2  'vbCrLf
End Type

Public Sub AbrirConexao()

Dim strConexao As String
    
    Set dbConexao = New ADODB.Connection

    'Constrói a ConnectionString para abertura da conexão
    strConexao = "File Name=" & App.Path & "\Carga_Light.udl;"

    'Abre a conexão
    dbConexao.Open strConexao

End Sub

Public Function gsReplicaPlics(ByVal vsArgumento As Variant) As String

   ' Descrição: Replica os Plics para argumentos tipo Texto
   ' Autor    : Alessandra Ribeiro Alexander
   

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

'---------------------------------------
'Realiza a leitura do arquivo .INI
'---------------------------------------
 Public Function LerArqIni(Secao, Chave, Arquivo)

    Dim strRetVal As String
    Dim intWorked As Integer

    strRetVal = String$(255, 0)
    intWorked = GetPrivateProfileString(Secao, Chave, "", strRetVal, Len(strRetVal), Arquivo)
    
    If intWorked = 0 Then
        LerArqIni = ""
    Else
        LerArqIni = Left(strRetVal, intWorked)
    End If

End Function


Public Function TrocarCarEsp_Enter(strValor As String) As String

    TrocarCarEsp_Enter = Replace(Replace(strValor, "**", Chr(10) & Chr(13)), "##", Chr(10) & Chr(13))

End Function

