VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAtualizarPercDt_TesteInt_Sin 
   Caption         =   "Atualização dos % e Datas Teste Integrado -> Sinergia"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmAtualizarPercDt_TesteInt_Sin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAtualizacaoTar 
      Height          =   750
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   3
         Top             =   270
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   1170
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   1185
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtualizarPercDt_TesteInt_Sin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    cmdCancela.Enabled = False
    
    fraAtualizacaoTar.Caption = ""
    
    If Atualizar_PercDt_TesteInt_Sin() Then
        MsgBox "Atualização dos % e Datas Teste Integrado -> Sinergia gerada com sucesso.", vbInformation + vbOKOnly, "Atualização dos % e Datas Teste Integrado -> Sinergia"
    Else
        MsgBox "Erro na Atualização dos % e Datas Teste Integrado -> Sinergia.", vbCritical + vbOKOnly, "Atualização dos % e Datas Teste Integrado -> Sinergia"
    End If
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 0
    fraAtualizacaoTar.Caption = ""
    
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Function Atualizar_PercDt_TesteInt_Sin() As Boolean

Dim objExcel          As Excel.Application
Dim wrkSheetDados     As Worksheet
Dim wrkBook           As Workbook

Dim rs1               As New ADODB.Recordset
Dim rs2               As New ADODB.Recordset
Dim rs3               As New ADODB.Recordset
Dim cmdAtualizar      As New ADODB.Command

Dim strSQL            As String
Dim intLinha          As Long
Dim strArqSaida       As String

On Error GoTo ErrAtualizar_PercDt_TesteInt_Sin

        fraAtualizacaoTar.Caption = "Atualização dos % e Datas Teste Integrado -> Sinergia"
        
        Atualizar_PercDt_TesteInt_Sin = False
        
        Screen.MousePointer = vbHourglass
        
        strArqSaida = "C:\Inc_APD_TIN_SIN_" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
        If Not (objExcel Is Nothing) Then

            wrkBook.Close
            objExcel.Quit
            
            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing
        
        End If
        
        'Criando uma nova instância para o objeto
        Set objExcel = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook = objExcel.Workbooks.Add
        'Setando um WorkSheet
        Set wrkSheetDados = wrkBook.Sheets(1)
        
        'Verificando se conseguiu inicializar o objeto
        If objExcel Is Nothing Then
            Screen.MousePointer = vbNormal
            MsgBox "O aplicativo Excel não pode ser inicializado." & vbCrLf, _
                   vbCritical + vbOKOnly, _
                   "ERRO"
            End
        End If
        
        'Deixando o Excel invisível
        objExcel.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel.DisplayAlerts = False
        
        wrkSheetDados.Cells(1, 1).Value = "Projeto Origem TIN"
        wrkSheetDados.Cells(1, 2).Value = "UID Origem TIN"
        wrkSheetDados.Cells(1, 3).Value = "Tarefa Origem TIN"
        wrkSheetDados.Cells(1, 4).Value = "%Complete Origem TIN"
        wrkSheetDados.Cells(1, 5).Value = "Projeto Destino SIN"
        wrkSheetDados.Cells(1, 6).Value = "UID Destino SIN"
        wrkSheetDados.Cells(1, 7).Value = "Tarefa Destino SIN"
        wrkSheetDados.Cells(1, 8).Value = "%Complete Destino SIN"
        
        intLinha = 2
        
        strSQL = "SP_LISTAR_ATUALIZA_PERCDT_TESTEINT_SIN "
                    
        rs1.CursorLocation = adUseClient
        rs2.CursorLocation = adUseClient
        rs3.CursorLocation = adUseClient
        
        rs1.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText
        
        prbExp.Max = rs1.RecordCount
        prbExp.Min = 0
        prbExp.Value = 0
        
        Do While Not rs1.EOF
        
            strSQL = "SP_LISTAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Orig"))) & ", " & Trim(Str(rs1("Cod_UID_Orig")))
            
            rs2.Open strSQL, dbConexaoTIN, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs2.EOF Then
                
                strSQL = "SP_LISTAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Dest"))) & ", " & Trim(Str(rs1("Cod_UID_Dest")))
                
                rs3.Open strSQL, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not rs3.EOF Then
                        
                    If rs2("TASK_PCT_COMP") >= rs3("TASK_PCT_COMP") Then
                    
                        With cmdAtualizar
                        
                            .ActiveConnection = dbConexaoSMR
                            .CommandType = 4
                            .CommandTimeout = 1200
                            .CommandText = "SP_ATUALIZAR_PERCDT_TESTEINT_SIN"
                            
                            .Parameters.Refresh
                            .Parameters(1).Value = rs3("PROJ_ID")
                            .Parameters(2).Value = rs3("TASK_UID")
                            .Parameters(3).Value = rs2("TASK_FINISH_DATE")
                            .Parameters(4).Value = rs2("TASK_PCT_COMP")
                        
                        End With
                        
                        cmdAtualizar.Execute
                        
                    Else
                        
                        With cmdAtualizar
                        
                            .ActiveConnection = dbConexaoSMR
                            .CommandType = 4
                            .CommandTimeout = 1200
                            .CommandText = "SP_ATUALIZAR_PERCDT_TESTEINT_SIN"
                            
                            .Parameters.Refresh
                            .Parameters(1).Value = rs3("PROJ_ID")
                            .Parameters(2).Value = rs3("TASK_UID")
                            .Parameters(3).Value = rs2("TASK_FINISH_DATE")
                            .Parameters(4).Value = Null
                        
                        End With
                        
                        cmdAtualizar.Execute
                    
                    
                        wrkSheetDados.Cells(intLinha, 1).Value = rs2("PROJ_NAME")
                        wrkSheetDados.Cells(intLinha, 2).Value = rs2("TASK_UID")
                        wrkSheetDados.Cells(intLinha, 3).Value = rs2("TASK_NAME")
                        wrkSheetDados.Cells(intLinha, 4).Value = rs2("TASK_PCT_COMP")
                        wrkSheetDados.Cells(intLinha, 5).Value = rs3("PROJ_NAME")
                        wrkSheetDados.Cells(intLinha, 6).Value = rs3("TASK_UID")
                        wrkSheetDados.Cells(intLinha, 7).Value = rs3("TASK_NAME")
                        wrkSheetDados.Cells(intLinha, 8).Value = rs3("TASK_PCT_COMP")
                        intLinha = intLinha + 1
                    
                    End If
                End If
            End If
            
            rs2.Close
            rs3.Close
            
            prbExp.Value = prbExp.Value + 1
            
            rs1.MoveNext
        
        Loop
        
        rs1.Close
        
        
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida) <> "" Then
            Kill strArqSaida
        End If

        'Salvando a planilha criada
        wrkBook.SaveAs strArqSaida
        
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
        
        Screen.MousePointer = vbNormal

        Atualizar_PercDt_TesteInt_Sin = True
        
        Exit Function
        
ErrAtualizar_PercDt_TesteInt_Sin:
   
   If Not (objExcel Is Nothing) Then
    
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
    End If
    
    Screen.MousePointer = vbNormal
    Atualizar_PercDt_TesteInt_Sin = False
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub

