VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAtualizarPerc_SC_SMR 
   Caption         =   "Atualização do Percentual EP -> SMR"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmAtualizarPerc_SC_SMR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAtualizacaoTar 
      Height          =   660
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   3
         Top             =   195
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
      Top             =   855
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3225
      TabIndex        =   0
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtualizarPerc_SC_SMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    cmdCancela.Enabled = False
    
    If AtualizarPerc_SC_SMR() Then
        MsgBox "Atualização do Percentual SC -> SMR gerada com sucesso.", vbInformation + vbOKOnly, "Atualização do Percentual SC -> SMR"
    Else
        MsgBox "Erro na Atualização do Percentual SC -> SMR.", vbCritical + vbOKOnly, "Atualização do Percentual SC -> SMR"
    End If
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 0
    
End Sub


Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


Private Function AtualizarPerc_SC_SMR() As Boolean

Dim objExcel          As Excel.Application
Dim wrkSheetDados     As Worksheet
Dim wrkBook           As Workbook

Dim objExcel1         As Excel.Application
Dim wrkSheetDados1    As Worksheet
Dim wrkBook1          As Workbook

Dim rs1               As New ADODB.Recordset
Dim rs2               As New ADODB.Recordset
Dim rs3               As New ADODB.Recordset

Dim strSql            As String
Dim intLinha          As Long
Dim strArqSaida       As String
Dim intLinha1         As Long
Dim strArqSaida1      As String

On Error GoTo ErrAtualizarPerc_SC_SMR

        fraAtualizacaoTar.Caption = "Atualização de Tarefas SC -> PMO"
        fraAtualizacaoTar.Refresh
        
        AtualizarPerc_SC_SMR = False
        
        Screen.MousePointer = vbHourglass
        
        strArqSaida = "C:\INC_AT_SC_PMO" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
        strArqSaida1 = "C:\INC_TS_SC_PMO" & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00")
        
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
        
        If Not (objExcel1 Is Nothing) Then

            wrkBook1.Close
            objExcel1.Quit
            
            Set wrkSheetDados1 = Nothing
            Set wrkBook1 = Nothing
            Set objExcel1 = Nothing
        
        End If
        
        'Criando uma nova instância para o objeto
        Set objExcel1 = CreateObject("Excel.Application")
        'Adicionando um WorkBook
        Set wrkBook1 = objExcel1.Workbooks.Add
        'Setando um WorkSheet
        Set wrkSheetDados1 = wrkBook1.Sheets(1)
        
        'Deixando o Excel invisível
        objExcel.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel.DisplayAlerts = False
        
        'Cria titulos das colunas da tabela INC_AT_SC_PMO
        wrkSheetDados.Cells(1, 1).Value = "Projeto Orig. SC."
        wrkSheetDados.Cells(1, 2).Value = "UID Orig. SC."
        wrkSheetDados.Cells(1, 3).Value = "Tarefa Orig. SC."
        wrkSheetDados.Cells(1, 4).Value = "%Complete Orig. SC."
        wrkSheetDados.Cells(1, 5).Value = "Projeto Dest. PMO."
        wrkSheetDados.Cells(1, 6).Value = "UID Dest. PMO."
        wrkSheetDados.Cells(1, 7).Value = "Tarefa Dest. PMO."
        wrkSheetDados.Cells(1, 8).Value = "%Complete Dest. PMO."
        
        'Deixando o Excel invisível
        objExcel1.Visible = False
        
        'Vai inibir algumas mensagens enviadas pelo Excel
        objExcel1.DisplayAlerts = False
        
        'Cria titulos das colunas da tabela INC_TS_PMO_SC
        wrkSheetDados1.Cells(1, 1).Value = "Projeto Orig. SC."
        wrkSheetDados1.Cells(1, 2).Value = "UID Orig. SC."
        wrkSheetDados1.Cells(1, 3).Value = "Tarefa Orig. SC."
        wrkSheetDados1.Cells(1, 5).Value = "Projeto Dest. PMO."
        wrkSheetDados1.Cells(1, 6).Value = "UID Dest. PMO."
        wrkSheetDados1.Cells(1, 7).Value = "Tarefa Dest. PMO."
        
        intLinha = 2
        intLinha1 = 2
        
        'Seleciona dados da tabela De/Para PMO->SC
        strSql = "SP_LISTAR_ATUALIZA_PERC_SC_SMR "
                    
        rs1.CursorLocation = adUseClient
        rs2.CursorLocation = adUseClient
        rs3.CursorLocation = adUseClient
        
        rs1.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        prbExp.Max = rs1.RecordCount
        prbExp.Min = 0
        prbExp.Value = 0
        
        Do While Not rs1.EOF
        
            'Seleciona dados da tarefa de origem
            strSql = "SP_LISTAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Orig"))) & ", " & Trim(Str(rs1("Cod_UID_Orig")))
            
            rs2.Open strSql, dbConexaoSC, adOpenStatic, adLockReadOnly, adCmdText
            
            If Not rs2.EOF Then
                
                'Seleciona dados da tarefa de destino
                strSql = "SP_LISTAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Dest"))) & ", " & Trim(Str(rs1("Cod_UID_Dest")))
                
                rs3.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not rs3.EOF Then
                    
                    'Se tarefa origem for atividade e tarefa destino for sumaria não atualizar e salvar na planilha INC_TS_PMO_SC
                    If rs2("TASK_IS_SUMMARY") = False And rs3("TASK_IS_SUMMARY") = True Then

                        wrkSheetDados1.Cells(intLinha1, 1).Value = rs2("PROJ_NAME")
                        wrkSheetDados1.Cells(intLinha1, 2).Value = rs2("TASK_UID")
                        wrkSheetDados1.Cells(intLinha1, 3).Value = rs2("TASK_NAME")
                        wrkSheetDados1.Cells(intLinha1, 5).Value = rs3("PROJ_NAME")
                        wrkSheetDados1.Cells(intLinha1, 6).Value = rs3("TASK_UID")
                        wrkSheetDados1.Cells(intLinha1, 7).Value = rs3("TASK_NAME")

                        intLinha1 = intLinha1 + 1

                    Else
                        
                        If rs2("TASK_PCT_COMP") >= rs3("TASK_PCT_COMP") Then
                            'Se o % de origem for maior ou igual a % de destino atualiza na base
                            
                            strSql = "SP_ATUALIZAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Dest"))) & ", " & Trim(Str(rs3("TASK_ID"))) & ", " & Trim(Str(rs1("Cod_UID_Dest"))) & ", " & Trim(Str(rs2("TASK_PCT_COMP")))
                            
                            dbConexaoSMR.Execute strSql
                        
                        Else
                            'Se o % de origem for menor a % de destino salvar na planilha INC_AT_PMO_SC
                            
                            strSql = "SP_ATUALIZAR_TAREFAS " & Trim(Str(rs1("Cod_Proj_Dest"))) & ", " & Trim(Str(rs3("TASK_ID"))) & ", " & Trim(Str(rs1("Cod_UID_Dest"))) & ", " & Trim(Str(rs2("TASK_PCT_COMP")))
                                                        
                            dbConexaoSMR.Execute strSql
                            
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
                    
                    
        'Deletando a planilha caso ela exista
        If Dir$(strArqSaida1) <> "" Then
            Kill strArqSaida1
        End If

        'Salvando a planilha criada
        wrkBook1.SaveAs strArqSaida1
        
        'Deixando o Excel visível
        objExcel1.Visible = True

        wrkBook1.Close
        objExcel1.Quit
        
        Set wrkSheetDados1 = Nothing
        Set wrkBook1 = Nothing
        Set objExcel1 = Nothing
                    
        Screen.MousePointer = vbNormal

        AtualizarPerc_SC_SMR = True
        
        Exit Function
        
ErrAtualizarPerc_SC_SMR:
   
   If Not (objExcel Is Nothing) Then
    
        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
    End If
    
   If Not (objExcel1 Is Nothing) Then
    
        'Deixando o Excel visível
        objExcel1.Visible = True

        wrkBook1.Close
        objExcel1.Quit
        
        Set wrkSheetDados1 = Nothing
        Set wrkBook1 = Nothing
        Set objExcel1 = Nothing
    
    End If
    
    Screen.MousePointer = vbNormal
    AtualizarPerc_SC_SMR = False
    
End Function

