VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExpArq_H 
   Caption         =   "Geração do Arquivo para Ferramenta H"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmArq_H.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   5775
      Begin MSComctlLib.ProgressBar prbExp 
         Height          =   360
         Left            =   105
         TabIndex        =   6
         Top             =   195
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivo"
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5775
      Begin VB.CommandButton cmdProcura 
         Caption         =   "Procurar"
         Height          =   345
         Left            =   4410
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txtDiretorio 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4545
      TabIndex        =   1
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1635
      Width           =   1215
   End
End
Attribute VB_Name = "frmExpArq_H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Dim strArq As String
Dim strDir As String

    strArq = txtDiretorio.Text

    strDir = Mid(strArq, 1, InStrRev(strArq, "\"))

    If Dir(strDir, vbDirectory) = "" Then
        MsgBox "Diretório não existe.", vbCritical + vbOKOnly, "Geração do Arquivo para Ferramenta H"
        Exit Sub
    End If

    If UCase(Right(strArq, 4)) <> ".XLS" Then
        MsgBox "Arquivo inválido.", vbCritical + vbOKOnly, "Geração do Arquivo para Ferramenta H"
        Exit Sub
    End If

    cmdOK.Enabled = False
    cmdProcura.Enabled = False
    cmdCancela.Enabled = False
    
    If ExportarArq(strArq) Then
        MsgBox "Exportação gerada com sucesso.", vbInformation + vbOKOnly, "Geração do Arquivo para Ferramenta H"
    Else
        MsgBox "Erro na geração do arquivo.", vbCritical + vbOKOnly, "Geração do Arquivo para Ferramenta H"
    End If
    
    cmdOK.Enabled = True
    cmdProcura.Enabled = True
    cmdCancela.Enabled = True
    prbExp.Value = 1
    
End Sub

Private Sub cmdProcura_Click()
Dim strDirInicial As String

On Error GoTo ErrcmdProcura

    strDirInicial = txtDiretorio.Text
    
    CommonDialog1.DialogTitle = "Arquivo Exportação F.H"
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Arquivos do MS-Excel (*.xls)|*.xls"
    CommonDialog1.CancelError = True
    CommonDialog1.ShowSave
    txtDiretorio.Text = CommonDialog1.FileName

Exit Sub

ErrcmdProcura:

txtDiretorio.Text = strDirInicial

End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtDiretorio = "C:\Planilha.xls"
End Sub

Private Function ExportarArq(strArquivo As String) As Boolean

Dim objExcel          As Excel.Application
Dim wrkSheetDados     As Worksheet
Dim wrkBook           As Workbook
Dim rs                As New ADODB.Recordset
Dim strSql            As String
Dim intLinha          As Integer

On Error GoTo ErrExportarArq

        ExportarArq = False
        
        Screen.MousePointer = vbHourglass
        
'        'Abre conexao com o Banco SQL-Server
'        Call gsAbrirConexao
    
        strSql = "Select TK.PROJ_ID, "
        strSql = strSql & "TK.TASK_UID, "
        strSql = strSql & "Substring(ltrim(rtrim(TK.TASK_NAME)),1,6) as COD_DESENV, "
        strSql = strSql & "TK.TASK_NAME, "
        strSql = strSql & "TK.TASK_START_DATE, "
        strSql = strSql & "TK.TASK_FINISH_DATE, "
        strSql = strSql & "UPPER(DD.TEXT_VALUE) as TIPO, "
        strSql = strSql & "DC.TEXT_VALUE as Causa, "
        strSql = strSql & "DA.DATE_VALUE as Atraso, "
        strSql = strSql & "DE.TEXT_VALUE as Equipe "
        
        strSql = strSql & "From MSP_TASKS TK INNER JOIN DadosDesenvolvimento DD on "
        strSql = strSql & "TK.PROJ_ID = DD.PROJ_ID And TK.TASK_UID = DD.TEXT_REF_UID "
        strSql = strSql & "left outer join DadosCausa DC on "
        strSql = strSql & "TK.PROJ_ID = DC.PROJ_ID And TK.TASK_UID = DC.TEXT_REF_UID "
        strSql = strSql & "left outer join DadosDtAtrazo DA on "
        strSql = strSql & "TK.PROJ_ID = DA.PROJ_ID And TK.TASK_UID = DA.DATE_REF_UID "
        strSql = strSql & "left outer join DadosEquipe DE on "
        strSql = strSql & "TK.PROJ_ID = DE.PROJ_ID And TK.TASK_UID = DE.TEXT_REF_UID "
        
        strSql = strSql & "Where (UPPER(LTrim(RTrim(DD.TEXT_VALUE))) = 'E' or UPPER(LTrim(RTrim(DD.TEXT_VALUE))) = 'D') "

        rs.Open strSql, dbConexaoSMR, adOpenStatic, adLockReadOnly, adCmdText
        
        prbExp.Max = rs.RecordCount + 2
        prbExp.Min = 1
    
        If Not (objExcel Is Nothing) Then

            wrkBook.Close
            objExcel.Quit

            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing

            'Criando uma nova instância para o objeto
            Set objExcel = CreateObject("Excel.Application")
            'Adicionando um WorkBook
            Set wrkBook = objExcel.Workbooks.Add
            'Setando um WorkSheet
            Set wrkSheetDados = wrkBook.Sheets(1)

        Else

            'Criando uma nova instância para o objeto
            Set objExcel = CreateObject("Excel.Application")
            'Adicionando um WorkBook
            Set wrkBook = objExcel.Workbooks.Add
            'Setando um WorkSheet
            Set wrkSheetDados = wrkBook.Sheets(1)

        End If

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
     
        'Dando um nome para a Sheet (Caption)
        wrkSheetDados.Name = "Dados"
        
        wrkSheetDados.Cells(1, 1).Value = "CodProjeto"
        wrkSheetDados.Cells(1, 2).Value = "UID"
        wrkSheetDados.Cells(1, 3).Value = "CodDes"
        wrkSheetDados.Cells(1, 4).Value = "Descricao"
        wrkSheetDados.Cells(1, 5).Value = "Data Inicio"
        wrkSheetDados.Cells(1, 6).Value = "Data Fim"
        wrkSheetDados.Cells(1, 7).Value = "Fim Previsto"
        wrkSheetDados.Cells(1, 8).Value = "Causa"
        wrkSheetDados.Cells(1, 9).Value = "Tipo"
        wrkSheetDados.Cells(1, 10).Value = "Equipe"

        intLinha = 2

        prbExp.Value = intLinha

        Do While Not rs.EOF
            'Montando o cabeçalho da planilha
            wrkSheetDados.Cells(intLinha, 1).Value = rs("PROJ_ID")
            wrkSheetDados.Cells(intLinha, 2).Value = rs("TASK_UID")
            wrkSheetDados.Cells(intLinha, 3).Value = rs("COD_DESENV")
            wrkSheetDados.Cells(intLinha, 4).Value = rs("TASK_NAME")
            wrkSheetDados.Cells(intLinha, 5).Value = rs("TASK_START_DATE") 'Format(Day(rs("TASK_START_DATE")), "00") & "/" & Format(Month(rs("TASK_START_DATE")), "00") & "/" & Format(Year(rs("TASK_START_DATE")), "0000")
            wrkSheetDados.Cells(intLinha, 6).Value = rs("TASK_FINISH_DATE") 'Format(Day(rs("TASK_FINISH_DATE")), "00") & "/" & Format(Month(rs("TASK_FINISH_DATE")), "00") & "/" & Format(Year(rs("TASK_FINISH_DATE")), "0000")
'            If Not IsNull(rs("Atraso")) Then
                wrkSheetDados.Cells(intLinha, 7).Value = rs("Atraso") 'Format(Day(rs("Atraso")), "00") & "/" & Format(Month(rs("Atraso")), "00") & "/" & Format(Year(rs("Atraso")), "0000")
'            Else
'                wrkSheetDados.Cells(intLinha, 7).Value = ""
'            End If
            wrkSheetDados.Cells(intLinha, 8).Value = rs("Causa")
            wrkSheetDados.Cells(intLinha, 9).Value = rs("TIPO")

            wrkSheetDados.Cells(intLinha, 10).Value = rs("Equipe")

            intLinha = intLinha + 1

            prbExp.Value = intLinha

            rs.MoveNext
        Loop

        'Deletando a planilha caso ela exista
        If Dir$(strArquivo) <> "" Then
            Kill strArquivo
        End If

        'Salvando a planilha criada
        wrkBook.SaveAs strArquivo

        Screen.MousePointer = vbNormal

        'Deixando o Excel visível
        objExcel.Visible = True

        wrkBook.Close
        objExcel.Quit
        
        Set wrkSheetDados = Nothing
        Set wrkBook = Nothing
        Set objExcel = Nothing
    
        Screen.MousePointer = vbNormal

        ExportarArq = True
        
        Exit Function
        
ErrExportarArq:
        
        Screen.MousePointer = vbNormal
        
        If Not (objExcel Is Nothing) Then
        
            objExcel.Visible = True
    
            wrkBook.Close
            objExcel.Quit
            
            Set wrkSheetDados = Nothing
            Set wrkBook = Nothing
            Set objExcel = Nothing
        
        End If
        
    ExportarArq = False
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    frmPrincipal.StatusBar.Panels(1).Text = ""
    
End Sub


