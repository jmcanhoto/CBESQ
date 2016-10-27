VERSION 5.00
Begin VB.Form frmMultaRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multa de Atraso nos Pagamentos das Mensalidades"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPercentagem 
      Caption         =   " Percentagem "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   960
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3150
      Begin VB.OptionButton optPercentagem 
         Caption         =   "20%"
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   945
      End
      Begin VB.OptionButton optPercentagem 
         Caption         =   "5%"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.OptionButton optPercentagem 
         Caption         =   "10%"
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdMultar 
      Caption         =   "&Multar"
      Height          =   900
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2002"
      Top             =   1500
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2003"
      Top             =   1500
      Width           =   1200
   End
   Begin VB.TextBox txtAtencao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1200
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "CBESQ051.frx":0000
      Top             =   120
      Width           =   5880
   End
End
Attribute VB_Name = "frmMultaRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSMultarRecibos As Workspace
Dim mBDMultarRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSMultarRecibos = DBEngine.CreateWorkspace("Alteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDMultarRecibos = mWSMultarRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Multar Recibos-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdMultar_Click()
    Dim mProcessamento As Processamento
    Dim qryMultarRecibos As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgbox = MsgBox("Confirma que quer Multar os recibos ?", vbQuestion + vbYesNo, _
                        "Multar Recibos")
    ' se resposta não sai
    If iRespMsgbox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSMultarRecibos.BeginTrans

    ' Percentagem 5%
    If optPercentagem(0).Value Then
        ' se for Julho, só multa CATL
        If Month(Date) = "7" Then
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 05 com Parametro")
            qryMultarRecibos.Parameters("COD_INST") = "001"
            qryMultarRecibos.Parameters("COD_SALA") = "008"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
        
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
                
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 05 com Parametro")
            qryMultarRecibos.Parameters("COD_INST") = "002"
            qryMultarRecibos.Parameters("COD_SALA") = "006"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
                
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        ' se for Agosto , não multa nada
        ElseIf Month(Date) = "8" Then
            ' NÃO MULTA NADA
        Else
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 05")
        
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
            
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        End If
    ' Percentagem 10%
    ElseIf optPercentagem(1).Value Then
        ' se for Julho, só multa CATL
        If Month(Date) = "7" Then
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 10 com Parametro")
            qryMultarRecibos.Parameters("COD_INST") = "001"
            qryMultarRecibos.Parameters("COD_SALA") = "008"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
        
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
                
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 10 com Parametro")
            qryMultarRecibos.Parameters("COD_INST") = "002"
            qryMultarRecibos.Parameters("COD_SALA") = "006"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
                
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        ' se for Agosto , não multa nada
        ElseIf Month(Date) = "8" Then
            ' NÃO MULTA NADA
        Else
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 10")
        
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
            
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        End If
    ' Percentagem 20%
    ElseIf optPercentagem(2).Value Then
        ' se for Agosto, só multa CATL
        If Month(Date) = "8" Then
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 20 com Parametro")
            qryMultarRecibos.Parameters("ANO") = Year(Date)
            qryMultarRecibos.Parameters("COD_MES") = "07"
            qryMultarRecibos.Parameters("COD_INST") = "001"
            qryMultarRecibos.Parameters("COD_SALA") = "008"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
        
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
                
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 20 com Parametro")
            qryMultarRecibos.Parameters("ANO") = Year(Date)
            qryMultarRecibos.Parameters("COD_MES") = "07"
            qryMultarRecibos.Parameters("COD_INST") = "002"
            qryMultarRecibos.Parameters("COD_SALA") = "006"
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
                
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        ' se for Setembro , não multa nada
        ElseIf Month(Date) = "9" Then
            ' NÃO MULTA NADA
        Else
            Set qryMultarRecibos = mBDMultarRecibos.QueryDefs("RECIBOS Multar 20")
            
            qryMultarRecibos.Parameters("ANO") = Year(Date)
            If Month(Date) = "1" Then
                qryMultarRecibos.Parameters("ANO") = Year(Date) - 1
                qryMultarRecibos.Parameters("COD_MES") = "12"
            Else
                qryMultarRecibos.Parameters("COD_MES") = Format(Month(Date) - 1, "00")
            End If
            qryMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
            
            ' executa a inserção
            qryMultarRecibos.Execute dbFailOnError
        End If
    End If
    
    mWSMultarRecibos.CommitTrans
    
    MsgBox "Multou os Recibos.", vbInformation + vbOKOnly, "Multar Recibos"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSMultarRecibos.Rollback
    Call ErrosGerais("Multar Recibos", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:

End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub



Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    
    txtAtencao.Text = " Atenção" & vbCrLf & _
                    " Vai Multar os recibos das Mensalidades." & vbCrLf & _
                    " Esta opção altera os recibos que ainda estão" & vbCrLf & _
                    " para  pagamento, acrescentando a Multa !"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSMultarRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSMultarRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmMultaRecibos = Nothing
End Sub


