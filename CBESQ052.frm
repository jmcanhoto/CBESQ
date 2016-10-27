VERSION 5.00
Begin VB.Form frmCAIFMultaRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Multa de Atraso nos Pagamentos das Mensalidades"
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
      Text            =   "CBESQ052.frx":0000
      Top             =   120
      Width           =   5880
   End
End
Attribute VB_Name = "frmCAIFMultaRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSCAIFMultarRecibos As Workspace
Dim mBDCAIFMultarRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSCAIFMultarRecibos = DBEngine.CreateWorkspace("Alteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDCAIFMultarRecibos = mWSCAIFMultarRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
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
    Dim qryCAIFMultarRecibos As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgbox = MsgBox("Confirma que quer Multar os recibos ?", vbQuestion + vbYesNo, _
                        "Multar Recibos")
    ' se resposta não sai
    If iRespMsgbox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSCAIFMultarRecibos.BeginTrans
    

    If optPercentagem(0).Value Then
        Set qryCAIFMultarRecibos = mBDCAIFMultarRecibos.QueryDefs("CAIF RECIBOS Multar 05")
    ElseIf optPercentagem(1).Value Then
        Set qryCAIFMultarRecibos = mBDCAIFMultarRecibos.QueryDefs("CAIF RECIBOS Multar 10")
    ElseIf optPercentagem(2).Value Then
        Set qryCAIFMultarRecibos = mBDCAIFMultarRecibos.QueryDefs("CAIF RECIBOS Multar 20")
    End If
    qryCAIFMultarRecibos.Parameters("COD_INST") = "501"
    qryCAIFMultarRecibos.Parameters("COD_SALA") = "001"
    qryCAIFMultarRecibos.Parameters("Utiliz") = gUtilizador.Nome
    
    ' executa a inserção
    qryCAIFMultarRecibos.Execute dbFailOnError
    mWSCAIFMultarRecibos.CommitTrans
    
    MsgBox "Multou os Recibos.", vbInformation + vbOKOnly, "Multar Recibos"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSCAIFMultarRecibos.Rollback
    Call ErrosGerais("Multar Recibos", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Call VerificaFecho
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Public Sub VerificaFecho()
    Dim recRECIBOS As Recordset
    
    cSql = "SELECT * FROM RECIBOS WHERE ISNULL(ESTADO_REC)"
    
    Set recRECIBOS = mBDCAIFMultarRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)

'    ' se existirem recibos para Multar não deixa criar mais
'    If (recRECIBOS.EOF And recRECIBOS.BOF) Then
'        cmdMultar.Visible = False
'    End If
    recRECIBOS.Close
    Set recRECIBOS = Nothing
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    
    Call VerificaFecho
    
    txtAtencao.Text = " Atenção" & vbCrLf & _
                    " Vai Multar os recibos das Mensalidades." & vbCrLf & _
                    " Esta opção altera os recibos que ainda estão" & vbCrLf & _
                    " para  pagamento, acrescentando a Multa !"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSCAIFMultarRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSCAIFMultarRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmMultaRecibos = Nothing
End Sub


