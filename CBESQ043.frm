VERSION 5.00
Begin VB.Form frmCAIFFecharRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Fechar Recibos de Mensalidades"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
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
   ScaleHeight     =   2355
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   900
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2002"
      Top             =   1380
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2003"
      Top             =   1380
      Width           =   1200
   End
   Begin VB.TextBox txtAtencao 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1200
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "CBESQ043.frx":0000
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "frmCAIFFecharRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSFecharRecibos As Workspace
Dim mBDFecharRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSFecharRecibos = DBEngine.CreateWorkspace("CAIFAlteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDFecharRecibos = mWSFecharRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Fechar Recibos-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFechar_Click()
    Dim mProcessamento As Processamento
    Dim qryFecharRecibos As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgBox = MsgBox("Confirma que quer fechar os recibos ?", vbQuestion + vbYesNo, _
                        "Fechar Recibos")
    ' se resposta não sai
    If iRespMsgBox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSFecharRecibos.BeginTrans
    
    Set qryFecharRecibos = mBDFecharRecibos.QueryDefs("CAIF RECIBOS Altera Estado")
    
    qryFecharRecibos.Parameters("Estado Novo") = "D"
    qryFecharRecibos.Parameters("Estado Velho") = "P"
    qryFecharRecibos.Parameters("Utiliz") = gUtilizador.Nome
    
    ' executa a inserção
    qryFecharRecibos.Execute dbFailOnError

    Set qryFecharRecibos = mBDFecharRecibos.QueryDefs("CAIF RECIBOS Fechar")
    
    qryFecharRecibos.Parameters("Utiliz") = gUtilizador.Nome
    
    ' executa a inserção
    qryFecharRecibos.Execute dbFailOnError
    mWSFecharRecibos.CommitTrans
    
    MsgBox "Criou definitivamente os recibos.", vbInformation + vbOKOnly, "Fechar Recibos"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSFecharRecibos.Rollback
    Call ErrosGerais("Fechar Recibos", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Call VerificaFecho
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Public Sub VerificaFecho()
    Dim recRECIBOS As Recordset
    
    cSql = "SELECT * FROM RECIBOS_IDOSOS WHERE ISNULL(ESTADO_REC)"
    
    Set recRECIBOS = mBDFecharRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)

    ' se existirem recibos para fechar não deixa criar mais
    If (recRECIBOS.EOF And recRECIBOS.BOF) Then
        cmdFechar.Visible = False
    End If
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
                    " Vai Fechar os recibos das Mensalidades." & vbCrLf & _
                    " Esta opção cria definitivamente os recibos," & vbCrLf & _
                    " pondo-os  disponíveis  para  pagamento."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSFecharRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSFecharRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFFecharRecibos = Nothing
End Sub


