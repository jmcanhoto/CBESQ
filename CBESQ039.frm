VERSION 5.00
Begin VB.Form frmCAIFActualizacaoMensalidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Actualização de Mensalidades"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2003"
      Top             =   1485
      Width           =   1200
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      Height          =   900
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   1485
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
      Text            =   "CBESQ039.frx":0000
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "frmCAIFActualizacaoMensalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSActualizaMensalidade As Workspace
Dim mBDActualizaMensalidade As Database
Dim mBDActualizaMensalidadeTemp As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSActualizaMensalidade = DBEngine.CreateWorkspace("CAIFActualizaMensalidade", gUtilizador.Nome, gUtilizador.Password)
    Set mBDActualizaMensalidade = mWSActualizaMensalidade.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDActualizaMensalidadeTemp = mWSActualizaMensalidade.OpenDatabase(cBDComNomeUtilizador)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Criar Quotas-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdActualizar_Click()
    Dim mProcessamento As Processamento
    Dim qryActualizaMensalidade As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgBox = MsgBox("Confirma que quer actualizar as Mensalidades dos Utentes do CAIF !!!", vbQuestion + vbYesNo, _
                        "Actualizar Mensalidades")
    ' se resposta não sai
    If iRespMsgBox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSActualizaMensalidade.BeginTrans
    
    Set qryActualizaMensalidade = mBDActualizaMensalidade.QueryDefs("CAIF UTENTES Actualizar Mensalidade")
    
    qryActualizaMensalidade.Parameters("Utiliz") = gUtilizador.Nome
        
    qryActualizaMensalidade.Execute dbFailOnError
    
    mWSActualizaMensalidade.CommitTrans
    
    ' actualizou as Mensalidades
    MsgBox "Actualização de Mensalidades concluída com sucesso !!!", vbInformation + vbOKOnly, "Criar Quotas"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSActualizaMensalidade.Rollback
    Call ErrosGerais("CAIF - Actualização de Mensalidades", Err.Number, Err.Description)
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
                    " Vai Actualizar as Mensalidades dos Utentes do CAIF."
    
    
    End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSActualizaMensalidade.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSActualizaMensalidade = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFActualizacaoMensalidades = Nothing
End Sub





