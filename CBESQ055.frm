VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmNomeEducadora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nome Educadora"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
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
   ScaleHeight     =   4725
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "2002"
      Top             =   3705
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   " Educadora "
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
      Height          =   930
      Left            =   150
      TabIndex        =   5
      Top             =   2160
      Width           =   5610
      Begin VB.TextBox txtNomeEducadora 
         Height          =   360
         Left            =   180
         MaxLength       =   50
         TabIndex        =   6
         Top             =   345
         Width           =   5250
      End
   End
   Begin VB.Frame fraSalas 
      Caption         =   " Sala "
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
      Height          =   930
      Left            =   150
      TabIndex        =   2
      Top             =   1155
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboSalas 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   390
         Width           =   5265
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   9287
         _ExtentY        =   582
         _StockProps     =   93
         Text            =   "<Todas as Salas>"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame fraEquipamento 
      Caption         =   " Equipamento "
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
      Height          =   930
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   390
         Width           =   5265
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   9287
         _ExtentY        =   582
         _StockProps     =   93
         Text            =   "<Todas as Instituições>"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "2003"
      Top             =   3705
      Width           =   1200
   End
End
Attribute VB_Name = "frmNomeEducadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSNomeEducadora As Workspace
Dim mBDNomeEducadora As Database
Dim mBDNomeEducadoraTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSNomeEducadora = DBEngine.CreateWorkspace("NomeEducadora", gUtilizador.Nome, gUtilizador.Password)
    Set mBDNomeEducadora = mWSNomeEducadora.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDNomeEducadoraTemp = mWSNomeEducadora.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Declarações de Frequência-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    cboSalas.Text = "<Todas as Salas>"
    txtNomeEducadora.Text = vbNullString
End Sub

Private Sub cboInstituicao_DropDown()
    ' carrega a combo
    Call CarregacboInstituicao(cboInstituicao)
End Sub

Private Sub cboInstituicao_InitColumnProps()
    With cboInstituicao
        .StyleSets.Add "Cabecalho"
        .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
        .StyleSets("Cabecalho").ForeColor = vbTitleBarText
        .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
        .StyleSets("Cabecalho").Font.Size = 10
        .StyleSets("Cabecalho").Font.Bold = True
        
        .AllowInput = False
        .BackColorOdd = dCorAmarelo
        .ForeColorEven = &H0&
        .FieldSeparator = vbTab
        .DataFieldList = "Column 0"
        .HeadStyleSet = "Cabecalho"
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10

                    
        ' coluna 0
        .Columns.Add 0
        .Columns(0).Caption = "Nome da Institução"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
End Sub

Private Sub cboSalas_Click()
    txtNomeEducadora.Text = cDescodificaNomeEducadora(cCodificaInstituicao(cboInstituicao.Text), cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text))
End Sub

Private Sub cboSalas_DropDown()
    ' carrega a combo
    Call CarregacboSalas(cboSalas, cboInstituicao.Text)
End Sub


Private Sub cboSalas_InitColumnProps()
    With cboSalas
        .StyleSets.Add "Cabecalho"
        .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
        .StyleSets("Cabecalho").ForeColor = vbTitleBarText
        .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
        .StyleSets("Cabecalho").Font.Size = 10
        .StyleSets("Cabecalho").Font.Bold = True
        
        .AllowInput = False
        .BackColorOdd = dCorAmarelo
        .ForeColorEven = &H0&
        .FieldSeparator = vbTab
        .DataFieldList = "Column 0"
        .HeadStyleSet = "Cabecalho"
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10
                    
        ' coluna 0
        .Columns.Add 0
        .Columns(0).Caption = "Nome da Sala"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
        
        ' coluna 2
        .Columns.Add 2
        .Columns(2).Visible = False
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub





Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryNomeEducadora As QueryDef
    
    Set mProcessamento = New Processamento
    
On Error GoTo TrataErro
    ' começa a transação
    mWSNomeEducadora.BeginTrans
    Set qryNomeEducadora = mBDNomeEducadora.QueryDefs("TABSALAS Altera NOME_EDUCADORA")
        
    ' parametros de input
    qryNomeEducadora.Parameters("COD_INST") = cCodificaInstituicao(cboInstituicao.Text)
    qryNomeEducadora.Parameters("COD_SALA") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)
    qryNomeEducadora.Parameters("NOME_EDUCADORA") = txtNomeEducadora.Text
    qryNomeEducadora.Parameters("Utiliz") = gUtilizador.Nome
    ' executa a query
    qryNomeEducadora.Execute dbFailOnError
    
    mWSNomeEducadora.CommitTrans
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSNomeEducadora.Rollback
    Call ErrosGerais("OK" & " Nome Educadora", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
'    Unload Me
End Sub





Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSNomeEducadora.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSNomeEducadora = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmNomeEducadora = Nothing
End Sub



