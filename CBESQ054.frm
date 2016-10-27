VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMudarSala 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mudar de Sala"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
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
   ScaleHeight     =   6225
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame Frame2 
      Caption         =   " Para "
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
      Height          =   2385
      Left            =   150
      TabIndex        =   5
      Top             =   2700
      Width           =   5820
      Begin VB.Frame Frame4 
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
         Left            =   105
         TabIndex        =   8
         Top             =   1320
         Width           =   5610
         Begin SSDataWidgets_B.SSDBCombo cboSalaPara 
            Height          =   330
            Left            =   165
            TabIndex        =   9
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
      Begin VB.Frame Frame3 
         Caption         =   " Instituição "
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
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   5610
         Begin SSDataWidgets_B.SSDBCombo cboInstituicaoPara 
            Height          =   330
            Left            =   165
            TabIndex        =   7
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
   End
   Begin VB.Frame Frame1 
      Caption         =   " De "
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
      Height          =   2385
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5820
      Begin VB.Frame fraInstituicao 
         Caption         =   " Instituição "
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
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   5610
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   330
            Left            =   165
            TabIndex        =   2
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
         Left            =   105
         TabIndex        =   3
         Top             =   1320
         Width           =   5610
         Begin SSDataWidgets_B.SSDBCombo cboSalas 
            Height          =   330
            Left            =   165
            TabIndex        =   4
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
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2003"
      Top             =   5175
      Width           =   1200
   End
   Begin VB.CommandButton cmdMudar 
      Caption         =   "&Mudar"
      Height          =   900
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2010"
      Top             =   5175
      Width           =   1200
   End
End
Attribute VB_Name = "frmMudarSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSMudar As Workspace
Dim mBDMudar As Database
Dim mBDMudarTemp As Database

Dim tBDAberta
Dim iRespMsgBox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSMudar = DBEngine.CreateWorkspace("Mudar", gUtilizador.Nome, gUtilizador.Password)
    Set mBDMudar = mWSMudar.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDMudarTemp = mWSMudar.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Mudar de Sala-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    cboSalas.Text = "<Todas as Salas>"
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

Private Sub cboInstituicaoPara_Click()
    cboSalaPara.Text = "<Todas as Salas>"
End Sub

Private Sub cboInstituicaoPara_DropDown()
    ' carrega a combo
    Call CarregacboInstituicao(cboInstituicaoPara)
End Sub


Private Sub cboInstituicaoPara_InitColumnProps()
    With cboInstituicaoPara
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

Private Sub cboSalaPara_DropDown()
    ' carrega a combo
    Call CarregacboSalas(cboSalaPara, cboInstituicaoPara.Text)
End Sub

Private Sub cboSalaPara_InitColumnProps()
    With cboSalaPara
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

Private Sub cmdMudar_Click()
    Dim mProcessamento As Processamento
    Dim qryMudarSala As QueryDef
    Dim cCod_Inst
    Dim cCod_Sala
    Dim cCod_Inst_Nova
    Dim cCod_Sala_Nova
    
    Set mProcessamento = New Processamento
    
    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)

    cCod_Inst_Nova = cCodificaInstituicao(cboInstituicaoPara.Text)
    cCod_Sala_Nova = cCodificaSala(cCod_Inst_Nova, cboSalaPara.Text)
   
     ' pede confirmação se quer continuar
    iRespMsgBox = MsgBox("Confirma que quer Mudar de Sala os Utentes !!!", vbQuestion + vbYesNo, _
                        "Mudar de Sala")
    ' se resposta não sai
    If iRespMsgBox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSMudar.BeginTrans
    
    Set qryMudarSala = mBDMudar.QueryDefs("UTENTES Mudar de Sala")
    
    qryMudarSala.Parameters("COD_INST") = cCod_Inst
    qryMudarSala.Parameters("COD_SALA") = cCod_Sala
    qryMudarSala.Parameters("COD_INST_NOVA") = cCod_Inst_Nova
    qryMudarSala.Parameters("COD_SALA_NOVA") = cCod_Sala_Nova
        
    qryMudarSala.Execute dbFailOnError
    
    mWSMudar.CommitTrans
    
    ' actualizou as Mensalidades
    MsgBox "Mudança de Sala concluída com sucesso !!!", vbInformation + vbOKOnly, "Criar Quotas"
       
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Mudar de Sala", Err.Number, Err.Description)
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSMudar.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSMudar = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmMudarSala = Nothing
End Sub



Private Sub SSDBCombo1_InitColumnProps()

End Sub


