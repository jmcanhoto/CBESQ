VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFDeclaracoesFrequencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Declarações de Frequência"
   ClientHeight    =   5670
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
   ScaleHeight     =   5670
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame fraAnoLectivo 
      Caption         =   " Data "
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
      Height          =   1020
      Left            =   150
      TabIndex        =   7
      Top             =   3585
      Width           =   5610
      Begin VB.TextBox txtAnoLec 
         Height          =   360
         Left            =   165
         MaxLength       =   50
         TabIndex        =   8
         Top             =   390
         Width           =   5280
      End
   End
   Begin VB.Frame fraUtente 
      Caption         =   " Utente "
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
      Height          =   1350
      Left            =   150
      TabIndex        =   4
      Top             =   2175
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboNumUtente 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   390
         Width           =   1350
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   2381
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboNome 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   780
         Width           =   5340
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   9419
         _ExtentY        =   582
         _StockProps     =   93
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
   Begin Crystal.CrystalReport rptDeclaracoesFreq 
      Left            =   2730
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2003"
      Top             =   4710
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2010"
      Top             =   4710
      Width           =   1200
   End
End
Attribute VB_Name = "frmCAIFDeclaracoesFrequencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSDeclaracoesFreq As Workspace
Dim mBDDeclaracoesFreq As Database
Dim mBDDeclaracoesFreqTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSDeclaracoesFreq = DBEngine.CreateWorkspace("DeclaracoesFreq", gUtilizador.Nome, gUtilizador.Password)
    Set mBDDeclaracoesFreq = mWSDeclaracoesFreq.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDDeclaracoesFreqTemp = mWSDeclaracoesFreq.OpenDatabase(cBDComNomeUtilizador)

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
    cboNumUtente.Text = vbNullString
    cboNome.Text = vbNullString
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
    cboNumUtente.Text = vbNullString
    cboNome.Text = vbNullString
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

Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    Dim lNum_Utente
    Dim cAno
    
    Set mProcessamento = New Processamento

    If cboNumUtente.Text = vbNullString Or cboNome.Text = vbNullString Then
        MsgBox "Não indicou o Utente!", vbInformation, "Declarações de Frequência"
        GoTo SairDoProcedimento
    End If
    
    cAno = str$(Year(Date))
    lNum_Utente = CLng(cboNumUtente.Text)
    cNomeMapa = "CBESQ042.RPT"

On Error GoTo TrataErro

    ' apaga os registos da Temp32
    cSql = "DELETE * FROM LISTA_DECLARACOES_FREQ_IDOSOS;"
    ' apaga o registo em Temp32
    mBDDeclaracoesFreqTemp.Execute cSql, dbFailOnError
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO LISTA_DECLARACOES_FREQ_IDOSOS (NUM_UTENTE,COD_INST,COD_SALA,NOME," & _
        "MENSALIDADE) IN '" & cBDComNomeUtilizador & "' " & _
        "SELECT UTENTES_IDOSOS.NUM_UTENTE,UTENTES_IDOSOS.COD_INST,UTENTES_IDOSOS.COD_SALA," & _
        "UTENTES_IDOSOS.NOME,UTENTES_IDOSOS.MENSALIDADE "
    cSql = cSql & "FROM UTENTES_IDOSOS " & _
        "WHERE UTENTES_IDOSOS.NUM_UTENTE=" & lNum_Utente & " AND ISNULL(UTENTES_IDOSOS.DATA_SAIDA);"
    
    ' insere o registo em Temp32
    mBDDeclaracoesFreq.Execute cSql, dbFailOnError
    
    Dim recDeclaracoesFreq As Recordset
    
    Set recDeclaracoesFreq = mBDDeclaracoesFreqTemp.OpenRecordset("LISTA_DECLARACOES_FREQ_IDOSOS", dbOpenDynaset)

    If Not (recDeclaracoesFreq.EOF Or recDeclaracoesFreq.BOF) Then
        recDeclaracoesFreq.MoveFirst
        recDeclaracoesFreq.Edit
        
        recDeclaracoesFreq.Fields("MENSALIDADE_EXT") = NumeroParaExtenso_Euro(recDeclaracoesFreq.Fields("MENSALIDADE"))
        
        recDeclaracoesFreq.Update
    Else
        MsgBox "O Utente indicado já não frequenta a Instituição !", vbInformation, "CAIF - Declarações de Frequência"
        GoTo SairDoProcedimento
    End If
    
    recDeclaracoesFreq.Close
    Set recDeclaracoesFreq = Nothing
    
    With rptDeclaracoesFreq
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Declarações de Frequência"
            GoTo SairDoProcedimento
        End If
'        .WindowParentHandle = fFrmMDIPrincipal.hwnd
'        .WindowTitle = "Declarações de Frequência"
'        .WindowState = crptMaximized
'        .WindowAllowDrillDown = False
'        .WindowBorderStyle = 2
'        .WindowControlBox = True
'        .WindowControls = True
'        .WindowMaxButton = False
'        .WindowMinButton = False
'        .WindowShowCloseBtn = True
'        .WindowShowExportBtn = False
'        .WindowShowGroupTree = False
'        .WindowShowNavigationCtls = True
'        .WindowShowPrintBtn = True
'        .WindowShowPrintSetupBtn = True
'        .WindowShowProgressCtls = True
'        .WindowShowZoomCtl = True
'        .WindowShowSearchBtn = False
'        .WindowShowRefreshBtn = False
        'Configura o destino e o numero de copias e de linhas para o Mapa
'        .Destination = crptToWindow
        .Destination = crptToPrinter
        .DataFiles(0) = cBDComNomeUtilizador
        .DataFiles(1) = cBDComNomeUtilizador
        .DataFiles(2) = cBDComNomeUtilizador
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_4 & "'"
        .Formulas(2) = "Titulo_3='" & gEmpresa.Linha4 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        .Formulas(4) = "Morada='" & gEmpresa.Linha1 & "'"
        .Formulas(5) = "Codigo Postal='" & gEmpresa.Linha2 & "'"
        .Formulas(6) = "Telefone='" & gEmpresa.Linha3 & "'"
        .Formulas(7) = "Ano Lectivo='" & txtAnoLec.Text & "'"
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("CAIF - Declarações de Frequência", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub cboNome_Click()
    ' poe o num de utente na combo
    cboNumUtente.Text = cboNome.Columns(1).Value
End Sub

Private Sub cboNome_DropDown()
    ' carrega a combo
    Call CarregacboNomeUtentesCAIF(cboNome, cboInstituicao.Text, cboSalas.Text)
End Sub

Private Sub cboNome_InitColumnProps()
    With cboNome
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
        .Columns(0).Caption = "Nome do Utente"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
End Sub

Private Sub cboNumUtente_Click()
    If Trim$(cboNumUtente.Text) <> vbNullString Then
        ' poe o num de utente na combo
        cboNome.Text = cboNumUtente.Columns(1).Value
    End If
End Sub

Private Sub cboNumUtente_DropDown()
    ' carrega a combo
    Call CarregacboNumUtentesCAIF(cboNumUtente, cboInstituicao.Text, cboSalas.Text)
End Sub

Private Sub cboNumUtente_InitColumnProps()
    With cboNumUtente
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
        .Columns(0).Caption = "Nº Utente"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
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
    For Each mBD In mWSDeclaracoesFreq.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSDeclaracoesFreq = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmDeclaracoesFrequencia = Nothing
End Sub



