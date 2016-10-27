VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRendimentoPercapitaMapa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendimento Percapita (Mapa)"
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
   Begin Crystal.CrystalReport rptRendPercapitaMapa 
      Left            =   2730
      Top             =   3795
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
      TabIndex        =   8
      Tag             =   "2003"
      Top             =   3705
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "2010"
      Top             =   3705
      Width           =   1200
   End
End
Attribute VB_Name = "frmRendimentoPercapitaMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSRendPercapitaMapa As Workspace
Dim mBDRendPercapitaMapa As Database
Dim mBDRendPercapitaMapaTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSRendPercapitaMapa = DBEngine.CreateWorkspace("RendPercapitaMapa", gUtilizador.Nome, gUtilizador.Password)
    Set mBDRendPercapitaMapa = mWSRendPercapitaMapa.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDRendPercapitaMapaTemp = mWSRendPercapitaMapa.OpenDatabase(cBDComNomeUtilizador)

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
    Dim cCod_Inst
    Dim cCod_Sala
    
    Set mProcessamento = New Processamento

    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)
    
    cNomeMapa = "CBESQ024M.RPT"

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDRendPercapitaMapaTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDRendPercapitaMapa.Execute cSql, dbFailOnError
    
    ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDRendPercapitaMapaTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDRendPercapitaMapa.Execute cSql, dbFailOnError
    
    ' apaga os registos da Temp32
    cSql = "DELETE * FROM LISTA_RENDIMENTO_PERCAPITA_MAPA;"
    ' apaga o registo em Temp32
    mBDRendPercapitaMapaTemp.Execute cSql, dbFailOnError
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO LISTA_RENDIMENTO_PERCAPITA_MAPA (NUM_UTENTE,COD_INST,COD_SALA,NOME,AGREGADO," & _
                                                    "MENSALIDADE_BASE,MENSALIDADE," & _
                                                    "R_TD_1,R_P_1,R_PA_1,R_TI_1,R_R_1,R_RSI_1,R_SD_1,R_AF_1,R_O_DESC,R_O_1," & _
                                                    "D_IRS_1,D_SS_1,D_IMI_1,D_JEH_1,R_TD_2,R_P_2,R_PA_2,R_TI_2,R_R_2,R_RSI_2," & _
                                                    "R_SD_2,R_AF_2,R_O_2,D_IRS_2,D_SS_2,D_IMI_2,D_JEH_2," & _
                                                    "D_SC_1,D_T_1,D_O_DESC,D_O_1,D_SC_2,D_T_2,D_O_2) IN '" & cBDComNomeUtilizador & "' " & _
        "SELECT UTENTES.NUM_UTENTE,UTENTES.COD_INST,UTENTES.COD_SALA," & _
        "UTENTES.NOME,UTENTES.AGREGADO,UTENTES.PROX_MENSALIDADE_BASE,UTENTES.PROX_MENSALIDADE," & _
        "UTENTES.R_TD_1,UTENTES.R_P_1,UTENTES.R_PA_1,UTENTES.R_TI_1,UTENTES.R_R_1,UTENTES.R_RSI_1,UTENTES.R_SD_1,UTENTES.R_AF_1,UTENTES.R_O_DESC,UTENTES.R_O_1," & _
        "UTENTES.D_IRS_1,UTENTES.D_SS_1,UTENTES.D_IMI_1,UTENTES.D_JEH_1," & _
        "UTENTES.R_TD_2,UTENTES.R_P_2,UTENTES.R_PA_2,UTENTES.R_TI_2,UTENTES.R_R_2,UTENTES.R_RSI_2,UTENTES.R_SD_2,UTENTES.R_AF_2,UTENTES.R_O_2," & _
        "UTENTES.D_IRS_2,UTENTES.D_SS_2,UTENTES.D_IMI_2,UTENTES.D_JEH_2," & _
        "UTENTES.D_SC_1,UTENTES.D_T_1,UTENTES.D_O_DESC,UTENTES.D_O_1,UTENTES.D_SC_2,UTENTES.D_T_2,UTENTES.D_O_2 FROM UTENTES WHERE ISNULL(DATA_SAIDA)"
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND UTENTES.COD_INST='" & cCod_Inst & "'"
    End If
    If cboSalas.Text <> "<Todas as Salas>" Then
        cSql = cSql & " AND UTENTES.COD_SALA='" & cCod_Sala & "'"
    End If
    If cboNumUtente.Text <> vbNullString Then
        cSql = cSql & " AND UTENTES.NUM_UTENTE=" & cboNumUtente.Text & " ;"
    End If
    ' insere o registo em Temp32
    mBDRendPercapitaMapa.Execute cSql, dbFailOnError
    
    With rptRendPercapitaMapa
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Declarações de Frequência"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Rendimento Percapita"
        .WindowState = crptMaximized
        .WindowAllowDrillDown = False
        .WindowBorderStyle = 2
        .WindowControlBox = True
        .WindowControls = True
        .WindowMaxButton = False
        .WindowMinButton = False
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = False
        .WindowShowGroupTree = False
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowProgressCtls = True
        .WindowShowZoomCtl = True
        .WindowShowSearchBtn = False
        .WindowShowRefreshBtn = False
        'Configura o destino e o numero de copias e de linhas para o Mapa
        .Destination = crptToWindow
'        .Destination = crptToPrinter
        .DataFiles(0) = cBDComNomeUtilizador
        .DataFiles(1) = cBDComNomeUtilizador
        .DataFiles(2) = cBDComNomeUtilizador
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Rendimento Percápita", Err.Number, Err.Description)
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
    Call CarregacboNomeUtentes(cboNome, cboInstituicao.Text, cboSalas.Text)
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
    Call CarregacboNumUtentes(cboNumUtente, cboInstituicao.Text, cboSalas.Text)
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
    For Each mBD In mWSRendPercapitaMapa.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSRendPercapitaMapa = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmRendimentoPercapitaMapa = Nothing
End Sub



