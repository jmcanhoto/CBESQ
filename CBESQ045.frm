VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFListaRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Listas de Recibos"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
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
   ScaleHeight     =   5625
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
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
      TabIndex        =   6
      Top             =   3540
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboSalas 
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
         Text            =   "<Todas as Valências>"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Enabled         =   0   'False
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
      TabIndex        =   4
      Top             =   2535
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
         Height          =   330
         Left            =   165
         TabIndex        =   5
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
         Enabled         =   0   'False
      End
   End
   Begin Crystal.CrystalReport rptListaRecibos 
      Left            =   2730
      Top             =   4665
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraListaUtentes 
      Caption         =   " Lista de "
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
      Height          =   2250
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin VB.OptionButton optListaRecibos 
         Caption         =   "Recibos Pagos (Dinheiro / Cheque)"
         Height          =   300
         Index           =   3
         Left            =   330
         TabIndex        =   11
         Top             =   1320
         Width           =   3900
      End
      Begin VB.OptionButton optListaRecibos 
         Caption         =   "Recibos Pagos (Transferências Bancárias)"
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   10
         Top             =   1650
         Width           =   4380
      End
      Begin VB.OptionButton optListaRecibos 
         Caption         =   "Recibos em Divida"
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   3
         Top             =   990
         Width           =   3900
      End
      Begin VB.OptionButton optListaRecibos 
         Caption         =   "Recibos Criados para Pagamento"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton optListaRecibos 
         Caption         =   "Recibos a Criar"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   3900
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2003"
      Top             =   4575
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2010"
      Top             =   4575
      Width           =   1200
   End
End
Attribute VB_Name = "frmCAIFListaRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaRecibos As Workspace
Dim mBDListaRecibos As Database
Dim mBDListaRecibosTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa
Dim iRespMsgBox
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaRecibos = DBEngine.CreateWorkspace("CAIFListaRecibos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaRecibos = mWSListaRecibos.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaRecibosTemp = mWSListaRecibos.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Lista de Recibos-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    cboSalas.Text = "<Todas as Valências>"
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
        
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    Dim cCod_Inst
    Dim cCod_Sala
    Dim cTipo_Pag
    
    Set mProcessamento = New Processamento

    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDListaRecibosTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDListaRecibos.Execute cSql, dbFailOnError
    
    ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDListaRecibosTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDListaRecibos.Execute cSql, dbFailOnError

    ' apaga os registos da Temp32
    cSql = "DELETE * FROM LISTA_RECIBOS_IDOSOS;"
    ' apaga o registo em Temp32
    mBDListaRecibosTemp.Execute cSql, dbFailOnError

    If optListaRecibos(0).Value Then
        ' Lista de Recibos em criacçao
        cNomeMapa = "CBESQ032.RPT"
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_RECIBOS_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM RECIBOS_IDOSOS WHERE ISNULL(ESTADO_REC)"
    ElseIf optListaRecibos(1).Value Then
        ' Lista de Recibos Criados para Pagamento
        cNomeMapa = "CBESQ033.RPT"
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_RECIBOS_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM RECIBOS_IDOSOS WHERE ESTADO_REC='P'"
    ElseIf optListaRecibos(2).Value Then
        ' Lista de Recibos em Divida
        cNomeMapa = "CBESQ034.RPT"
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_RECIBOS_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM RECIBOS_IDOSOS WHERE (ESTADO_REC='P' OR ESTADO_REC='D')"
    ElseIf optListaRecibos(3).Value Then
        cTipo_Pag = "D"
        ' Lista de Recibos Pagos
        cNomeMapa = "CBESQ035.RPT"
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_RECIBOS_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM RECIBOS_IDOSOS WHERE ESTADO_REC='L' AND TIPO_PAG='D'"
    ElseIf optListaRecibos(4).Value Then
        cTipo_Pag = "T"
        ' Lista de Recibos Pagos
        cNomeMapa = "CBESQ035.RPT"
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_RECIBOS_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM RECIBOS_IDOSOS WHERE ESTADO_REC='L' AND TIPO_PAG='T'"
    End If
    
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCod_Inst & "'"
    End If
    If cboSalas.Text <> "<Todas as Valências>" Then
        cSql = cSql & " AND COD_SALA='" & cCod_Sala & "'"
    End If
    cSql = cSql & ";"
    
    ' insere o registo em Temp32
    mBDListaRecibos.Execute cSql, dbFailOnError

    With rptListaRecibos
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Lista de Recibos"
        .WindowState = crptMaximized
        .WindowAllowDrillDown = False
        .WindowBorderStyle = 2
        .WindowControlBox = True
        .WindowControls = True
        .WindowMaxButton = False
        .WindowMinButton = False
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = False
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowProgressCtls = True
        .WindowShowZoomCtl = True
        .WindowShowSearchBtn = False
        .WindowShowRefreshBtn = False
        'Configura o destino e o numero de copias e de linhas para o Mapa
        If optListaRecibos(3).Value Or optListaRecibos(4).Value Then
'            .Destination = crptToWindow
            .Destination = crptToPrinter
        Else
            .Destination = crptToWindow
        End If
        .DataFiles(0) = cBDComNomeUtilizador
        .DataFiles(1) = cBDComNomeUtilizador
        .DataFiles(2) = cBDComNomeUtilizador
        .PrintFileLinesPerPage = 60
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_4 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_5 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        'executa o Report
        .Action = 1
    End With
        
    If optListaRecibos(3).Value Or optListaRecibos(4).Value Then
        Dim qryPassaHistorico As QueryDef

        iRespMsgBox = MsgBox("Confirme se a Lista de Recibos Pagos" & _
                            vbCrLf & "foi impressa correctamente ?", vbQuestion + vbYesNo + vbDefaultButton2, _
                            "Lista de Recibos Pagos")
        If iRespMsgBox = vbYes Then
            Set qryPassaHistorico = mBDListaRecibos.QueryDefs("CAIF RECIBOS Altera Estado TipoPag")

            qryPassaHistorico.Parameters("Estado Novo") = "H"
            qryPassaHistorico.Parameters("Estado Velho") = "L"
            qryPassaHistorico.Parameters("Utiliz") = gUtilizador.Nome
            qryPassaHistorico.Parameters("TipoPag") = cTipo_Pag

            ' executa a inserção
            qryPassaHistorico.Execute dbFailOnError
        End If
    End If
        
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Lista de Recibos", Err.Number, Err.Description)
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
    For Each mBD In mWSListaRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFListaRecibos = Nothing
End Sub


Private Sub optListaRecibos_Click(Index As Integer)
Select Case Index
    Case 3, 4
        cboInstituicao.Text = "<Todas as Instituições>"
        cboInstituicao.Enabled = False
        cboSalas.Text = "<Todas as Valências>"
        cboSalas.Enabled = False
    Case Else
        cboInstituicao.Enabled = True
        cboSalas.Enabled = True
End Select
End Sub


