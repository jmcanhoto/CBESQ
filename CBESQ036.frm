VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFListaUtentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Lista de Utentes"
   ClientHeight    =   7830
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
   ScaleHeight     =   7830
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame fraRangeUtentes 
      Enabled         =   0   'False
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
      Height          =   975
      Left            =   135
      TabIndex        =   15
      Top             =   5775
      Width           =   5610
      Begin GTMaskNum.GTMaskNum txtNumUtenteIni 
         Height          =   330
         Left            =   720
         TabIndex        =   17
         Top             =   375
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         CalcDropDown    =   0   'False
         BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskType        =   0
         DataType        =   1
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GTMaskNum.GTMaskNum txtNumUtenteFim 
         Height          =   330
         Left            =   2205
         TabIndex        =   19
         Top             =   375
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "9999"
         CalcDropDown    =   0   'False
         BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskType        =   0
         DataType        =   1
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GTMaskNum.GTMaskNum txtQuantidade 
         Height          =   330
         Left            =   4470
         TabIndex        =   21
         Top             =   375
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         CalcDropDown    =   0   'False
         BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskType        =   0
         DataType        =   1
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         Height          =   240
         Index           =   2
         Left            =   3270
         TabIndex        =   20
         Top             =   390
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         Caption         =   "de"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   16
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblLabel 
         Caption         =   "a"
         Height          =   255
         Index           =   1
         Left            =   1815
         TabIndex        =   18
         Top             =   390
         Width           =   225
      End
   End
   Begin VB.Frame fraSalas 
      Caption         =   " Valências"
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
      Left            =   135
      TabIndex        =   13
      Top             =   4800
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboSalas 
         Height          =   330
         Left            =   165
         TabIndex        =   14
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
      Left            =   135
      TabIndex        =   11
      Top             =   3795
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
         Height          =   330
         Left            =   165
         TabIndex        =   12
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
   Begin Crystal.CrystalReport rptListaUtentes 
      Left            =   2715
      Top             =   6900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraListaUtentes 
      Caption         =   " Lista de Utentes por "
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
      Height          =   3510
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome , Data Nasc. e Nº Seg. Social"
         Height          =   300
         Index           =   8
         Left            =   330
         TabIndex        =   10
         Top             =   2970
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome e B.I."
         Height          =   300
         Index           =   7
         Left            =   330
         TabIndex        =   9
         Top             =   2640
         Width           =   5205
      End
      Begin VB.CheckBox chkDatas 
         Caption         =   "Com Datas"
         Height          =   240
         Left            =   3225
         TabIndex        =   8
         Top             =   2310
         Width           =   1995
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Etiquetas dos Utentes"
         Height          =   300
         Index           =   6
         Left            =   330
         TabIndex        =   7
         Top             =   2310
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Carta aos Utentes"
         Height          =   300
         Index           =   5
         Left            =   330
         TabIndex        =   6
         Top             =   1980
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Valências com Mensalidade  e Próxima Mensalidade"
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   5
         Top             =   1650
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Valências com Valor da Mensalidade"
         Height          =   300
         Index           =   3
         Left            =   330
         TabIndex        =   4
         Top             =   1320
         Width           =   3900
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Valências"
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   3
         Top             =   990
         Width           =   3900
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nº de Utente"
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
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "2003"
      Top             =   6810
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "2010"
      Top             =   6810
      Width           =   1200
   End
End
Attribute VB_Name = "frmCAIFListaUtentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaUtentes_Idosos As Workspace
Dim mBDListaUtentes_Idosos As Database
Dim mBDListaUtentes_IdososTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

Dim i As Integer
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaUtentes_Idosos = DBEngine.CreateWorkspace("ListaUtentes_Idosos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaUtentes_Idosos = mWSListaUtentes_Idosos.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaUtentes_IdososTemp = mWSListaUtentes_Idosos.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Lista de Utentes-Abrir BD", Err.Number, Err.Description)
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
        .Columns(0).Caption = "Nome da Valência"
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
    
    Set mProcessamento = New Processamento

    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDListaUtentes_IdososTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDListaUtentes_Idosos.Execute cSql, dbFailOnError
    
    ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDListaUtentes_IdososTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDListaUtentes_Idosos.Execute cSql, dbFailOnError
    
    If optListaUtentes(6).Value Then
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM LISTA_UTENTES_IDOSOS_ETIQ;"
    Else
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM LISTA_UTENTES_IDOSOS;"
    End If
    ' apaga o registo em Temp32
    mBDListaUtentes_IdososTemp.Execute cSql, dbFailOnError
    
    If optListaUtentes(5).Value Then
        cSql = "INSERT INTO LISTA_UTENTES_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES_IDOSOS WHERE "
        cSql = cSql & "NUM_UTENTE BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    ElseIf optListaUtentes(6).Value Then
        cSql = "INSERT INTO LISTA_UTENTES_IDOSOS_ETIQ IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES_IDOSOS WHERE "
        cSql = cSql & "NUM_CAMA BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    Else
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_UTENTES_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES_IDOSOS WHERE ISNULL(DATA_SAIDA)"
    End If
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCod_Inst & "'"
    End If
    If cboSalas.Text <> "<Todas as Valências>" Then
        cSql = cSql & " AND COD_SALA='" & cCod_Sala & "'"
    End If
   
    If optListaUtentes(0).Value Then
        ' Lista de Utentes por Número
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ029.RPT"
    ElseIf optListaUtentes(1).Value Then
        ' Lista de Utentes por Nome
        cSql = cSql & " ORDER BY NOME ASC;"
        cNomeMapa = "CBESQ029.RPT"
    ElseIf optListaUtentes(2).Value Then
        ' Lista de Utentes por Valências
        cSql = cSql & ";"
        cNomeMapa = "CBESQ030.RPT"
    ElseIf optListaUtentes(3).Value Then
        ' Lista de Utentes por Salas com Valor das Mensalidades
        cSql = cSql & ";"
        cNomeMapa = "CBESQ039.RPT"
    ElseIf optListaUtentes(4).Value Then
        ' Lista de Utentes por Salas com Valor da Mensalidade e Próxima Mensalidade
        cSql = cSql & ";"
        cNomeMapa = "CBESQ040.RPT"
    ElseIf optListaUtentes(5).Value Then
        ' Lista de Carta aos Utentes
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ043.RPT"
    ElseIf optListaUtentes(6).Value Then
        ' Lista de Etiquetas
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        If chkDatas.Value = Checked Then
            cNomeMapa = "CBESQ047.RPT"
        Else
            cNomeMapa = "CBESQ046.RPT"
        End If
    ElseIf optListaUtentes(7).Value Then
        ' Lista de Nomes e B.I
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ050.RPT"
    ElseIf optListaUtentes(8).Value Then
        ' Lista de Nomes e Data Nasc e Nº Seg Social
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ051.RPT"
    End If
    
    ' insere o registo em Temp32
    mBDListaUtentes_Idosos.Execute cSql, dbFailOnError
    
    ' insere a Quantidade de Etiquetas
    If optListaUtentes(6).Value Then
        For i = 1 To CInt(txtQuantidade.Value) - 1
            ' insere o registo em Temp32
            mBDListaUtentes_Idosos.Execute cSql, dbFailOnError
        Next i
    End If

    With rptListaUtentes
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Lista de Utentes"
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
        .Destination = crptToWindow
'        .Destination = crptToPrinter
        If optListaUtentes(5).Value Then
            .DataFiles(0) = cBDComNomeUtilizador
            .DataFiles(1) = vbNullString
            .DataFiles(2) = vbNullString
        ElseIf optListaUtentes(6).Value Then
            .DataFiles(0) = cBDComNomeUtilizador
            .DataFiles(1) = vbNullString
            .DataFiles(2) = vbNullString
        Else
            .DataFiles(0) = cBDComNomeUtilizador
            .DataFiles(1) = cBDComNomeUtilizador
            .DataFiles(2) = cBDComNomeUtilizador
        End If
        .PrintFileLinesPerPage = 60
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        If Not optListaUtentes(5).Value Then
            .Formulas(1) = "Titulo_1='" & Mapa.Titulo_1 & "'"
            .Formulas(2) = "Titulo_2='" & Mapa.Titulo_4 & "'"
            .Formulas(3) = "Titulo_3='" & Mapa.Titulo_5 & "'"
        End If
        If optListaUtentes(0).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nº de Utente'"
        ElseIf optListaUtentes(1).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nome'"
        ElseIf optListaUtentes(2).Value Then
            .Formulas(4) = "Descricao Parametros Mapa=''"
        ElseIf optListaUtentes(3).Value Then
            .Formulas(4) = "Descricao Parametros Mapa=''"
        ElseIf optListaUtentes(4).Value Then
            .Formulas(4) = "Descricao Parametros Mapa=''"
        End If
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("CAIF - Lista Utentes", Err.Number, Err.Description)
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
    For Each mBD In mWSListaUtentes_Idosos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaUtentes_Idosos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFListaUtentes = Nothing
End Sub


Private Sub optListaUtentes_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 2, 3, 4
            fraRangeUtentes.Enabled = False
            txtQuantidade.Value = 1
        Case 5
            fraRangeUtentes.Enabled = True
            txtQuantidade.Enabled = False
            txtQuantidade.Value = 1
        Case 6
            fraRangeUtentes.Enabled = True
            txtQuantidade.Enabled = True
            txtQuantidade.Value = 1
    End Select
    txtNumUtenteIni.Value = 1
    txtNumUtenteFim.Value = 9999
End Sub


