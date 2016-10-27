VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFListaInscricoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Lista de Inscrições"
   ClientHeight    =   6795
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
   ScaleHeight     =   6795
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame Frame2 
      Caption         =   "Sexo"
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
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   5565
      Begin VB.OptionButton optSexo 
         Height          =   300
         Index           =   2
         Left            =   3510
         TabIndex        =   19
         Top             =   285
         Value           =   -1  'True
         Width           =   210
      End
      Begin VB.OptionButton optSexo 
         Height          =   300
         Index           =   0
         Left            =   675
         TabIndex        =   16
         Top             =   300
         Width           =   210
      End
      Begin VB.OptionButton optSexo 
         Height          =   300
         Index           =   1
         Left            =   2115
         TabIndex        =   15
         Top             =   300
         Width           =   210
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Ambos"
         Height          =   240
         Index           =   0
         Left            =   3780
         TabIndex        =   20
         Top             =   285
         Width           =   645
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Masculino"
         Height          =   240
         Index           =   18
         Left            =   945
         TabIndex        =   18
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Feminino"
         Height          =   240
         Index           =   19
         Left            =   2385
         TabIndex        =   17
         Top             =   300
         Width           =   825
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
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboSalas 
         Height          =   330
         Left            =   165
         TabIndex        =   13
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
      Left            =   150
      TabIndex        =   8
      Top             =   2985
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
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
         Text            =   "<Todas as Instituições>"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
   End
   Begin Crystal.CrystalReport rptListaInscricoes 
      Left            =   2730
      Top             =   5850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraDatas 
      Caption         =   " Intervalo de Datas de Nascimento "
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
      TabIndex        =   3
      Top             =   1995
      Width           =   5610
      Begin GTMaskDate.GTMaskDate gtmData 
         Height          =   405
         Index           =   0
         Left            =   810
         TabIndex        =   5
         Top             =   330
         Width           =   1500
         _Version        =   65537
         _ExtentX        =   2646
         _ExtentY        =   714
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
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalWidth        =   200
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalDayCaption1  =   "Dom"
         CalDayCaption2  =   "Seg"
         CalDayCaption3  =   "Ter"
         CalDayCaption4  =   "Qua"
         CalDayCaption5  =   "Qui"
         CalDayCaption6  =   "Sex"
         CalDayCaption7  =   "Sáb"
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin GTMaskDate.GTMaskDate gtmData 
         Height          =   405
         Index           =   1
         Left            =   3900
         TabIndex        =   7
         Top             =   330
         Width           =   1500
         _Version        =   65537
         _ExtentX        =   2646
         _ExtentY        =   714
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
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalWidth        =   200
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalDayCaption1  =   "Dom"
         CalDayCaption2  =   "Seg"
         CalDayCaption3  =   "Ter"
         CalDayCaption4  =   "Qua"
         CalDayCaption5  =   "Qui"
         CalDayCaption6  =   "Sex"
         CalDayCaption7  =   "Sáb"
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Caption         =   "a"
         Height          =   255
         Index           =   1
         Left            =   3420
         TabIndex        =   6
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblLabel 
         Caption         =   "de"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   390
         Width           =   225
      End
   End
   Begin VB.Frame fraListaInscricoes 
      Caption         =   " Lista Inscrições por "
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
      Height          =   1800
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5595
      Begin VB.OptionButton optListaInscricoes 
         Caption         =   "Data de Nascimento"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton optListaInscricoes 
         Caption         =   "Nº de Inscrição"
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
      TabIndex        =   11
      Tag             =   "2003"
      ToolTipText     =   "Sair da Lista de Inscrições"
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2010"
      ToolTipText     =   "Imprimir a Lista de Inscrições seleccionada"
      Top             =   5760
      Width           =   1200
   End
End
Attribute VB_Name = "frmCAIFListaInscricoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaInscricoes_Idosos As Workspace
Dim mBDListaInscricoes_Idosos As Database
Dim mBDListaInscricoes_IdososTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaInscricoes_Idosos = DBEngine.CreateWorkspace("ListaInscricoes_Idosos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaInscricoes_Idosos = mWSListaInscricoes_Idosos.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaInscricoes_IdososTemp = mWSListaInscricoes_Idosos.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Lista de Inscrições-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    cboSalas.Text = "<Todas as Valências>"
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
    
    Set mProcessamento = New Processamento

    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDListaInscricoes_IdososTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDListaInscricoes_Idosos.Execute cSql, dbFailOnError
    
     ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDListaInscricoes_IdososTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDListaInscricoes_Idosos.Execute cSql, dbFailOnError
    
    ' apaga os registos da Temp32
    cSql = "DELETE * FROM LISTA_INSCRICOES_IDOSOS;"
    ' apaga o registo em Temp32
    mBDListaInscricoes_IdososTemp.Execute cSql, dbFailOnError

    ' carrega a variavela com o Sql
    cSql = "INSERT INTO LISTA_INSCRICOES_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM INSCRICOES_IDOSOS WHERE UTENTE=False AND " & _
            "DATA_NASC BETWEEN #" & Format$(gtmData(0).DateValue, "yyyy/mm/dd") & "# AND #" & _
            Format$(gtmData(1).DateValue, "yyyy/mm/dd") & "#"
    
    ' se seleccionar uma instituiçao
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND (COD_INST='" & cCod_Inst & "' OR COD_INST='000')"
    End If
    If cboSalas.Text <> "<Todas as Valências>" Then
        cSql = cSql & " AND COD_SALA='" & cCod_Sala & "'"
    End If
    
    If optSexo(0).Value = True Then
        cSql = cSql & " AND SEXO='M'"
    ElseIf optSexo(1).Value = True Then
        cSql = cSql & " AND SEXO='F'"
    End If
    
    If optListaInscricoes(0).Value Then
        cSql = cSql & " ORDER BY NUM_INSCRICAO ASC;"
    ElseIf optListaInscricoes(1).Value Then
        cSql = cSql & " ORDER BY DATA_NASC DESC,NUM_INSCRICAO ASC;"
    End If
    
    ' insere o registo em Temp32
    mBDListaInscricoes_Idosos.Execute cSql, dbFailOnError

    With rptListaInscricoes
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Lista de Inscrições"
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
        If optListaInscricoes(0).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nº de Inscrição, entre " & gtmData(0).DateValue & _
                            " e " & gtmData(1).DateValue & "'"
        ElseIf optListaInscricoes(1).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Data de Nascimento, entre " & gtmData(0).DateValue & _
                            " e " & gtmData(1).DateValue & "'"
        End If
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Lista Inscrições", Err.Number, Err.Description)
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
    
    Call CarregacboInstituicao(cboInstituicao)

    ' carrega a Data
    gtmData(0).DateValue = DateAdd("yyyy", -100, Date)
'    gtmData(0).DateValue = DateAdd("d", -30, Date)
    gtmData(1).DateValue = Date
    cNomeMapa = "CBESQ027.RPT"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSListaInscricoes_Idosos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaInscricoes_Idosos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFListaInscricoes = Nothing
End Sub


Private Sub gtmData_DropDown(Index As Integer)
    If Not IsDate(gtmData(Index).Text) Then
        gtmData(Index).DateValue = Date
    End If
End Sub


