VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaFuncionarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Funcionários"
   ClientHeight    =   4815
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
   ScaleHeight     =   4815
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin Crystal.CrystalReport rptListaFuncionarios 
      Left            =   2730
      Top             =   3885
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraListaFuncionarios 
      Caption         =   " Lista de Funcionários por "
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
      Height          =   2700
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Mapa para o Seguro"
         Height          =   300
         Index           =   5
         Left            =   330
         TabIndex        =   8
         Top             =   1980
         Width           =   3900
      End
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Instituição"
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   5
         Top             =   1650
         Width           =   3900
      End
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Categoria"
         Height          =   300
         Index           =   3
         Left            =   330
         TabIndex        =   4
         Top             =   1320
         Width           =   3900
      End
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Data de Admissão"
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   3
         Top             =   990
         Width           =   3900
      End
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Nome"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton optListaFuncionarios 
         Caption         =   "Nº de Funcionário"
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
      TabIndex        =   7
      Tag             =   "2003"
      Top             =   3795
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "2010"
      Top             =   3795
      Width           =   1200
   End
End
Attribute VB_Name = "frmListaFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaFuncionarios As Workspace
Dim mBDListaFuncionarios As Database
Dim mBDListaFuncionariosTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaFuncionarios = DBEngine.CreateWorkspace("ListaFuncionarios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaFuncionarios = mWSListaFuncionarios.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaFuncionariosTemp = mWSListaFuncionarios.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Lista de Funcionários-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    Dim cCod_Inst
    Dim cCod_Sala
    
    Set mProcessamento = New Processamento

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDListaFuncionariosTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDListaFuncionarios.Execute cSql, dbFailOnError
    
      ' apaga registos da TABICATEGORIA
    cSql = "DELETE * FROM TABCATEGORIA;"
    mBDListaFuncionariosTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABCATEGORIA
    cSql = "INSERT INTO TABCATEGORIA IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABCATEGORIA"
    mBDListaFuncionarios.Execute cSql, dbFailOnError
   
    ' apaga os registos da Temp32
    cSql = "DELETE * FROM LISTA_FUNCIONARIOS;"
    ' apaga o registo em Temp32
    mBDListaFuncionariosTemp.Execute cSql, dbFailOnError

    ' carrega a variavela com o Sql
    cSql = "INSERT INTO LISTA_FUNCIONARIOS IN '" & cBDComNomeUtilizador & _
        "' SELECT * FROM FUNCIONARIOS WHERE ISNULL(DATA_DEMISSAO)"
    
    If optListaFuncionarios(0).Value Then
        cNomeMapa = "CBESQ017.RPT"
        ' Lista de Funcionarios por Número
        cSql = cSql & " ORDER BY NUM_FUNCIONARIO ASC;"
    ElseIf optListaFuncionarios(1).Value Then
        cNomeMapa = "CBESQ017.RPT"
        ' Lista de Funcionarios por Nome
        cSql = cSql & " ORDER BY NOME ASC;"
    ElseIf optListaFuncionarios(2).Value Then
        cNomeMapa = "CBESQ017.RPT"
        ' Lista de Funcionarios por Data de Admissão
        cSql = cSql & " ORDER BY DATA_ADMISSAO ASC;"
    ElseIf optListaFuncionarios(3).Value Then
        cNomeMapa = "CBESQ044.RPT"
        ' Lista de Funcionarios por Categoria
        cSql = cSql & " ORDER BY COD_CATEGORIA ASC;"
    ElseIf optListaFuncionarios(4).Value Then
        cNomeMapa = "CBESQ045.RPT"
        ' Lista de Funcionarios por Instituição
        cSql = cSql & " ORDER BY COD_INST ASC;"
    ElseIf optListaFuncionarios(5).Value Then
        cNomeMapa = "CBESQ048.RPT"
    End If
    
    ' insere o registo em Temp32
    mBDListaFuncionarios.Execute cSql, dbFailOnError

    With rptListaFuncionarios
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Lista de Funcionários"
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
        .PrintFileLinesPerPage = 60
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        If optListaFuncionarios(0).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nº de Funcionário'"
        ElseIf optListaFuncionarios(1).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nome do Funcionário'"
        ElseIf optListaFuncionarios(2).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Data de Admissão'"
        ElseIf optListaFuncionarios(3).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Categoria'"
        ElseIf optListaFuncionarios(4).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Instituição'"
        ElseIf optListaFuncionarios(5).Value Then
            .Formulas(4) = "Descricao Parametros Mapa=''"
        End If
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Lista Funcionários", Err.Number, Err.Description)
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
    For Each mBD In mWSListaFuncionarios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaFuncionarios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmListaFuncionarios = Nothing
End Sub


