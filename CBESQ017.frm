VERSION 5.00
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Sócios"
   ClientHeight    =   6075
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
   ScaleHeight     =   6075
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame fraRangeSocios 
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
      Height          =   1410
      Left            =   150
      TabIndex        =   4
      Top             =   3465
      Width           =   5610
      Begin VB.TextBox txtAnoQuotas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   720
         MaxLength       =   4
         TabIndex        =   14
         Top             =   870
         Width           =   930
      End
      Begin GTMaskNum.GTMaskNum txtNumSocioIni 
         Height          =   330
         Left            =   720
         TabIndex        =   6
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
      Begin GTMaskNum.GTMaskNum txtNumSocioFim 
         Height          =   330
         Left            =   3240
         TabIndex        =   8
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
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Ano"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   13
         Top             =   900
         Width           =   360
      End
      Begin VB.Label lblLabel 
         Caption         =   "a"
         Height          =   255
         Index           =   1
         Left            =   2850
         TabIndex        =   7
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblLabel 
         Caption         =   "de"
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   390
         Width           =   225
      End
   End
   Begin Crystal.CrystalReport rptListaSocios 
      Left            =   2730
      Top             =   5130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraListaSocios 
      Caption         =   " Lista de Sócios por "
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
      Height          =   3225
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Quotas por Morada"
         Height          =   300
         Index           =   7
         Left            =   330
         TabIndex        =   17
         Top             =   2640
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Moradas"
         Height          =   300
         Index           =   6
         Left            =   330
         TabIndex        =   16
         Top             =   2310
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Carta aos Sócios"
         Height          =   300
         Index           =   5
         Left            =   330
         TabIndex        =   15
         Top             =   1980
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Quotas"
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   12
         Top             =   1650
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Quotas em Dívida"
         Height          =   300
         Index           =   3
         Left            =   330
         TabIndex        =   11
         Top             =   1320
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Etiquetas com a Morada do Sócio"
         Height          =   300
         Index           =   2
         Left            =   345
         TabIndex        =   3
         Top             =   990
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Nome"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   2
         Top             =   660
         Width           =   3900
      End
      Begin VB.OptionButton optListaSocios 
         Caption         =   "Nº de Sócio"
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
      TabIndex        =   10
      Tag             =   "2003"
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2010"
      Top             =   5040
      Width           =   1200
   End
End
Attribute VB_Name = "frmListaSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaSocios As Workspace
Dim mBDListaSocios As Database
Dim mBDListaSociosTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaSocios = DBEngine.CreateWorkspace("ListaSocios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaSocios = mWSListaSocios.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaSociosTemp = mWSListaSocios.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Lista de Socios-Abrir BD", Err.Number, Err.Description)
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

    ' apaga os registos da Temp32
    If optListaSocios(2).Value Then
        cSql = "DELETE * FROM SOCIOS_ETIQ_MORADAS;"
    ElseIf optListaSocios(3).Value Then
        cSql = "DELETE * FROM SOCIOS_QUOTAS_DIVIDA;"
    ElseIf optListaSocios(4).Value Then
        cSql = "DELETE * FROM SOCIOS_QUOTAS;"
    ElseIf optListaSocios(7).Value Then
        cSql = "DELETE * FROM SOCIOS_QUOTAS;"
    Else
        cSql = "DELETE * FROM LISTA_SOCIOS;"
    End If
    ' apaga o registo em Temp32
    mBDListaSociosTemp.Execute cSql, dbFailOnError

    ' carrega a variavela com o Sql
    If optListaSocios(2).Value Then
        cSql = "INSERT INTO SOCIOS_ETIQ_MORADAS IN '" & _
            cBDComNomeUtilizador & "' SELECT NUM_SOCIO,"
        cSql = cSql & "NOME,"
        cSql = cSql & "MORADA,"
        cSql = cSql & "COD_POSTAL,"
        cSql = cSql & "LOCAL FROM SOCIOS WHERE "
        cSql = cSql & "NUM_SOCIO BETWEEN " & txtNumSocioIni.Value & _
                " AND " & txtNumSocioFim.Value & " AND ISNULL(DATA_DEMISSAO)"
    ElseIf optListaSocios(3).Value Then
        ' insere os registos na temp32
        cSql = "INSERT INTO SOCIOS_QUOTAS_DIVIDA IN '" & _
            cBDComNomeUtilizador & "' SELECT QUOTAS.NUM_SOCIO,"
        cSql = cSql & "QUOTAS.ANO,"
        cSql = cSql & "QUOTAS.TRIMESTRE,"
        cSql = cSql & "QUOTAS.VALOR,"
        cSql = cSql & "QUOTAS.DATA_PAG, SOCIOS.NOME FROM QUOTAS INNER JOIN SOCIOS ON QUOTAS.NUM_SOCIO = SOCIOS.NUM_SOCIO "
        cSql = cSql & "WHERE (((QUOTAS.NUM_SOCIO) BETWEEN " & txtNumSocioIni.Value & _
                    " AND " & txtNumSocioFim.Value & ") AND ((QUOTAS.DATA_PAG) Is Null))"
'        cSql = cSql & "WHERE NUM_SOCIO BETWEEN " & txtNumSocioIni.Value & _
'                " AND " & txtNumSocioFim.Value & " AND ISNULL(DATA_PAG)"
    ElseIf optListaSocios(4).Value Then
        ' insere os registos na temp32
        cSql = "INSERT INTO SOCIOS_QUOTAS IN '" & _
            cBDComNomeUtilizador & "' SELECT NUM_SOCIO,"
        cSql = cSql & "NOME,"
        cSql = cSql & "MORADA,"
        cSql = cSql & "COD_POSTAL,"
        cSql = cSql & "LOCAL,"
        cSql = cSql & "TELEFONE,"
        cSql = cSql & "MORADA_COB,"
        cSql = cSql & "COD_POSTAL_COB,"
        cSql = cSql & "LOCAL_COB,"
        cSql = cSql & "QUOTA FROM SOCIOS WHERE "
        cSql = cSql & "NUM_SOCIO BETWEEN " & txtNumSocioIni.Value & _
                " AND " & txtNumSocioFim.Value & " AND ISNULL(DATA_DEMISSAO)"
    ElseIf optListaSocios(5).Value Then
        cSql = "INSERT INTO LISTA_SOCIOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM SOCIOS WHERE "
        cSql = cSql & "NUM_SOCIO BETWEEN " & txtNumSocioIni.Value & _
                " AND " & txtNumSocioFim.Value & " AND ISNULL(DATA_DEMISSAO)"
    ElseIf optListaSocios(7).Value Then
        ' insere os registos na temp32
        cSql = "INSERT INTO SOCIOS_QUOTAS IN '" & _
            cBDComNomeUtilizador & "' SELECT NUM_SOCIO,"
        cSql = cSql & "NOME,"
        cSql = cSql & "MORADA,"
        cSql = cSql & "COD_POSTAL,"
        cSql = cSql & "LOCAL,"
        cSql = cSql & "TELEFONE,"
        cSql = cSql & "MORADA_COB,"
        cSql = cSql & "COD_POSTAL_COB,"
        cSql = cSql & "LOCAL_COB,"
        cSql = cSql & "QUOTA FROM SOCIOS WHERE "
        cSql = cSql & "NUM_SOCIO BETWEEN " & txtNumSocioIni.Value & _
                " AND " & txtNumSocioFim.Value & " AND ISNULL(DATA_DEMISSAO)"
    Else
        cSql = "INSERT INTO LISTA_SOCIOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM SOCIOS WHERE ISNULL(DATA_DEMISSAO)"
    End If
    
    If optListaSocios(0).Value Then
        ' Lista de Socios por Número
        cSql = cSql & " ORDER BY NUM_SOCIO ASC;"
        cNomeMapa = "CBESQ015.RPT"
    ElseIf optListaSocios(1).Value Then
        ' Lista de Socios por Data de Nascimento
        cSql = cSql & " ORDER BY NOME ASC;"
        cNomeMapa = "CBESQ015.RPT"
    ElseIf optListaSocios(2).Value Then
        ' Lista de Socios para Etiquetas com moradas
        cSql = cSql & " ORDER BY NUM_SOCIO;"
        cNomeMapa = "CBESQ021.RPT"
    ElseIf optListaSocios(3).Value Then
        ' Lista de Quotas em Dívida dos Socios
        cSql = cSql & ";"
        cNomeMapa = "CBESQ056.RPT"
    ElseIf optListaSocios(4).Value Then
        ' Lista de Quotas de Socios
        cSql = cSql & " ORDER BY NUM_SOCIO;"
        cNomeMapa = "CBESQ020.RPT"
    ElseIf optListaSocios(5).Value Then
        ' Lista de Carta aos Socios
        cSql = cSql & " ORDER BY NUM_SOCIO;"
        cNomeMapa = "CBESQ025.RPT"
    ElseIf optListaSocios(6).Value Then
        ' Lista de Moradas dos Socios
        cSql = cSql & ";"
        cNomeMapa = "CBESQ055.RPT"
    ElseIf optListaSocios(7).Value Then
        ' Lista de Quotas de Socios
        cSql = cSql & " ORDER BY MORADA;"
        cNomeMapa = "CBESQ020.RPT"
    End If
    
    ' insere o registo em Temp32
    mBDListaSocios.Execute cSql, dbFailOnError

    With rptListaSocios
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
        .WindowTitle = "Lista de Socios"
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
        If optListaSocios(2).Value Then
            .Destination = crptToPrinter
        Else
            .Destination = crptToWindow
        End If
        .DataFiles(0) = cBDComNomeUtilizador
        .PrintFileLinesPerPage = 60
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        If Not optListaSocios(2).Value Then
            .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
            .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
            .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
            .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        End If
        If optListaSocios(0).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nº de Sócio'"
        ElseIf optListaSocios(1).Value Then
            .Formulas(4) = "Descricao Parametros Mapa='Ordenada por Nome do Sócio'"
        ElseIf optListaSocios(4).Value Then
            .Formulas(4) = "Ano='" & txtAnoQuotas.Text & "'"
        ElseIf optListaSocios(7).Value Then
            .Formulas(4) = "Ano='" & txtAnoQuotas.Text & "'"
        End If
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Lista Socios", Err.Number, Err.Description)
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
    
    txtNumSocioIni.Value = 1
    txtNumSocioFim.Value = 9999
    txtAnoQuotas.Text = Year(Date)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSListaSocios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaSocios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmListaSocios = Nothing
End Sub


Private Sub optListaSocios_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 3, 6
            fraRangeSocios.Enabled = False
        Case 2, 5
            fraRangeSocios.Enabled = True
            txtAnoQuotas.Enabled = False
        Case 4
            fraRangeSocios.Enabled = True
            txtAnoQuotas.Enabled = True
    End Select
    
    txtNumSocioIni.Value = 1
    txtNumSocioFim.Value = 9999
    txtAnoQuotas.Text = Year(Date)
End Sub

