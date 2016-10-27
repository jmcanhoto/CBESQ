VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListaUtentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Utentes"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
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
   ScaleHeight     =   5835
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Tag             =   "1003"
   Begin VB.Frame Frame1 
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
      Left            =   5805
      TabIndex        =   30
      Top             =   4710
      Width           =   3090
      Begin GTMaskNum.GTMaskNum txtQuantidade 
         Height          =   330
         Left            =   1500
         TabIndex        =   31
         Top             =   360
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
         Left            =   150
         TabIndex        =   32
         Top             =   390
         Width           =   1050
      End
   End
   Begin VB.Frame fraRangeUtentes 
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
      Left            =   5805
      TabIndex        =   22
      Top             =   3705
      Width           =   3090
      Begin GTMaskNum.GTMaskNum txtNumUtenteIni 
         Height          =   330
         Left            =   540
         TabIndex        =   24
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
         Left            =   2025
         TabIndex        =   26
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
         Caption         =   "a"
         Height          =   255
         Index           =   1
         Left            =   1635
         TabIndex        =   25
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblLabel 
         Caption         =   "de"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   390
         Width           =   225
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
      TabIndex        =   20
      Top             =   4710
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboSalas 
         Height          =   330
         Left            =   165
         TabIndex        =   21
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
      TabIndex        =   18
      Top             =   3705
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
         Height          =   330
         Left            =   165
         TabIndex        =   19
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
      Left            =   10995
      Top             =   4800
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
      Width           =   11265
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Fichas de Utente (Alteração/Atualização)"
         Height          =   330
         Index           =   17
         Left            =   6000
         TabIndex        =   29
         Top             =   2970
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Contrato de Prestação de Serviços"
         Height          =   300
         Index           =   16
         Left            =   6000
         TabIndex        =   17
         Top             =   2640
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome , Data Nasc. e Nº S.N.S."
         Height          =   300
         Index           =   15
         Left            =   6000
         TabIndex        =   16
         Top             =   2310
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome , Data Nasc. e NISS"
         Height          =   300
         Index           =   14
         Left            =   6000
         TabIndex        =   15
         Top             =   1980
         Width           =   5205
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Almoços"
         Height          =   330
         Index           =   13
         Left            =   6000
         TabIndex        =   14
         Top             =   1650
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Presenças na Sala"
         Height          =   330
         Index           =   12
         Left            =   6000
         TabIndex        =   13
         Top             =   1320
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Fichas de Utente"
         Height          =   330
         Index           =   11
         Left            =   6000
         TabIndex        =   12
         Top             =   990
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Carta Dívida dos Utentes"
         Height          =   330
         Index           =   10
         Left            =   6000
         TabIndex        =   11
         Top             =   660
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Etiquetas dos Utentes"
         Height          =   330
         Index           =   9
         Left            =   6000
         TabIndex        =   10
         Top             =   330
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome, Encarregado de Educação"
         Height          =   300
         Index           =   8
         Left            =   330
         TabIndex        =   9
         Top             =   2970
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nome, Cédula Pessoal e B.I."
         Height          =   300
         Index           =   7
         Left            =   330
         TabIndex        =   8
         Top             =   2640
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Salas com Moradas"
         Height          =   300
         Index           =   6
         Left            =   330
         TabIndex        =   7
         Top             =   2310
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Salas com Mensalidade  e Próxima Mensalidade"
         Height          =   300
         Index           =   5
         Left            =   330
         TabIndex        =   6
         Top             =   1980
         Width           =   4800
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Nº de Utente para Seguros"
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   5
         Top             =   1650
         Width           =   3900
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Salas com Valor da Mensalidade"
         Height          =   300
         Index           =   3
         Left            =   330
         TabIndex        =   4
         Top             =   1320
         Width           =   3900
      End
      Begin VB.OptionButton optListaUtentes 
         Caption         =   "Salas"
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
      Left            =   10215
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "2003"
      Top             =   3810
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "2010"
      Top             =   3810
      Width           =   1200
   End
End
Attribute VB_Name = "frmListaUtentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSListaUtentes As Workspace
Dim mBDListaUtentes As Database
Dim mBDListaUtentesTemp As Database

Dim tBDAberta
Dim cSql
Dim cNomeMapa

Dim i As Integer

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSListaUtentes = DBEngine.CreateWorkspace("ListaUtentes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDListaUtentes = mWSListaUtentes.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDListaUtentesTemp = mWSListaUtentes.OpenDatabase(cBDComNomeUtilizador)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Lista de Utentes-Abrir BD", Err.Number, Err.Description)
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
    Dim cAno
    
    Set mProcessamento = New Processamento

    cCod_Inst = cCodificaInstituicao(cboInstituicao.Text)
    cCod_Sala = cCodificaSala(cCod_Inst, cboSalas.Text)
    cAno = str$(Year(Date))

On Error GoTo TrataErro

     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDListaUtentesTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDListaUtentes.Execute cSql, dbFailOnError
    
    ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDListaUtentesTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDListaUtentes.Execute cSql, dbFailOnError
    
    ' apaga os registos da Temp32
    If optListaUtentes(9).Value Then
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM LISTA_UTENTES_ETIQ;"
    ElseIf optListaUtentes(10).Value Then
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM LISTA_UTENTES_CARTA_DIV;"
    ElseIf optListaUtentes(11).Value Then
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM FICHA_UTENTES;"
    ElseIf optListaUtentes(17).Value Then
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM FICHA_UTENTES;"
    Else
        ' apaga os registos da Temp32
        cSql = "DELETE * FROM LISTA_UTENTES;"
    End If
    ' apaga o registo em Temp32
    mBDListaUtentesTemp.Execute cSql, dbFailOnError

    If optListaUtentes(9).Value Then
        cSql = "INSERT INTO LISTA_UTENTES_ETIQ IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE "
        cSql = cSql & "NUM_UTENTE BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    ElseIf optListaUtentes(10).Value Then
        cSql = "INSERT INTO LISTA_UTENTES_CARTA_DIV IN '" & cBDComNomeUtilizador
        cSql = cSql & "' SELECT RECIBOS.ANO,RECIBOS.COD_MES,RECIBOS.NUM_UTENTE,RECIBOS.COD_INST,RECIBOS.COD_SALA,"
        cSql = cSql & "RECIBOS.NOME,RECIBOS.TOTAL_MENSALIDADE,RECIBOS.NUM_RECIBO,RECIBOS.ESTADO_REC,"
        cSql = cSql & "UTENTES.NOME_ENC_EDU,UTENTES.MORADA,UTENTES.COD_POSTAL,UTENTES.LOCAL"
        cSql = cSql & " FROM UTENTES INNER JOIN RECIBOS ON UTENTES.NUM_UTENTE = RECIBOS.NUM_UTENTE"
        cSql = cSql & " WHERE (RECIBOS.COD_INST<>'999' AND RECIBOS.COD_SALA<>'999') AND (RECIBOS.ESTADO_REC='D' OR RECIBOS.ESTADO_REC='P')"
    ElseIf optListaUtentes(11).Value Then
        cSql = "INSERT INTO FICHA_UTENTES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE "
        cSql = cSql & "NUM_UTENTE BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    ElseIf optListaUtentes(16).Value Then
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_UTENTES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE "
        cSql = cSql & "NUM_UTENTE BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    ElseIf optListaUtentes(17).Value Then
        cSql = "INSERT INTO FICHA_UTENTES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE "
        cSql = cSql & "NUM_UTENTE BETWEEN " & txtNumUtenteIni.Value & _
                " AND " & txtNumUtenteFim.Value & " AND ISNULL(DATA_SAIDA)"
    Else
        ' carrega a variavela com o Sql
        cSql = "INSERT INTO LISTA_UTENTES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE ISNULL(DATA_SAIDA)"
    End If
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        If optListaUtentes(10).Value Then
            cSql = cSql & " AND RECIBOS.COD_INST='" & cCod_Inst & "'"
        Else
            cSql = cSql & " AND COD_INST='" & cCod_Inst & "'"
        End If
    End If
    If cboSalas.Text <> "<Todas as Salas>" Then
        If optListaUtentes(10).Value Then
            cSql = cSql & " AND RECIBOS.COD_SALA='" & cCod_Sala & "'"
        Else
            cSql = cSql & " AND COD_SALA='" & cCod_Sala & "'"
        End If
    End If
    
    If optListaUtentes(0).Value Then
        ' Lista de Utentes por Número
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ004.RPT"
    ElseIf optListaUtentes(1).Value Then
        ' Lista de Utentes por Nome
        cSql = cSql & " ORDER BY NOME ASC;"
        cNomeMapa = "CBESQ004.RPT"
    ElseIf optListaUtentes(2).Value Then
        ' Lista de Utentes por Salas
        cSql = cSql & ";"
        cNomeMapa = "CBESQ005.RPT"
    ElseIf optListaUtentes(3).Value Then
        ' Lista de Utentes por Salas com Valor das Mensalidades
        cSql = cSql & ";"
        cNomeMapa = "CBESQ006.RPT"
    ElseIf optListaUtentes(4).Value Then
        ' Lista de Utentes por Número para Seguros
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ019.RPT"
    ElseIf optListaUtentes(5).Value Then
        ' Lista de Utentes por Salas com Valor da Mensalidade e Próxima Mensalidade
        cSql = cSql & ";"
        cNomeMapa = "CBESQ023.RPT"
    ElseIf optListaUtentes(6).Value Then
        ' Lista de Utentes por Salas com Morada
        cSql = cSql & ";"
        cNomeMapa = "CBESQ026.RPT"
    ElseIf optListaUtentes(7).Value Then
        ' Lista de Utentes por Nome, Céd. e B.I
        cSql = cSql & ";"
        cNomeMapa = "CBESQ049.RPT"
    ElseIf optListaUtentes(8).Value Then
        ' Lista de Utentes por Nome, Nome Encarregado de Educação
        cSql = cSql & ";"
        cNomeMapa = "CBESQ052.RPT"
    ElseIf optListaUtentes(9).Value Then
        ' Lista de Etiquetas Utentes
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ054.RPT"
    ElseIf optListaUtentes(10).Value Then
        ' Carta Divida
        cSql = cSql & " ORDER BY RECIBOS.NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ053.RPT"
    ElseIf optListaUtentes(11).Value Then
        ' Ficha de Utente
        cSql = cSql & ";"
        cNomeMapa = "CBESQ003.RPT"
    ElseIf optListaUtentes(12).Value Then
        ' Lista de Utentes - Presenças
        cSql = cSql & ";"
        cNomeMapa = "CBESQ057.RPT"
    ElseIf optListaUtentes(13).Value Then
        ' Lista de Utentes - Almoços
        cSql = cSql & ";"
        cNomeMapa = "CBESQ058.RPT"
    ElseIf optListaUtentes(14).Value Then
        ' Lista de Nomes e Data Nasc e Nº Seg Social
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ059.RPT"
    ElseIf optListaUtentes(15).Value Then
        ' Lista de Nomes e Data Nasc e Nº Seg Social
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ060.RPT"
    ElseIf optListaUtentes(16).Value Then
        ' Contratos de Prestação de Serviços
        cSql = cSql & " ORDER BY NUM_UTENTE ASC;"
        cNomeMapa = "CBESQ061.RPT"
    ElseIf optListaUtentes(17).Value Then
        ' Ficha de Utente
        cSql = cSql & ";"
        cNomeMapa = "CBESQ003A.RPT"
    End If
    
    ' insere o registo em Temp32
    mBDListaUtentes.Execute cSql, dbFailOnError

    ' insere a Quantidade de Etiquetas
    If optListaUtentes(9).Value Then
        For i = 1 To CInt(txtQuantidade.Value) - 1
            ' insere o registo em Temp32
            mBDListaUtentes.Execute cSql, dbFailOnError
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
        If optListaUtentes(9).Value Then
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
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        
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
        ElseIf optListaUtentes(5).Value Then
            .Formulas(4) = "Descricao Parametros Mapa=''"
        ElseIf optListaUtentes(10).Value Then
            .Formulas(4) = "Mapa=''"
        ElseIf optListaUtentes(11).Value Then
            .Formulas(4) = "AnoLectivo='" & cAno & "/" & (cAno + 1) & "'"
        ElseIf optListaUtentes(17).Value Then
            .Formulas(4) = "AnoLectivo='" & cAno & "/" & (cAno + 1) & "'"
        End If
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Lista Utentes", Err.Number, Err.Description)
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
    
    txtNumUtenteIni.Enabled = False
    txtNumUtenteFim.Enabled = False
    txtQuantidade.Enabled = False
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSListaUtentes.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSListaUtentes = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmListaUtentes = Nothing
End Sub


Private Sub optListaUtentes_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 15
            txtNumUtenteIni.Enabled = False
            txtNumUtenteFim.Enabled = False
            txtQuantidade.Enabled = False
        Case 9
            txtNumUtenteIni.Enabled = True
            txtNumUtenteFim.Enabled = True
            txtQuantidade.Enabled = True
        Case 11, 16, 17
            txtNumUtenteIni.Enabled = True
            txtNumUtenteFim.Enabled = True
            txtQuantidade.Enabled = False
    End Select
    txtNumUtenteIni.Value = 1
    txtNumUtenteFim.Value = 9999
    txtQuantidade.Value = 1
End Sub


