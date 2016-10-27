VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGestaoRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestao de Recibos"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
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
   ScaleHeight     =   6720
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Tag             =   "1002"
   Begin VB.Frame frmTipoPag 
      Caption         =   " Tipo de Pagamento "
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
      Height          =   1200
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   2550
      Begin VB.OptionButton optTipoPagamento 
         Caption         =   "Dinheiro / Cheque"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   12
         Top             =   330
         Value           =   -1  'True
         Width           =   2160
      End
      Begin VB.OptionButton optTipoPagamento 
         Caption         =   "Transf. Bancária"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   13
         Top             =   690
         Width           =   2160
      End
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "&Visualizar Mensalidade"
      Height          =   1050
      Left            =   4875
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   2100
   End
   Begin VB.CommandButton cmdNotasDebito 
      Caption         =   "&Notas de Débito"
      Height          =   420
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   2940
   End
   Begin Crystal.CrystalReport rptRecibosMensalidade 
      Left            =   0
      Top             =   5430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "&Pagar Mensalidade"
      Height          =   1050
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "2011"
      Top             =   5520
      Width           =   2100
   End
   Begin VB.Frame fraOrdenar 
      Caption         =   " Ordenar por "
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
      Height          =   1200
      Left            =   5505
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nome do Utente"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   8
         Top             =   690
         Width           =   2325
      End
      Begin VB.OptionButton optOrdenacao 
         Caption         =   " Nº de Utente"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   1050
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "2003"
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Data datUtentes 
      Caption         =   "datUtentes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtProcura 
      Height          =   330
      Left            =   120
      MaxLength       =   25
      TabIndex        =   4
      Top             =   390
      Width           =   3000
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoUtentes 
      Bindings        =   "CBESQ060.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   8400
      _Version        =   196617
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   14817
      _ExtentY        =   5821
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
      Height          =   330
      Left            =   135
      TabIndex        =   5
      Top             =   990
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
   Begin SSDataWidgets_B.SSDBCombo cboSalas 
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   1590
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
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar os Utentes da Sala"
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
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar os Utentes do Equipamento"
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
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4185
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por Nº de Utente"
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2685
   End
End
Attribute VB_Name = "frmGestaoRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSRecibos As Workspace
Dim mBDRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox
Dim cNomeMapa

Private Sub CarregaGridUtentes(ByRef Grid As SSDBGrid, ByVal iOrdem)
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    
    Grid.Redraw = False
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME,DATA_SAIDA FROM UTENTES "
    ' se seleccionou Instituição tem de filtrar
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " WHERE COD_INST='" & cCodificaInstituicao(cboInstituicao.Text) & "'"
        ' se seleccionou Sala tem de filtrar
        If cboSalas.Text <> "<Todas as Salas>" Then
            cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) & "'"
        End If
    Else
        cSql = cSql & " WHERE COD_INST<>'999'"
        cSql = cSql & " AND COD_SALA<>'999'"
    End If
    
    ' Estabelace a ordem dos registos
    Select Case iOrdem
        Case 0
            cSql = cSql & " ORDER BY NUM_UTENTE ASC"
        Case 1
            cSql = cSql & " ORDER BY NOME ASC"
    End Select
    ' abre a tabela
    Set recTabelaUTENTES = mBDRecibos.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' carrega o data control da grid
    Set datUtentes.Recordset = recTabelaUTENTES
    
SairDoProcedimento:
    Grid.Redraw = True
End Sub

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
On Error GoTo TrataErro
    Set mWSRecibos = DBEngine.CreateWorkspace("Recibos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDRecibos = mWSRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Gestão de Recibos-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    ' Carrega a grid com nova Ordenação
    txtProcura.Text = vbNullString
    cboSalas.Text = "<Todas as Salas>"
    Call CarregaGridUtentes(sgrdGestaoUtentes, IIf(optOrdenacao(0).Value, 0, 1))
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


Private Sub cboSalas_Click()
    ' Carrega a grid com nova Ordenação
    txtProcura.Text = vbNullString
    Call CarregaGridUtentes(sgrdGestaoUtentes, IIf(optOrdenacao(0).Value, 0, 1))
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


Private Sub cmdNotasDebito_Click()
    Dim mProcessamento As Processamento
    Dim recRECIBOS As Recordset
    Dim qryPagaRecibo As QueryDef
    
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
    
        Set mProcessamento = New Processamento

On Error GoTo TrataErro
        ' carrega a variavel com o Sql
        cSql = "SELECT RECIBOS.ANO,RECIBOS.COD_MES,TABMESES.NOME,RECIBOS.COD_INST,TABINSTITUICAO.NOME,RECIBOS.COD_SALA,TABSALAS.NOME," & _
            "RECIBOS.NOME,RECIBOS.MENSALIDADE,RECIBOS.TOTAL_MENSALIDADE,RECIBOS.VALOR1," & _
            "RECIBOS.VALOR2,RECIBOS.VALOR3,RECIBOS.VALOR4,RECIBOS.VALOR5,RECIBOS.NUM_RECIBO,RECIBOS.MENSALIDADE_PCTG "
        cSql = cSql & "FROM ((RECIBOS INNER JOIN TABMESES ON " & _
            "RECIBOS.COD_MES = TABMESES.COD_MES) INNER JOIN TABINSTITUICAO ON " & _
            "RECIBOS.COD_INST = TABINSTITUICAO.COD_INST) INNER JOIN TABSALAS ON " & _
            "(RECIBOS.COD_SALA = TABSALAS.COD_SALA) AND (RECIBOS.COD_INST = TABSALAS.COD_INST) "
        cSql = cSql & "WHERE (((RECIBOS.NUM_UTENTE)=" & sgrdGestaoUtentes.Columns(0).Value & ") " & _
            "AND ((RECIBOS.ESTADO_REC)='P' Or (RECIBOS.ESTADO_REC)='D')) " & _
            "AND (((RECIBOS.ANO)='2003' AND (RECIBOS.COD_MES)>='11') OR (RECIBOS.ANO)>'2003') " & _
            "ORDER BY RECIBOS.ANO ASC, RECIBOS.COD_MES ASC;"
    
        ' insere o registo em Temp32
        Set recRECIBOS = mBDRecibos.OpenRecordset(cSql, dbOpenSnapshot)
        
        If (recRECIBOS.EOF And recRECIBOS.BOF) Then
            MsgBox "O Utente Nº " & sgrdGestaoUtentes.Columns(0).Value & ", não tem " & vbCrLf & _
                "Notas de Débito !", vbInformation + vbOKOnly, "Notas de Débito"
            GoTo SairDoProcedimento
        Else
            cNomeMapa = "CBESQ007ND.RPT"
            
            recRECIBOS.MoveFirst

            Do While Not recRECIBOS.EOF
                With rptRecibosMensalidade
                    'Carrega o Nome do Report se ele existir
                    If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
                        .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
                    Else
                        MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
                        GoTo SairDoProcedimento
                    End If
                    .WindowParentHandle = fFrmMDIPrincipal.hwnd
                    'Configura o destino e o numero de copias e de linhas para o Mapa
'                    .Destination = crptToWindow
                    .Destination = crptToPrinter
                    .CopiesToPrinter = 1
                    'Passa para o Mapa os dados da Empresa
                    .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
                    .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
                    .Formulas(2) = "Titulo_3='" & gEmpresa.Linha4 & "'"
                    .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                    .Formulas(4) = "Morada='" & gEmpresa.Linha1 & "'"
                    .Formulas(5) = "Codigo Postal='" & gEmpresa.Linha2 & "'"
                    .Formulas(6) = "Telefone='" & gEmpresa.Linha3 & "'"
                    .Formulas(7) = "Instituicao='" & recRECIBOS.Fields("TABINSTITUICAO.NOME").Value & "'"
                    .Formulas(8) = "Sala='" & recRECIBOS.Fields("TABSALAS.NOME").Value & "'"
                    .Formulas(9) = "Num Utente=" & sgrdGestaoUtentes.Columns(0).Value
                    .Formulas(10) = "Nome='" & recRECIBOS.Fields("RECIBOS.NOME").Value & "'"
                    .Formulas(11) = "Num Recibo=" & recRECIBOS.Fields("NUM_RECIBO").Value
                    .Formulas(12) = "Mes='" & recRECIBOS.Fields("TABMESES.NOME").Value & "'"
                    .Formulas(13) = "Ano='" & recRECIBOS.Fields("ANO").Value & "'"
                    .Formulas(14) = "Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE").Value))
                    .Formulas(15) = "Valor 1=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR1").Value))
                    .Formulas(16) = "Valor 2=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR2").Value))
                    .Formulas(17) = "Valor 3=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR3").Value))
                    .Formulas(18) = "Valor 4=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR4").Value))
                    .Formulas(19) = "Valor 5=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR5").Value))
                    .Formulas(20) = "Percentagem=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE_PCTG").Value))
                    .Formulas(21) = "COD_INST='" & recRECIBOS.Fields("COD_INST").Value & "'"
                    .Formulas(22) = "COD_SALA='" & recRECIBOS.Fields("COD_SALA").Value & "'"
                    'executa o Report
                    .Action = 1
                End With
                recRECIBOS.MoveNext
            Loop
        End If
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Pagar Mensalidade", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
        recRECIBOS.Close
        Set recRECIBOS = Nothing
    
    End If
End Sub

Private Sub cmdPagar_Click()
    Dim mProcessamento As Processamento
    Dim recRECIBOS As Recordset
    Dim qryPagaRecibo As QueryDef
    
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
    
        Set mProcessamento = New Processamento

On Error GoTo TrataErro
        ' carrega a variavel com o Sql
        cSql = "SELECT RECIBOS.ANO,RECIBOS.COD_MES,TABMESES.NOME,RECIBOS.COD_INST,TABINSTITUICAO.NOME,RECIBOS.COD_SALA,TABSALAS.NOME," & _
            "RECIBOS.NOME,RECIBOS.MENSALIDADE,RECIBOS.TOTAL_MENSALIDADE,RECIBOS.VALOR1," & _
            "RECIBOS.VALOR2,RECIBOS.VALOR3,RECIBOS.VALOR4,RECIBOS.VALOR5,RECIBOS.NUM_RECIBO,RECIBOS.MENSALIDADE_PCTG,RECIBOS.TOTAL_MENSALIDADE,TABSALAS.VALENCIA "
        cSql = cSql & "FROM ((RECIBOS INNER JOIN TABMESES ON " & _
            "RECIBOS.COD_MES = TABMESES.COD_MES) INNER JOIN TABINSTITUICAO ON " & _
            "RECIBOS.COD_INST = TABINSTITUICAO.COD_INST) INNER JOIN TABSALAS ON " & _
            "(RECIBOS.COD_SALA = TABSALAS.COD_SALA) AND (RECIBOS.COD_INST = TABSALAS.COD_INST) "
        cSql = cSql & "WHERE (((RECIBOS.NUM_UTENTE)=" & sgrdGestaoUtentes.Columns(0).Value & ") " & _
            "AND ((RECIBOS.ESTADO_REC)='P' Or (RECIBOS.ESTADO_REC)='D')) " & _
            "ORDER BY RECIBOS.ANO ASC, RECIBOS.COD_MES ASC, RECIBOS.NUM_RECIBO ASC;"
    
        ' insere o registo em Temp32
        Set recRECIBOS = mBDRecibos.OpenRecordset(cSql, dbOpenSnapshot)
        
        If (recRECIBOS.EOF And recRECIBOS.BOF) Then
            MsgBox "O Utente Nº " & sgrdGestaoUtentes.Columns(0).Value & ", não tem " & vbCrLf & _
                "Mensalidades para pagar !", vbInformation + vbOKOnly, "Pagar Mensalidade"
            GoTo SairDoProcedimento
        Else
            
            recRECIBOS.MoveFirst
            
            If recRECIBOS.Fields("COD_MES").Value = "07" Then
                If (recRECIBOS.Fields("COD_INST").Value = "001" And recRECIBOS.Fields("COD_SALA").Value = "008") Or _
                (recRECIBOS.Fields("COD_INST").Value = "002" And recRECIBOS.Fields("COD_SALA").Value = "006") Then
                    cNomeMapa = "CBESQ007.RPT"
                Else
                    cNomeMapa = "CBESQ007RM.RPT"
                End If
            ElseIf recRECIBOS.Fields("COD_MES").Value = "08" Then
                If (recRECIBOS.Fields("COD_INST").Value = "001" And recRECIBOS.Fields("COD_SALA").Value = "008") Or _
                (recRECIBOS.Fields("COD_INST").Value = "002" And recRECIBOS.Fields("COD_SALA").Value = "006") Then
                    cNomeMapa = "CBESQ007RM.RPT"
                Else
                    cNomeMapa = "CBESQ007.RPT"
                End If
            Else
                cNomeMapa = "CBESQ007.RPT"
            End If
            
            With rptRecibosMensalidade
                'Carrega o Nome do Report se ele existir
                If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
                    .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
                Else
                    MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
                    GoTo SairDoProcedimento
                End If
                .WindowParentHandle = fFrmMDIPrincipal.hwnd
                'Configura o destino e o numero de copias e de linhas para o Mapa
'                .Destination = crptToWindow
                .Destination = crptToPrinter
               .CopiesToPrinter = 1
                'Passa para o Mapa os dados da Empresa
                .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
                .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
                .Formulas(2) = "Titulo_3='" & gEmpresa.Linha4 & "'"
                .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                .Formulas(4) = "Morada='" & gEmpresa.Linha1 & "'"
                .Formulas(5) = "Codigo Postal='" & gEmpresa.Linha2 & "'"
                .Formulas(6) = "Telefone='" & gEmpresa.Linha3 & "'"
                .Formulas(7) = "Instituicao='" & recRECIBOS.Fields("TABINSTITUICAO.NOME").Value & "'"
                .Formulas(8) = "Sala='" & recRECIBOS.Fields("TABSALAS.NOME").Value & "'"
                .Formulas(9) = "Num Utente=" & sgrdGestaoUtentes.Columns(0).Value
                .Formulas(10) = "Nome='" & recRECIBOS.Fields("RECIBOS.NOME").Value & "'"
                .Formulas(11) = "Num Recibo=" & recRECIBOS.Fields("NUM_RECIBO").Value
                .Formulas(12) = "Mes='" & recRECIBOS.Fields("TABMESES.NOME").Value & "'"
                .Formulas(13) = "Ano='" & recRECIBOS.Fields("ANO").Value & "'"
                .Formulas(14) = "Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE").Value))
                .Formulas(15) = "Valor 1=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR1").Value))
                .Formulas(16) = "Valor 2=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR2").Value))
                .Formulas(17) = "Valor 3=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR3").Value))
                .Formulas(18) = "Valor 4=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR4").Value))
                .Formulas(19) = "Valor 5=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR5").Value))
                .Formulas(20) = "Percentagem=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE_PCTG").Value))
                .Formulas(21) = "COD_INST='" & recRECIBOS.Fields("COD_INST").Value & "'"
                .Formulas(22) = "COD_SALA='" & recRECIBOS.Fields("COD_SALA").Value & "'"
                .Formulas(23) = "Total_Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("TOTAL_MENSALIDADE").Value))
                .Formulas(24) = "NIF='" & cNIF(sgrdGestaoUtentes.Columns(0).Value) & "'"
                .Formulas(25) = "RespSocial='" & recRECIBOS.Fields("VALENCIA").Value & "'"
                'executa o Report
                .Action = 1
            End With
        
            iRespMsgBox = MsgBox("Confirme se o Recibo foi" & vbCrLf & _
                                "impresso correctamente ?", vbQuestion + vbYesNo + vbDefaultButton2, _
                                "Pagar Mensalidade")
                                
            If iRespMsgBox = vbYes Then
                Set qryPagaRecibo = mBDRecibos.QueryDefs("RECIBOS Pagar")
                ' parametros de input
                qryPagaRecibo.Parameters("Utiliz") = gUtilizador.Nome
                qryPagaRecibo.Parameters("Ano") = recRECIBOS.Fields("ANO")
                qryPagaRecibo.Parameters("Mes") = recRECIBOS.Fields("COD_MES")
                qryPagaRecibo.Parameters("Num_Utente") = sgrdGestaoUtentes.Columns(0).Value
                If optTipoPagamento(0).Value Then
                    qryPagaRecibo.Parameters("TIPO_PAG") = "D"
                ElseIf optTipoPagamento(1).Value Then
                    qryPagaRecibo.Parameters("TIPO_PAG") = "T"
                End If
                qryPagaRecibo.Parameters("NUM_RECIBO") = recRECIBOS.Fields("NUM_RECIBO")
                
                ' executa a query
                qryPagaRecibo.Execute dbFailOnError
            End If
        End If
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Pagar Mensalidade", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
        recRECIBOS.Close
        Set recRECIBOS = Nothing
    
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub




Private Sub cmdVisualizar_Click()
    Dim mProcessamento As Processamento
    Dim recRECIBOS As Recordset
    Dim qryPagaRecibo As QueryDef
    
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
    
        Set mProcessamento = New Processamento

On Error GoTo TrataErro
        ' carrega a variavel com o Sql
        cSql = "SELECT RECIBOS.ANO,RECIBOS.COD_MES,TABMESES.NOME,RECIBOS.COD_INST,TABINSTITUICAO.NOME,RECIBOS.COD_SALA,TABSALAS.NOME," & _
            "RECIBOS.NOME,RECIBOS.MENSALIDADE,RECIBOS.TOTAL_MENSALIDADE,RECIBOS.VALOR1," & _
            "RECIBOS.VALOR2,RECIBOS.VALOR3,RECIBOS.VALOR4,RECIBOS.VALOR5,RECIBOS.NUM_RECIBO,RECIBOS.MENSALIDADE_PCTG,RECIBOS.TOTAL_MENSALIDADE,TABSALAS.VALENCIA "
        cSql = cSql & "FROM ((RECIBOS INNER JOIN TABMESES ON " & _
            "RECIBOS.COD_MES = TABMESES.COD_MES) INNER JOIN TABINSTITUICAO ON " & _
            "RECIBOS.COD_INST = TABINSTITUICAO.COD_INST) INNER JOIN TABSALAS ON " & _
            "(RECIBOS.COD_SALA = TABSALAS.COD_SALA) AND (RECIBOS.COD_INST = TABSALAS.COD_INST) "
        cSql = cSql & "WHERE (((RECIBOS.NUM_UTENTE)=" & sgrdGestaoUtentes.Columns(0).Value & ") " & _
            "AND ((RECIBOS.ESTADO_REC)='P' Or (RECIBOS.ESTADO_REC)='D')) " & _
            "ORDER BY RECIBOS.ANO ASC, RECIBOS.COD_MES ASC, RECIBOS.NUM_RECIBO ASC;"
    
        ' insere o registo em Temp32
        Set recRECIBOS = mBDRecibos.OpenRecordset(cSql, dbOpenSnapshot)
        
        If (recRECIBOS.EOF And recRECIBOS.BOF) Then
            MsgBox "O Utente Nº " & sgrdGestaoUtentes.Columns(0).Value & ", não tem " & vbCrLf & _
                "Mensalidades para pagar !", vbInformation + vbOKOnly, "Pagar Mensalidade"
            GoTo SairDoProcedimento
        Else
            
            recRECIBOS.MoveFirst
            
            If recRECIBOS.Fields("COD_MES").Value = "07" Then
                If (recRECIBOS.Fields("COD_INST").Value = "001" And recRECIBOS.Fields("COD_SALA").Value = "008") Or _
                (recRECIBOS.Fields("COD_INST").Value = "002" And recRECIBOS.Fields("COD_SALA").Value = "006") Then
                    cNomeMapa = "CBESQ007.RPT"
                Else
                    cNomeMapa = "CBESQ007RM.RPT"
                End If
            ElseIf recRECIBOS.Fields("COD_MES").Value = "08" Then
                If (recRECIBOS.Fields("COD_INST").Value = "001" And recRECIBOS.Fields("COD_SALA").Value = "008") Or _
                (recRECIBOS.Fields("COD_INST").Value = "002" And recRECIBOS.Fields("COD_SALA").Value = "006") Then
                    cNomeMapa = "CBESQ007RM.RPT"
                Else
                    cNomeMapa = "CBESQ007.RPT"
                End If
            Else
                cNomeMapa = "CBESQ007.RPT"
            End If
            
            With rptRecibosMensalidade
                'Carrega o Nome do Report se ele existir
                If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
                    .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
                Else
                    MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
                    GoTo SairDoProcedimento
                End If
                .WindowParentHandle = fFrmMDIPrincipal.hwnd
                .WindowTitle = "Visualizar Recibo"
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
                .WindowShowNavigationCtls = False
                .WindowShowPrintBtn = False
                .WindowShowPrintSetupBtn = False
                .WindowShowProgressCtls = False
                .WindowShowZoomCtl = True
                .WindowShowSearchBtn = False
                .WindowShowRefreshBtn = False
                
                'Configura o destino e o numero de copias e de linhas para o Mapa
                .Destination = crptToWindow
'                .Destination = crptToPrinter
                .CopiesToPrinter = 1
                'Passa para o Mapa os dados da Empresa
                .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
                .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
                .Formulas(2) = "Titulo_3='" & gEmpresa.Linha4 & "'"
                .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                .Formulas(4) = "Morada='" & gEmpresa.Linha1 & "'"
                .Formulas(5) = "Codigo Postal='" & gEmpresa.Linha2 & "'"
                .Formulas(6) = "Telefone='" & gEmpresa.Linha3 & "'"
                .Formulas(7) = "Instituicao='" & recRECIBOS.Fields("TABINSTITUICAO.NOME").Value & "'"
                .Formulas(8) = "Sala='" & recRECIBOS.Fields("TABSALAS.NOME").Value & "'"
                .Formulas(9) = "Num Utente=" & sgrdGestaoUtentes.Columns(0).Value
                .Formulas(10) = "Nome='" & recRECIBOS.Fields("RECIBOS.NOME").Value & "'"
                .Formulas(11) = "Num Recibo=" & recRECIBOS.Fields("NUM_RECIBO").Value
                .Formulas(12) = "Mes='" & recRECIBOS.Fields("TABMESES.NOME").Value & "'"
                .Formulas(13) = "Ano='" & recRECIBOS.Fields("ANO").Value & "'"
                .Formulas(14) = "Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE").Value))
                .Formulas(15) = "Valor 1=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR1").Value))
                .Formulas(16) = "Valor 2=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR2").Value))
                .Formulas(17) = "Valor 3=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR3").Value))
                .Formulas(18) = "Valor 4=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR4").Value))
                .Formulas(19) = "Valor 5=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR5").Value))
                .Formulas(20) = "Percentagem=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE_PCTG").Value))
                .Formulas(21) = "COD_INST='" & recRECIBOS.Fields("COD_INST").Value & "'"
                .Formulas(22) = "COD_SALA='" & recRECIBOS.Fields("COD_SALA").Value & "'"
                .Formulas(23) = "Total_Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("TOTAL_MENSALIDADE").Value))
                .Formulas(24) = "NIF='" & cNIF(sgrdGestaoUtentes.Columns(0).Value) & "'"
                .Formulas(25) = "RespSocial='" & recRECIBOS.Fields("VALENCIA").Value & "'"
                'executa o Report
                .Action = 1
            End With
        End If
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Pagar Mensalidade", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
        recRECIBOS.Close
        Set recRECIBOS = Nothing
    
    End If
End Sub

Private Sub Form_Activate()
    ' marca a janela activa na lista de janelas abertas
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    ' centra a janela
    CenterMe Me
    ' le as imagens da form
    LoadResStrings Me
    ' adiciona na lista de janelas abertas
    Call AlteraWindowList(Me.Caption)
    ' poe a textbox só a aceitar numeros
    Call SetNumber(txtProcura, True)
    ' mostra a janela
    Me.Show
    DoEvents
    ' abre a base de dados
    tBDAberta = tAbreBD
    ' carrega a grid com os dados seleccionados
    Call CarregaGridUtentes(sgrdGestaoUtentes, 0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' fecha as tabelas abertas e as bases de dados abertas no Workspace
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    ' limpa o workspace
    Set mWSRecibos = Nothing
    ' apaga a janela da lista de janelas abertas
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    ' limpa a form de memoria
    Set frmGestaoRecibos = Nothing
End Sub


Private Sub optOrdenacao_Click(Index As Integer)
    If Index = 0 Then
        lblTexto(0).Caption = "Procurar por Nº de Utente"
        Call SetNumber(txtProcura, True)
        txtProcura.Text = vbNullString
    Else
        lblTexto(0).Caption = "Procurar por Nome do Utente"
        Call SetNumber(txtProcura, False)
        txtProcura.Text = vbNullString
    End If
    ' Carrega a grid com nova Ordenação
    Call CarregaGridUtentes(sgrdGestaoUtentes, Index)
End Sub


Private Sub sgrdGestaoUtentes_InitColumnProps()
    With sgrdGestaoUtentes
        If .StyleSets.Count = 0 Then
            .StyleSets.Add "Cabecalho"
            .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
            .StyleSets("Cabecalho").ForeColor = vbTitleBarText
            .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
            .StyleSets("Cabecalho").Font.Size = 10
            .StyleSets("Cabecalho").Font.Bold = True
        End If
        
        .AllowAddNew = False
        .AllowColumnMoving = ssRelocateAnywhere
        .AllowColumnShrinking = False
        .AllowColumnSizing = False
        .AllowColumnSwapping = ssRelocateAnywhere
        .AllowDelete = False
        .AllowDragDrop = False
        .AllowGroupMoving = False
        .AllowGroupShrinking = False
        .AllowGroupSizing = False
        .AllowRowSizing = False
        .AllowUpdate = False
        .BackColorOdd = dCorAmarelo
        .Caption = "Lista de Utentes"
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10
        .ForeColorEven = &H0&
        .FieldSeparator = vbTab
        .HeadFont.Name = "MS Sans Serif"
        .HeadFont.Size = 10
        .HeadFont.Bold = True
        .RowSelectionStyle = ssRowSelectionStyle3D
        .ScrollBars = ssScrollBarsVertical
        .SelectByCell = False
        .SelectTypeCol = ssSelectionTypeNone
        .SelectTypeRow = ssSelectionTypeSingleSelect
       
        'Nº Utente
        .Columns(0).Alignment = ssCaptionAlignmentRight
        .Columns(0).Caption = "NºUtente"
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 2 'Integer
        .Columns(0).Width = 1200
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).NumberFormat = "#####0"
        .Columns(0).Visible = True
        
        'Nome do Utente
        .Columns(1).Alignment = ssCaptionAlignmentLeft
        .Columns(1).Caption = "Nome do Utente"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 8 ' Text
        .Columns(1).Width = 6600
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
        
        'Data de Saida
        .Columns(2).DataType = 7 'date
        .Columns(2).Visible = False
    End With
End Sub


Private Sub txtProcura_Change()
    If Trim$(txtProcura.Text) <> vbNullString Then
        If optOrdenacao(0).Value Then
            datUtentes.Recordset.FindFirst "NUM_UTENTE=" & CInt(txtProcura.Text)
        Else
            datUtentes.Recordset.FindFirst "NOME LIKE '" & txtProcura.Text & "*'"
        End If
    End If
End Sub


