VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFGestaoUtentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Gestao de Utentes"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
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
   ScaleHeight     =   6435
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Tag             =   "1002"
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Saida"
      Height          =   900
      Left            =   6105
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "2012"
      Top             =   5430
      Width           =   1200
   End
   Begin Crystal.CrystalReport rptRecibosMensalidade 
      Left            =   2100
      Top             =   5565
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "&Pagar Mensalidade"
      Height          =   900
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2011"
      Top             =   5430
      Visible         =   0   'False
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
      TabIndex        =   6
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
      Height          =   900
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "2003"
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Nova"
      Height          =   900
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "2007"
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   900
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "2008"
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdFicha 
      Caption         =   "&Ficha"
      Height          =   900
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2006"
      Top             =   5430
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
      TabIndex        =   1
      Top             =   390
      Width           =   3000
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoUtentes 
      Bindings        =   "CBESQ035.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   9
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
      TabIndex        =   3
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1320
      Width           =   3300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar os Utentes da Instituição"
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
      Width           =   3870
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
      TabIndex        =   0
      Top             =   120
      Width           =   2685
   End
End
Attribute VB_Name = "frmCAIFGestaoUtentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSUtentes_Idosos As Workspace
Dim mBDUtentes_Idosos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox
Dim cNomeMapa

Private Sub CarregaGridUtentes(ByRef Grid As SSDBGrid, ByVal iOrdem)
    Dim recTabelaUTENTES As Recordset
    Dim cLinha
    
    Grid.Redraw = False
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME,DATA_SAIDA FROM UTENTES_IDOSOS "
    
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
    Set recTabelaUTENTES = mBDUtentes_Idosos.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
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
    Set mWSUtentes_Idosos = DBEngine.CreateWorkspace("Utentes_Idosos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDUtentes_Idosos = mWSUtentes_Idosos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Gestão de Utentes-Abrir BD", Err.Number, Err.Description)
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

Private Sub cmdAlterar_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaUtenteCAIF = New frmCAIFFichaUtente
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Altera"
        ' faz o load da form
        fFichaUtenteCAIF.Show
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
        cSql = "SELECT RECIBOS_IDOSOS.ANO,RECIBOS_IDOSOS.COD_MES,TABMESES.NOME,TABINSTITUICAO.NOME,TABSALAS.NOME," & _
            "RECIBOS_IDOSOS.NOME,RECIBOS_IDOSOS.MENSALIDADE,RECIBOS_IDOSOS.TOTAL_MENSALIDADE,RECIBOS_IDOSOS.VALOR1," & _
            "RECIBOS_IDOSOS.VALOR2,RECIBOS_IDOSOS.VALOR3,RECIBOS_IDOSOS.VALOR4,RECIBOS_IDOSOS.VALOR5,RECIBOS_IDOSOS.NUM_RECIBO "
        cSql = cSql & "FROM ((RECIBOS_IDOSOS INNER JOIN TABMESES ON " & _
            "RECIBOS_IDOSOS.COD_MES = TABMESES.COD_MES) INNER JOIN TABINSTITUICAO ON " & _
            "RECIBOS_IDOSOS.COD_INST = TABINSTITUICAO.COD_INST) INNER JOIN TABSALAS ON " & _
            "(RECIBOS_IDOSOS.COD_SALA = TABSALAS.COD_SALA) AND (RECIBOS_IDOSOS.COD_INST = TABSALAS.COD_INST) "
        cSql = cSql & "WHERE (((RECIBOS_IDOSOS.NUM_UTENTE)=" & sgrdGestaoUtentes.Columns(0).Value & ") " & _
            "AND ((RECIBOS_IDOSOS.ESTADO_REC)='P' Or (RECIBOS_IDOSOS.ESTADO_REC)='D')) " & _
            "ORDER BY RECIBOS_IDOSOS.ANO ASC, RECIBOS_IDOSOS.COD_MES ASC;"
    
        ' insere o registo em Temp32
        Set recRECIBOS = mBDUtentes_Idosos.OpenRecordset(cSql, dbOpenSnapshot)
        
        If (recRECIBOS.EOF And recRECIBOS.BOF) Then
            MsgBox "O Utente Nº " & sgrdGestaoUtentes.Columns(0).Value & ", não tem " & vbCrLf & _
                "Mensalidades para pagar !", vbInformation + vbOKOnly, "Pagar Mensalidade"
            GoTo SairDoProcedimento
        Else
            cNomeMapa = "CBESQ031.RPT"
            
            recRECIBOS.MoveFirst

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
                .Formulas(10) = "Nome='" & recRECIBOS.Fields("RECIBOS_IDOSOS.NOME").Value & "'"
                .Formulas(11) = "Num Recibo=" & recRECIBOS.Fields("NUM_RECIBO").Value
                .Formulas(12) = "Mes='" & recRECIBOS.Fields("TABMESES.NOME").Value & "'"
                .Formulas(13) = "Ano='" & recRECIBOS.Fields("ANO").Value & "'"
                .Formulas(14) = "Mensalidade=" & cAlteraVparaP(CStr(recRECIBOS.Fields("MENSALIDADE").Value))
                .Formulas(15) = "Valor 1=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR1").Value))
                .Formulas(16) = "Valor 2=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR2").Value))
                .Formulas(17) = "Valor 3=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR3").Value))
                .Formulas(18) = "Valor 4=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR4").Value))
                .Formulas(19) = "Valor 5=" & cAlteraVparaP(CStr(recRECIBOS.Fields("VALOR5").Value))
                .Formulas(20) = "NIF='" & cNIF_CAIF(sgrdGestaoUtentes.Columns(0).Value) & "'"
                'executa o Report
                .Action = 1
            End With
        
            iRespMsgbox = MsgBox("Confirme se o Recibo foi" & vbCrLf & _
                                "impresso correctamente ?", vbQuestion + vbYesNo + vbDefaultButton2, _
                                "Pagar Mensalidade")
                                
            If iRespMsgbox = vbYes Then
                Set qryPagaRecibo = mBDUtentes_Idosos.QueryDefs("CAIF RECIBOS Pagar")
                ' parametros de input
                qryPagaRecibo.Parameters("Utiliz") = gUtilizador.Nome
                qryPagaRecibo.Parameters("Ano") = recRECIBOS.Fields("ANO")
                qryPagaRecibo.Parameters("Mes") = recRECIBOS.Fields("COD_MES")
                qryPagaRecibo.Parameters("Num_Utente") = sgrdGestaoUtentes.Columns(0).Value
                ' executa a query
                qryPagaRecibo.Execute dbFailOnError
            End If
        End If
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("CAIF - Pagar Mensalidade", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
        recRECIBOS.Close
        Set recRECIBOS = Nothing
    
    End If
End Sub

Private Sub cmdSaida_Click()
    Dim mProcessamento As Processamento
    Dim qrySaidaUtente As QueryDef
    Dim recRECIBOS As Recordset
    Dim cSql
    
    Set mProcessamento = New Processamento
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
        ' confirma se quer mesmo apagar
        iRespMsgbox = MsgBox("Confirme se o Utente nº " & sgrdGestaoUtentes.Columns(0).Value & "," & vbCrLf & _
                    "vai sair da Instituição ?", vbQuestion + vbYesNo, "Saida de Utente")
        If iRespMsgbox = vbYes Then
            If IsDate(sgrdGestaoUtentes.Columns(2).Value) Then
                MsgBox "Este Utente já tem Data de Saida registada !!!", vbOKOnly, "Saida de Utente"
                Exit Sub
            End If
On Error GoTo TrataErro
            mWSUtentes_Idosos.BeginTrans
            ' dar saida de utente
            Set qrySaidaUtente = mBDUtentes_Idosos.QueryDefs("CAIF UTENTES Data Saida")
    
            qrySaidaUtente.Parameters("Num_Utente") = CLng(sgrdGestaoUtentes.Columns(0).Value)
            qrySaidaUtente.Parameters("Utiliz") = gUtilizador.Nome

            ' executa a Saida
            qrySaidaUtente.Execute dbFailOnError
            
            ' altera os recibos do Utente
            Set qrySaidaUtente = mBDUtentes_Idosos.QueryDefs("CAIF UTENTES Data Saida RECIBOS")
    
            qrySaidaUtente.Parameters("Num_Utente") = CLng(sgrdGestaoUtentes.Columns(0).Value)
            qrySaidaUtente.Parameters("Utiliz") = gUtilizador.Nome

            ' executa alteração de Cod_Inst, Cod_Sala dos recibos do Utente
            qrySaidaUtente.Execute dbFailOnError
            
            
            mWSUtentes_Idosos.CommitTrans
On Error GoTo 0
            ' faz o refresh
            datUtentes.Refresh
            sgrdGestaoUtentes.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSUtentes_Idosos.Rollback
    Call ErrosGerais("CAIF - Saida de Utente", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFicha_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoUtentes.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaUtenteCAIF = New frmCAIFFichaUtente
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Ficha"
        ' faz o load da form
        fFichaUtenteCAIF.Show
    End If
End Sub

Private Sub cmdNovo_Click()
    ' cria nova instancia da Form
    Set fFichaUtenteCAIF = New frmCAIFFichaUtente
    ' Carrega a Variavel para passar o Botao de onde vai
    cBotaoOrigem = "Novo"
    ' faz o load da form
    fFichaUtenteCAIF.Show
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
    
    If gUtilizador.Perfil = "CONSULTA" Then
        cmdPagar.Visible = False
        cmdNovo.Visible = False
        cmdAlterar.Visible = False
        cmdSaida.Visible = False
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' fecha as tabelas abertas e as bases de dados abertas no Workspace
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSUtentes_Idosos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    ' limpa o workspace
    Set mWSUtentes_Idosos = Nothing
    ' apaga a janela da lista de janelas abertas
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    ' limpa a form de memoria
    Set frmCAIFGestaoUtentes = Nothing
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


