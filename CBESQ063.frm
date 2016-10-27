VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAlterarRecibosCriados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestao de Recibos de Utentes (Fechados)"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   ScaleHeight     =   10185
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin VB.Data datRecibos 
      Caption         =   "datRecibos"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5115
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4935
      Visible         =   0   'False
      Width           =   3315
   End
   Begin SSDataWidgets_B.SSDBCombo cboNumUtente 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   1590
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2003"
      Top             =   9045
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "2002"
      Top             =   9045
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdAlteracoes 
      Height          =   2940
      Left            =   120
      TabIndex        =   2
      Top             =   5835
      Width           =   8550
      ScrollBars      =   0
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   0
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   15081
      _ExtentY        =   5186
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
   Begin SSDataWidgets_B.SSDBCombo cboSalas 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   990
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
   Begin SSDataWidgets_B.SSDBCombo cboNome 
      Height          =   330
      Left            =   1605
      TabIndex        =   9
      Top             =   1590
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
   Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
      Height          =   330
      Left            =   120
      TabIndex        =   10
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
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoRecibos 
      Bindings        =   "CBESQ063.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   12
      Top             =   2055
      Width           =   8550
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
      _ExtentX        =   15081
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
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Recibo do Mês de "
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
      TabIndex        =   11
      Top             =   5520
      Width           =   1980
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
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3870
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3300
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Utente"
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
      Left            =   1590
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nº de Utente"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmAlterarRecibosCriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSAlteraRecibosCriados As Workspace
Dim mBDAlteraRecibosCriados As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox

Private Sub CarregaGridRecibos(ByRef Grid As SSDBGrid, ByVal iOrdem)
    Dim recTabelaRECIBOS As Recordset
    Dim cLinha
    
    Grid.Redraw = False
    ' começa a construir o Sql
    cSql = "SELECT NUM_UTENTE,NOME,NUM_RECIBO FROM RECIBOS WHERE (ESTADO_REC='P' OR ESTADO_REC='D')"
    
    ' se seleccionou Instituição tem de filtrar
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cboInstituicao.Text) & "'"
        ' se seleccionou Sala tem de filtrar
        If cboSalas.Text <> "<Todas as Salas>" Then
            cSql = cSql & " AND COD_SALA='" & cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) & "'"
        End If
    Else
        cSql = cSql & " AND COD_INST<>'999'"
        cSql = cSql & " AND COD_SALA<>'999'"
    End If
    
    ' se seleccionou Utente
    If cboNumUtente.Text <> vbNullString Then
        cSql = cSql & " AND NUM_UTENTE=" & cboNumUtente.Text
    End If
    
    ' Estabelace a ordem dos registos
    cSql = cSql & " ORDER BY NUM_UTENTE ASC, NUM_RECIBO ASC"
    
    
    ' abre a tabela
    Set recTabelaRECIBOS = mBDAlteraRecibosCriados.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' carrega o data control da grid
    Set datRecibos.Recordset = recTabelaRECIBOS
    
SairDoProcedimento:
    Grid.Redraw = True
End Sub


'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSAlteraRecibosCriados = DBEngine.CreateWorkspace("Alteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDAlteraRecibosCriados = mWSAlteraRecibosCriados.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Alteração de Recibos -Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    ' Carrega a grid com nova Ordenação
    cboSalas.Text = "<Todas as Salas>"
    cboNumUtente.Text = vbNullString
    cboNome.Text = vbNullString
    sgrdAlteracoes.Redraw = False
    sgrdAlteracoes.RemoveAll
    sgrdAlteracoes.Redraw = True
    
    
    ' Carrega a grid com nova Ordenação
    Call CarregaGridRecibos(sgrdGestaoRecibos, 0)
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
    
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

Private Sub cboNome_Click()
    ' poe o num de utente na combo
    cboNumUtente.Text = cboNome.Columns(1).Value
    ' Carrega a grid com nova Ordenação
    Call CarregaGridRecibos(sgrdGestaoRecibos, 0)
    ' carrega a grid
    Call CarregaGridAlteracoes(sgrdAlteracoes)
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
        ' Carrega a grid com nova Ordenação
        Call CarregaGridRecibos(sgrdGestaoRecibos, 0)
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
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


Private Sub cboSalas_Click()
    cboNumUtente.Text = vbNullString
    cboNome.Text = vbNullString
    sgrdAlteracoes.Redraw = False
    sgrdAlteracoes.RemoveAll
    sgrdAlteracoes.Redraw = True
    
    ' Carrega a grid com nova Ordenação
    Call CarregaGridRecibos(sgrdGestaoRecibos, 0)
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
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

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryAlteraRecibo As QueryDef

'    If Trim$(cboNumUtente.Text) <> vbNullString Then

        Set mProcessamento = New Processamento
        
        ' pede confirmação se quer continuar
        iRespMsgBox = MsgBox("Confirma que quer Alterar os Dados do Recibo.", _
                            vbQuestion + vbYesNo, "Alterar Recibos")
        ' se resposta não sai
        If iRespMsgBox = vbNo Then
            GoTo SairDoProcedimento
        End If
On Error GoTo TrataErro
        mWSAlteraRecibosCriados.BeginTrans
    
        Set qryAlteraRecibo = mBDAlteraRecibosCriados.QueryDefs("RECIBOS Altera Recibo Criado")
        
        sgrdAlteracoes.Row = 5
        qryAlteraRecibo.Parameters("Mensalidade") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 6
        qryAlteraRecibo.Parameters("Mensalidade_PCTG") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 0
        qryAlteraRecibo.Parameters("Valor1") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 1
        qryAlteraRecibo.Parameters("Valor2") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 2
        qryAlteraRecibo.Parameters("Valor3") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 3
        qryAlteraRecibo.Parameters("Valor4") = sgrdAlteracoes.Columns(1).Value
        sgrdAlteracoes.Row = 4
        qryAlteraRecibo.Parameters("Valor5") = sgrdAlteracoes.Columns(1).Value
        qryAlteraRecibo.Parameters("Utiliz") = gUtilizador.Nome
        qryAlteraRecibo.Parameters("Num_Recibo") = sgrdGestaoRecibos.Columns(2).Value
    
        ' executa a inserção
        qryAlteraRecibo.Execute dbFailOnError
        mWSAlteraRecibosCriados.CommitTrans
        
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
        
        MsgBox "Alterou os dados relativos ao Utente.", vbInformation + vbOKOnly, "Alterar Recibos"
    
'    End If
        
    GoTo SairDoProcedimento
    
TrataErro:
    mWSAlteraRecibosCriados.Rollback
    Call ErrosGerais("Altera Recibos", Err.Number, Err.Description)
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
    
    ' carrega a grid com os dados seleccionados
    Call CarregaGridRecibos(sgrdGestaoRecibos, 0)
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSAlteraRecibosCriados.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSAlteraRecibosCriados = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFAlterarRecibos = Nothing
End Sub


Private Sub sgrdAlteracoes_BeforeRowColChange(Cancel As Integer)
    If IsEmpty(sgrdAlteracoes.Columns(1).Value) Then
        sgrdAlteracoes.Columns(1).Value = 0
    End If
End Sub

Private Sub sgrdAlteracoes_InitColumnProps()
    With sgrdAlteracoes
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
        .AllowUpdate = True
        .BackColorOdd = dCorAmarelo
        .Caption = "Alterações"
        .DataMode = ssDataModeAddItem
        .FieldSeparator = vbTab
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10
        .ForeColorEven = &H0&
        .FieldSeparator = vbTab
        .HeadFont.Name = "MS Sans Serif"
        .HeadFont.Size = 10
        .HeadFont.Bold = True
        .RowSelectionStyle = ssRowSelectionStyle3D
        .ScrollBars = ssScrollBarsNone
        .SelectByCell = False
        .SelectTypeCol = ssSelectionTypeNone
        .SelectTypeRow = ssSelectionTypeSingleSelect
       
        ' Descrição da Alteração
        .Columns(0).Alignment = ssCaptionAlignmentLeft
        .Columns(0).Caption = "Alteração"
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 8 ' Text
        .Columns(0).Width = 5500
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).Visible = True
        .Columns(0).Locked = True
        
        'Valor
        .Columns(1).Alignment = ssCaptionAlignmentRight
        .Columns(1).Caption = "Valor"
        .Columns(1).CaptionAlignment = ssColCapAlignCenter
        .Columns(1).DataType = 5 ' Double
        .Columns(1).Width = 2700
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).NumberFormat = "#####0.00"
        .Columns(1).Visible = True
    End With
End Sub


Private Sub sgrdAlteracoes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 46 Then
        KeyAscii = Asc(cSeparadorDecimal)
    End If
End Sub

Private Sub CarregaGridAlteracoes(ByRef Grid As SSDBGrid)
    Dim recALTERACOES As Recordset
    With Grid
        .Redraw = False
        .RemoveAll
        .Caption = "Alterações"
         
        If datRecibos.Recordset.RecordCount > 0 Then
            cSql = "SELECT * FROM RECIBOS WHERE NUM_RECIBO=" & sgrdGestaoRecibos.Columns(2).Value
            
            Set recALTERACOES = mBDAlteraRecibosCriados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
            If Not (recALTERACOES.EOF And recALTERACOES.BOF) Then
                lblTexto(2).Caption = "Recibo Nº " & recALTERACOES!NUM_RECIBO & " do Mês de " & _
                        cDescodificaMes(recALTERACOES!COD_MES) & " de " & recALTERACOES!ANO
    
                .AddItem "Desconto por Ausência" & vbTab & recALTERACOES!VALOR1
                .AddItem "Refeições (Almoço)" & vbTab & recALTERACOES!VALOR2
                .AddItem "Multa por Atraso" & vbTab & recALTERACOES!VALOR3
                .AddItem "Seguro" & vbTab & recALTERACOES!VALOR4
                .AddItem "Outras" & vbTab & recALTERACOES!VALOR5
                .AddItem "Mensalidade" & vbTab & recALTERACOES!MENSALIDADE
                .AddItem "1/10 ou 1/11" & vbTab & recALTERACOES!MENSALIDADE_PCTG
                .AddItem "Total do Recibo" & vbTab & recALTERACOES!TOTAL_MENSALIDADE
            End If
            recALTERACOES.Close
            Set recALTERACOES = Nothing
        Else
            lblTexto(2).Caption = "Recibo Nº  do Mês de  de "
        End If
        
        .Redraw = True
    End With
End Sub

Private Sub sgrdGestaoRecibos_Click()
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
End Sub

Private Sub sgrdGestaoRecibos_InitColumnProps()
    With sgrdGestaoRecibos
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
        .Caption = "Lista de Recibos"
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
        .Columns(1).Width = 5500
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
        
        'Nº Recibo
        .Columns(2).Alignment = ssCaptionAlignmentRight
        .Columns(2).Caption = "NºRecibo"
        .Columns(2).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(2).DataType = 3 'Long
        .Columns(2).Width = 1200
        .Columns(2).HeadStyleSet = "Cabecalho"
        .Columns(2).NumberFormat = "#####0"
        .Columns(2).Visible = True
    End With
End Sub


Private Sub sgrdGestaoRecibos_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
End Sub


