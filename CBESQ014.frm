VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAlterarRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestao de Recibos de Utentes"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
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
   ScaleHeight     =   6690
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
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
      Left            =   5745
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2003"
      Top             =   5595
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "2002"
      Top             =   5595
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdAlteracoes 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   2445
      Width           =   6825
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
      _ExtentX        =   12039
      _ExtentY        =   5318
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
      Top             =   2115
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
Attribute VB_Name = "frmAlterarRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSAlteraRecibos As Workspace
Dim mBDAlteraRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox



'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSAlteraRecibos = DBEngine.CreateWorkspace("Alteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDAlteraRecibos = mWSAlteraRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Alteração de Recibos Mensais-Abrir BD", Err.Number, Err.Description)
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


Private Sub cboSalas_DropDown()
    ' carrega a combo
    Call CarregacboSalas(cboSalas, cboInstituicao.Text)
    cboNumUtente.Text = vbNullString
    cboNome.Text = vbNullString
    sgrdAlteracoes.Redraw = False
    sgrdAlteracoes.RemoveAll
    sgrdAlteracoes.Redraw = True
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

    If Trim$(cboNumUtente.Text) <> vbNullString Then

        Set mProcessamento = New Processamento
        
        ' pede confirmação se quer continuar
        iRespMsgBox = MsgBox("Confirma que quer Alterar os Dados do Recibo.", _
                            vbQuestion + vbYesNo, "Alterar Recibos")
        ' se resposta não sai
        If iRespMsgBox = vbNo Then
            GoTo SairDoProcedimento
        End If
On Error GoTo TrataErro
        mWSAlteraRecibos.BeginTrans
    
        Set qryAlteraRecibo = mBDAlteraRecibos.QueryDefs("RECIBOS Altera Recibo")
        
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
        qryAlteraRecibo.Parameters("Num_Utente") = CLng(cboNumUtente.Text)
    
        ' executa a inserção
        qryAlteraRecibo.Execute dbFailOnError
        mWSAlteraRecibos.CommitTrans
        
        ' carrega a grid
        Call CarregaGridAlteracoes(sgrdAlteracoes)
        
        MsgBox "Alterou os dados relativos ao Utente.", vbInformation + vbOKOnly, "Alterar Recibos"
    
    End If
        
    GoTo SairDoProcedimento
    
TrataErro:
    mWSAlteraRecibos.Rollback
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
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSAlteraRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSAlteraRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmAlterarRecibos = Nothing
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
        .Columns(0).Width = 4500
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).Visible = True
        .Columns(0).Locked = True
        
        'Valor
        .Columns(1).Alignment = ssCaptionAlignmentRight
        .Columns(1).Caption = "Valor"
        .Columns(1).CaptionAlignment = ssColCapAlignCenter
        .Columns(1).DataType = 5 ' Double
        .Columns(1).Width = 2000
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
         
        cSql = "SELECT * FROM RECIBOS WHERE NUM_UTENTE=" & _
            CLng(cboNumUtente.Text) & " AND (ESTADO_REC) IS NULL"
        
        Set recALTERACOES = mBDAlteraRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
        If Not (recALTERACOES.EOF And recALTERACOES.BOF) Then
            lblTexto(2).Caption = "Recibo do Mês de " & _
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
        
        .Redraw = True
    End With
End Sub

