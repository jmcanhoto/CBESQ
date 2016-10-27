VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmGestaoFuncionarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestão de Funcionários"
   ClientHeight    =   5865
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
   ScaleHeight     =   5865
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   4
      Top             =   120
      Width           =   3000
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nome do Funcionário"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   6
         Top             =   690
         Width           =   2475
      End
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nº de Funcionário"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "2003"
      Top             =   4860
      Width           =   1200
   End
   Begin VB.CommandButton cmdDemissao 
      Caption         =   "&Demissão"
      Height          =   900
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2012"
      Top             =   4845
      Width           =   1200
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   900
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2007"
      Top             =   4875
      Width           =   1200
   End
   Begin VB.TextBox txtProcura 
      Height          =   330
      Left            =   120
      MaxLength       =   25
      TabIndex        =   1
      Top             =   390
      Width           =   3000
   End
   Begin VB.Data datFuncionarios 
      Caption         =   "datFuncionarios"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4350
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   900
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2008"
      Top             =   4860
      Width           =   1200
   End
   Begin VB.CommandButton cmdFicha 
      Caption         =   "&Ficha"
      Height          =   900
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2006"
      Top             =   4875
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoFuncionarios 
      Bindings        =   "CBESQ007.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   7
      Top             =   1455
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
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar os Funcionários da Instituição"
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
      Width           =   4410
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por Nº de Funcionário"
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
      Width           =   3225
   End
End
Attribute VB_Name = "frmGestaoFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSFuncionarios As Workspace
Dim mBDFuncionarios As Database

Dim tBDAberta
Dim tSelecao
Dim iRespMsgbox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSFuncionarios = DBEngine.CreateWorkspace("Funcionarios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDFuncionarios = mWSFuncionarios.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Gestão de Funcionários-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    ' Carrega a grid com nova Ordenação
    Call CarregaGridFuncionarios(sgrdGestaoFuncionarios, IIf(optOrdenacao(0).Value, 0, 1))
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

Private Sub cmdAlterar_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoFuncionarios.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaFuncionario = New frmFichaFuncionario
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Altera"
        ' faz o load da form
        fFichaFuncionario.Show
    End If
End Sub

Private Sub cmdDemissao_Click()
    Dim mProcessamento As Processamento
    Dim qryApagarFuncionario As QueryDef
    
    Set mProcessamento = New Processamento
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoFuncionarios.Rows <> 0 Then
        ' confirma se quer mesmo apagar
        iRespMsgbox = MsgBox("Confirme se o Funcionário nº " & sgrdGestaoFuncionarios.Columns(0).Value & "," & vbCrLf & _
                    "vai sair da Instituição ?", vbQuestion + vbYesNo, "Saida de Funcionário")
        If iRespMsgbox = vbYes Then
            If IsDate(sgrdGestaoFuncionarios.Columns(2).Value) Then
                MsgBox "Este Funcionário já tem Data de Demissão registada !!!", vbOKOnly, "Saida de Utente"
                Exit Sub
            End If
On Error GoTo TrataErro
            mWSFuncionarios.BeginTrans
            ' apagar Inscrição
            Set qryApagarFuncionario = mBDFuncionarios.QueryDefs("FUNCIONARIOS Data Demissao")
    
            qryApagarFuncionario.Parameters("Num_Func") = CLng(sgrdGestaoFuncionarios.Columns(0).Value)
            qryApagarFuncionario.Parameters("Utiliz") = gUtilizador.Nome
            
            ' executa a query
            qryApagarFuncionario.Execute dbFailOnError
            mWSFuncionarios.CommitTrans
On Error GoTo 0
            ' faz o refresh
            datFuncionarios.Refresh
            sgrdGestaoFuncionarios.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSFuncionarios.Rollback
    Call ErrosGerais("Data de Demissão de Funcionário", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFicha_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoFuncionarios.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaFuncionario = New frmFichaFuncionario
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Ficha"
        ' faz o load da form
        fFichaFuncionario.Show
    End If
End Sub

Private Sub cmdNovo_Click()
    ' cria nova instancia da Form
    Set fFichaFuncionario = New frmFichaFuncionario
    ' Carrega a Variavel para passar o Botao de onde vai
    cBotaoOrigem = "Novo"
    ' faz o load da form
    fFichaFuncionario.Show
End Sub


Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridFuncionarios(ByRef Grid As SSDBGrid, ByVal iOrdem)
    Dim recTabelaFUNCIONARIOS As Recordset
    Dim cLinha
    Dim cSql
    
    Grid.Redraw = False
    
    cSql = "SELECT NUM_FUNCIONARIO,NOME,DATA_DEMISSAO FROM FUNCIONARIOS WHERE ISNULL(DATA_DEMISSAO)"
    
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cboInstituicao.Text) & "'"
    End If
    
    Select Case iOrdem
        Case 0
            cSql = cSql & " ORDER BY NUM_FUNCIONARIO ASC"
        Case 1
            cSql = cSql & " ORDER BY NOME ASC"
    End Select
   
    Set recTabelaFUNCIONARIOS = mBDFuncionarios.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    Set datFuncionarios.Recordset = recTabelaFUNCIONARIOS
    
SairDoProcedimento:
    Grid.Redraw = True
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    tSelecao = True
    
    Call CarregaGridFuncionarios(sgrdGestaoFuncionarios, 0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSFuncionarios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSFuncionarios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmGestaoFuncionarios = Nothing
End Sub


Private Sub optOrdenacao_Click(Index As Integer)
    If Index = 0 Then
        lblTexto(0).Caption = "Procurar por Nº de Funcionário"
    Else
        lblTexto(0).Caption = "Procurar por Nome do Funcionário"
    End If
    ' Carrega a grid com nova Ordenação
    Call CarregaGridFuncionarios(sgrdGestaoFuncionarios, Index)
End Sub


Private Sub sgrdGestaoFuncionarios_InitColumnProps()
    With sgrdGestaoFuncionarios
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
        .Caption = "Lista de Funcionários"
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
       
        'Nº Sócio
        .Columns(0).Alignment = ssCaptionAlignmentRight
        .Columns(0).Caption = "NºFunc."
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 2 'Integer
        .Columns(0).Width = 1200
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).NumberFormat = "#####0"
        .Columns(0).Visible = True
        
        'Nome do Sócio
        .Columns(1).Alignment = ssCaptionAlignmentLeft
        .Columns(1).Caption = "Nome do Funcionário"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 8 ' Text
        .Columns(1).Width = 6600
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
        
        'Data de Demissao
        .Columns(2).DataType = 7 'date
        .Columns(2).Visible = False
    
    End With
End Sub


Private Sub txtProcura_Change()
    If Trim$(txtProcura.Text) <> vbNullString Then
        If optOrdenacao(0).Value Then
            datFuncionarios.Recordset.FindFirst "NUM_FUNCIONARIO=" & CInt(txtProcura.Text)
        Else
            datFuncionarios.Recordset.FindFirst "NOME LIKE '" & txtProcura.Text & "*'"
        End If
    End If
End Sub
