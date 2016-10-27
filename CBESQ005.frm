VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmGestaoInscricoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestão de Inscrições"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Tag             =   "1002"
   Begin VB.CheckBox chkTodos 
      Caption         =   "Apagar Insc. Ano Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6225
      TabIndex        =   14
      Top             =   4830
      Width           =   2730
   End
   Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
      Height          =   330
      Left            =   135
      TabIndex        =   3
      ToolTipText     =   "Seleção da Instituição que deseja consultar"
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
      Height          =   1215
      Left            =   5565
      TabIndex        =   4
      Top             =   90
      Width           =   3000
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nome"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   6
         Top             =   690
         Width           =   2550
      End
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nº de Inscrição"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "2003"
      ToolTipText     =   "Sair da Gestão de Inscrições"
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdApagar 
      Caption         =   "A&pagar"
      Height          =   900
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "2004"
      ToolTipText     =   "Apagar a Inscrição"
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Nova"
      Height          =   900
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2007"
      ToolTipText     =   "Nova Inscrição"
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   900
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2008"
      ToolTipText     =   "Alterar a Ficha de Inscrição"
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdFicha 
      Caption         =   "&Ficha"
      Height          =   900
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2006"
      ToolTipText     =   "Visualizar a Ficha de Inscrição"
      Top             =   5130
      Width           =   1200
   End
   Begin VB.CommandButton cmdPassar 
      Caption         =   "&Passar Utente"
      Height          =   900
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2005"
      ToolTipText     =   "Passar a Inscrição para Novo Utente"
      Top             =   5130
      Width           =   1500
   End
   Begin VB.Data datInscricoes 
      Caption         =   "datInscricoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4260
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtProcura 
      Height          =   330
      Left            =   120
      MaxLength       =   25
      TabIndex        =   1
      ToolTipText     =   "Digite o que deseja procurar"
      Top             =   390
      Width           =   3000
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoInscricoes 
      Bindings        =   "CBESQ005.frx":0000
      Height          =   3270
      Left            =   120
      TabIndex        =   7
      Top             =   1455
      Width           =   8490
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
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   14975
      _ExtentY        =   5768
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
      Caption         =   "Seleccionar Inscrições da Instituição"
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
      Width           =   3825
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por Nº de Inscrição"
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
      Width           =   2955
   End
End
Attribute VB_Name = "frmGestaoInscricoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSInscricoes As Workspace
Dim mBDInscricoes As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSInscricoes = DBEngine.CreateWorkspace("GestaoInscricoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDInscricoes = mWSInscricoes.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Gestão de Inscrições-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    ' Carrega a grid com nova Ordenação
    txtProcura.Text = vbNullString
    Call CarregaGridInscricoes(sgrdGestaoInscricoes, IIf(optOrdenacao(0).Value, 0, 1))
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
    If sgrdGestaoInscricoes.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaInscricao = New frmFichaInscricao
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Altera"
        ' faz o load da form
        fFichaInscricao.Show
    End If
End Sub

Private Sub cmdApagar_Click()
    Dim mProcessamento As Processamento
    Dim qryApagarInscricao As QueryDef
    
    Set mProcessamento = New Processamento
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoInscricoes.Rows <> 0 Then
        ' confirma se quer mesmo apagar
            If chkTodos.Value = vbUnchecked Then
                iRespMsgBox = MsgBox("Confirme se quer apagar a Inscrição nº " & sgrdGestaoInscricoes.Columns(0).Value & " ?", vbYesNo, "Apagar Inscrição")
            Else
                iRespMsgBox = MsgBox("Confirme se quer apagar todas as Inscrições do Ano anterior ?", vbYesNo, "Apagar Inscrição")
            End If
        If iRespMsgBox = vbYes Then
On Error GoTo TrataErro
            mWSInscricoes.BeginTrans
            ' apagar Inscrição
            If chkTodos.Value = vbUnchecked Then
                Set qryApagarInscricao = mBDInscricoes.QueryDefs("INSCRICOES Apagar")
        
                qryApagarInscricao.Parameters("Num_Inscricao") = CLng(sgrdGestaoInscricoes.Columns(0).Value)
            Else
'                Set qryApagarInscricao = mBDInscricoes.QueryDefs("INSCRICOES Apagar Tudo")
                Set qryApagarInscricao = mBDInscricoes.QueryDefs("INSCRICOES Apagar Ano Anterior")
                
                qryApagarInscricao.Parameters("Inicio") = CDate("01-01-" & CStr(Year(Date) - 1))
                qryApagarInscricao.Parameters("Fim") = CDate("31-12-" & CStr(Year(Date) - 1))
            End If
            ' executa a query
            qryApagarInscricao.Execute dbFailOnError
            mWSInscricoes.CommitTrans
On Error GoTo 0
            ' faz o refresh
            datInscricoes.Refresh
           sgrdGestaoInscricoes.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSInscricoes.Rollback
    Call ErrosGerais("Apagar Inscrição", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFicha_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoInscricoes.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaInscricao = New frmFichaInscricao
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Ficha"
        ' faz o load da form
        fFichaInscricao.Show
    End If
End Sub

Private Sub cmdNovo_Click()
    ' cria nova instancia da Form
    Set fFichaInscricao = New frmFichaInscricao
    ' Carrega a Variavel para passar o Botao de onde vai
    cBotaoOrigem = "Novo"
    ' faz o load da form
    fFichaInscricao.Show
End Sub

Private Sub cmdPassar_Click()
    Dim mProcessamento As Processamento
    Dim qryPassaUtente As QueryDef
    
    Set mProcessamento = New Processamento
    ' passa o registo da inscrição a utente
    If sgrdGestaoInscricoes.Rows <> 0 Then
        ' confirma se quer mesmo apagar
        iRespMsgBox = MsgBox("Confirme se quer aceitar esta Inscrição nº " & sgrdGestaoInscricoes.Columns(0).Value & " ?", vbYesNo, "Passar a Utente")
        If iRespMsgBox = vbYes Then
            ' começa a transação
            mWSInscricoes.BeginTrans
            ' Passar a Inscrição a Utente
            Set qryPassaUtente = mBDInscricoes.QueryDefs("INSCRICOES Passa a Utente")
            ' parametros de input
            qryPassaUtente.Parameters("Num_Inscricao") = sgrdGestaoInscricoes.Columns(0).Value
            qryPassaUtente.Parameters("Utiliz") = gUtilizador.Nome
            ' executa a query
            qryPassaUtente.Execute dbFailOnError
            ' Copia os Daos para Utentes
            Set qryPassaUtente = mBDInscricoes.QueryDefs("INSCRICOES Copia para UTENTES")
            ' parametros de input
            qryPassaUtente.Parameters("Num_Utente") = lNovoNumUtente
            qryPassaUtente.Parameters("Num_Inscricao") = sgrdGestaoInscricoes.Columns(0).Value
            qryPassaUtente.Parameters("Utiliz") = gUtilizador.Nome
            ' executa a query
            qryPassaUtente.Execute dbFailOnError
            ' fecha a transação
            mWSInscricoes.CommitTrans
            ' faz o refresh
            datInscricoes.Refresh
            sgrdGestaoInscricoes.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSInscricoes.Rollback
    Call ErrosGerais("Passa a Utente", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub Form_Activate()
    ' marca a janela activa na lista de janelas abertas
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridInscricoes(ByRef Grid As SSDBGrid, ByVal iOrdem)
    ' define a variavel da tabela
    Dim recTabelaINSCRICOES As Recordset
    
    Grid.Redraw = False
    
    ' atribui o SQL a variavel
    cSql = "SELECT NUM_INSCRICAO,DATA_NASC,NOME FROM INSCRICOES WHERE UTENTE=False"
    ' se houver uma intituição seleccionada
    If cboInstituicao.Text <> "<Todas as Instituições>" Then
        cSql = cSql & " AND COD_INST='" & cCodificaInstituicao(cboInstituicao.Text) & "'"
    End If
    ' seleciona a ordem dos dados
    Select Case iOrdem
        Case 0
            cSql = cSql & " ORDER BY NUM_INSCRICAO ASC"
        Case 1
            cSql = cSql & " ORDER BY NOME ASC"
    End Select
    ' abre a tabela
    Set recTabelaINSCRICOES = mBDInscricoes.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'passa os dados para o  data control
    Set datInscricoes.Recordset = recTabelaINSCRICOES
    
    Grid.Redraw = True
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
    Call CarregaGridInscricoes(sgrdGestaoInscricoes, 0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' fecha as tabelas abertas e as bases de dados abertas no Workspace
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSInscricoes.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    ' limpa o workspace
    Set mWSInscricoes = Nothing
    ' apaga a janela da lista de janelas abertas
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    ' limpa a form de memoria
    Set frmGestaoInscricoes = Nothing
End Sub


Private Sub optOrdenacao_Click(Index As Integer)
    If Index = 0 Then
        lblTexto(0).Caption = "Procurar por Nº de Inscrição"
        Call SetNumber(txtProcura, True)
        txtProcura.Text = vbNullString
    Else
        lblTexto(0).Caption = "Procurar por Nome"
        Call SetNumber(txtProcura, False)
        txtProcura.Text = vbNullString
    End If
    ' Carrega a grid com nova Ordenação
    Call CarregaGridInscricoes(sgrdGestaoInscricoes, Index)
End Sub

Private Sub sgrdGestaoInscricoes_InitColumnProps()
    With sgrdGestaoInscricoes
        If .StyleSets.Count = 0 Then
            .StyleSets.Add "Cabecalho"
            .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
            .StyleSets("Cabecalho").ForeColor = vbTitleBarText
            .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
            .StyleSets("Cabecalho").Font.Size = 10
            .StyleSets("Cabecalho").Font.Bold = True
        End If
        
        .AllowAddNew = False
        .AllowColumnMoving = ssRelocateNotAllowed
        .AllowColumnShrinking = False
        .AllowColumnSizing = False
        .AllowColumnSwapping = ssRelocateNotAllowed
        .AllowDelete = False
        .AllowDragDrop = False
        .AllowGroupMoving = False
        .AllowGroupShrinking = False
        .AllowGroupSizing = False
        .AllowRowSizing = False
        .AllowUpdate = False
        .BackColorOdd = dCorAmarelo
        .Caption = "Lista de Inscrições"
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
       
        'Nº Inscrição
        .Columns(0).Alignment = ssCaptionAlignmentRight
        .Columns(0).Caption = "Nº Insc."
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 3 'Long
        .Columns(0).Width = 1100
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).NumberFormat = "#####0"
        .Columns(0).Visible = True
        
        'Data de Nascimento
        .Columns(1).Alignment = ssCaptionAlignmentCenter
        .Columns(1).Caption = "Data Nasc."
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 7 ' Date
        .Columns(1).Width = 1300
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
        
        'Nome da Incrição
        .Columns(2).Alignment = ssCaptionAlignmentLeft
        .Columns(2).Caption = "Nome"
        .Columns(2).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(2).DataType = 8 ' Text
        .Columns(2).Width = 5750
        .Columns(2).HeadStyleSet = "Cabecalho"
        .Columns(2).Visible = True
    End With
End Sub


Private Sub txtProcura_Change()
    If Trim$(txtProcura.Text) <> vbNullString Then
        If optOrdenacao(0).Value Then
            datInscricoes.Recordset.FindFirst "NUM_INSCRICAO=" & CLng(txtProcura.Text)
        Else
            datInscricoes.Recordset.FindFirst "NOME LIKE '" & txtProcura.Text & "*'"
        End If
    End If
End Sub


