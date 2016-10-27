VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmGestaoSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestão de Sócios"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
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
   ScaleHeight     =   5895
   ScaleWidth      =   8685
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
      TabIndex        =   2
      Top             =   120
      Width           =   3000
      Begin VB.OptionButton optOrdenacao 
         Caption         =   " Nº de Sócio"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   330
         Value           =   -1  'True
         Width           =   2325
      End
      Begin VB.OptionButton optOrdenacao 
         Caption         =   "Nome do Sócio"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   4
         Top             =   660
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdQuotas 
      Caption         =   "Pagar &Quotas"
      Height          =   900
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "2011"
      Top             =   4890
      Width           =   1500
   End
   Begin VB.CommandButton cmdFicha 
      Caption         =   "&Ficha"
      Height          =   900
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "2006"
      Top             =   4890
      Width           =   1200
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   900
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2008"
      Top             =   4890
      Width           =   1200
   End
   Begin VB.Data datSocios 
      Caption         =   "datSocios"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
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
   Begin SSDataWidgets_B.SSDBGrid sgrdGestaoSocios 
      Bindings        =   "CBESQ003.frx":0000
      Height          =   3300
      Left            =   120
      TabIndex        =   5
      Top             =   1500
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
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   900
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2007"
      Top             =   4890
      Width           =   1200
   End
   Begin VB.CommandButton cmdDemissao 
      Caption         =   "Demissão"
      Height          =   900
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2012"
      Top             =   4890
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2003"
      Top             =   4890
      Width           =   1200
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Procurar por Nº de Sócio"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2610
   End
End
Attribute VB_Name = "frmGestaoSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSSocios As Workspace
Dim mBDSocios As Database

Dim tBDAberta
Dim tSelecao
Dim iRespMsgbox

Private Sub cmdAlterar_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoSocios.Rows <> 0 Then
        Set fFichaSocio = New frmFichaSocio
            
        cBotaoOrigem = "Altera"
            
        fFichaSocio.Show
    End If
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDemissao_Click()
    Dim mProcessamento As Processamento
    Dim qryDemissaoSocio As QueryDef
    
    Set mProcessamento = New Processamento
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoSocios.Rows <> 0 Then
        ' confirma se quer mesmo apagar
        iRespMsgbox = MsgBox("Confirme se o Sócio nº " & sgrdGestaoSocios.Columns(0).Value & "," & vbCrLf & _
                    "vai sair da Instituição ?", vbQuestion + vbYesNo, "Saida de Sócio")
        If iRespMsgbox = vbYes Then
On Error GoTo TrataErro
            mWSSocios.BeginTrans
            ' dar saida de utente
            
            Set qryDemissaoSocio = mBDSocios.QueryDefs("SOCIOS Data Demissao")
    
            qryDemissaoSocio.Parameters("Num_Socio") = CLng(sgrdGestaoSocios.Columns(0).Value)
            qryDemissaoSocio.Parameters("Utiliz") = gUtilizador.Nome
            
            ' executa a Saida
            qryDemissaoSocio.Execute dbFailOnError
            mWSSocios.CommitTrans
On Error GoTo 0
            ' faz o refresh
            datSocios.Refresh
            sgrdGestaoSocios.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSSocios.Rollback
    Call ErrosGerais("Saida de Sócio", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdFicha_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoSocios.Rows <> 0 Then
        Set fFichaSocio = New frmFichaSocio
            
        cBotaoOrigem = "Ficha"
            
        fFichaSocio.Show
    End If
End Sub


Private Sub cmdNovo_Click()
    Set fFichaSocio = New frmFichaSocio
        
    cBotaoOrigem = "Novo"
        
    fFichaSocio.Show
End Sub


Private Sub cmdQuotas_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdGestaoSocios.Rows <> 0 Then
        Set fPagaQuotaSocio = New frmPagaQuotaSocio
            
        fPagaQuotaSocio.Show
    End If
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridSocios(ByRef Grid As SSDBGrid, ByVal iOrdem)
    Dim recTabelaSOCIOS As Recordset
    Dim cSql
    
    Grid.Redraw = False
    
    cSql = "SELECT NUM_SOCIO,NOME FROM SOCIOS WHERE ISNULL(DATA_DEMISSAO)"
    
    Select Case iOrdem
        Case 0
            cSql = cSql & " ORDER BY NUM_SOCIO ASC"
        Case 1
            cSql = cSql & " ORDER BY NOME ASC"
    End Select
   
    Set recTabelaSOCIOS = mBDSocios.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    Set datSocios.Recordset = recTabelaSOCIOS
    
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
    
    Call CarregaGridSocios(sgrdGestaoSocios, 0)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSSocios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSSocios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmGestaoSocios = Nothing
End Sub
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSSocios = DBEngine.CreateWorkspace("Socios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDSocios = mWSSocios.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Gestão de Sócios-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function


Private Sub optOrdenacao_Click(Index As Integer)
    If Index = 0 Then
        lblTexto.Caption = "Procurar por Nº de Sócio"
    Else
        lblTexto.Caption = "Procurar por Nome do Sócio"
    End If
    ' Carrega a grid com nova Ordenação
    Call CarregaGridSocios(sgrdGestaoSocios, Index)
End Sub

Private Sub sgrdGestaoSocios_InitColumnProps()
    With sgrdGestaoSocios
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
        .Caption = "Lista de Sócios"
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
        .Columns(0).Caption = "NºSócio"
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 2 'Integer
        .Columns(0).Width = 1200
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).NumberFormat = "#####0"
        .Columns(0).Visible = True
        
        'Nome do Sócio
        .Columns(1).Alignment = ssCaptionAlignmentLeft
        .Columns(1).Caption = "Nome do Sócio"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 8 ' Text
        .Columns(1).Width = 6600
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
    End With
End Sub


Private Sub txtProcura_Change()
    If Trim$(txtProcura.Text) <> vbNullString Then
        If optOrdenacao(0).Value Then
            datSocios.Recordset.FindFirst "NUM_SOCIO=" & CInt(txtProcura.Text)
        Else
            datSocios.Recordset.FindFirst "NOME LIKE '" & txtProcura.Text & "*'"
        End If
    End If
End Sub


