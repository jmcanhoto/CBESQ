VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAbsentismo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Absentismo"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
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
   ScaleHeight     =   5880
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApagar 
      Caption         =   "A&pagar"
      Height          =   900
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2004"
      ToolTipText     =   "Apagar a Inscrição"
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   900
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "2008"
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2003"
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   900
      Left            =   870
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "2007"
      Top             =   4950
      Width           =   1200
   End
   Begin VB.Data datAbsentismo 
      Caption         =   "datAbsentismo"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4545
      Visible         =   0   'False
      Width           =   2535
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdAbsentismo 
      Bindings        =   "CBESQ024.frx":0000
      Height          =   2355
      Left            =   150
      TabIndex        =   5
      Top             =   2520
      Width           =   5610
      _Version        =   196617
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   9895
      _ExtentY        =   4154
      _StockProps     =   79
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
      TabIndex        =   0
      Top             =   150
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   390
         Width           =   5265
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
         Columns(0).Width=   3200
         _ExtentX        =   9287
         _ExtentY        =   582
         _StockProps     =   93
         Text            =   "<Todas as Instituições>"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraFuncionario 
      Caption         =   " Funcionário "
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
      Height          =   1350
      Left            =   150
      TabIndex        =   2
      Top             =   1095
      Width           =   5610
      Begin SSDataWidgets_B.SSDBCombo cboNumFunc 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   390
         Width           =   1350
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
         Columns(0).Width=   3200
         _ExtentX        =   2381
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBCombo cboNome 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   780
         Width           =   5340
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
         Columns(0).Width=   3200
         _ExtentX        =   9419
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAbsentismo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSAbsentismo As Workspace
Dim mBDAbsentismo As Database

Dim tBDAberta
Dim cSql

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSAbsentismo = DBEngine.CreateWorkspace("Absentismo", gUtilizador.Nome, gUtilizador.Password)
    Set mBDAbsentismo = mWSAbsentismo.OpenDatabase(cBD_Path & cNomeBD)

    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Absentismo-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboInstituicao_Click()
    cboNumFunc.Text = vbNullString
    cboNome.Text = vbNullString
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
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
End Sub


Private Sub cboNome_Click()
    ' poe o num de Funcionario na combo
    cboNumFunc.Text = cboNome.Columns(1).Value
    Call CarregaGridAbsentismo(sgrdAbsentismo, cboNome.Columns(1).Value)
End Sub

Private Sub cboNome_DropDown()
    ' carrega a combo
    Call CarregacboNomeFunc(cboNome, cboInstituicao.Text)
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
        .Columns(0).Caption = "Nome do Func."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
End Sub

Private Sub cboNumFunc_Click()
'    If Trim$(cboNumFunc.Text) <> vbNullString Then
        ' poe o num de utente na combo
        cboNome.Text = cboNumFunc.Columns(1).Text
        Call CarregaGridAbsentismo(sgrdAbsentismo, cboNumFunc.Columns(0).Value)
'    End If
End Sub

Private Sub cboNumFunc_DropDown()
    ' carrega a combo
    Call CarregacboNumFunc(cboNumFunc, cboInstituicao.Text)
End Sub

Private Sub cboNumFunc_InitColumnProps()
    With cboNumFunc
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
        .Columns(0).Caption = "Nº Func."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
                
    End With
End Sub

Private Sub cmdAlterar_Click()
    ' se houver registos Edita o que estiver selecionado
    If sgrdAbsentismo.Rows <> 0 Then
        ' cria nova instancia da Form
        Set fFichaAbsent = New frmFichaAbsent
        ' Carrega a Variavel para passar o Botao de onde vai
        cBotaoOrigem = "Altera"
        ' faz o load da form
        fFichaAbsent.Show
    End If
End Sub

Private Sub cmdApagar_Click()
    Dim mProcessamento As Processamento
    Dim qryApagarAbsent As QueryDef
    Dim iRespMsgBox
    Set mProcessamento = New Processamento
    ' se houver registos Edita o que estiver selecionado
    If sgrdAbsentismo.Rows <> 0 Then
        ' confirma se quer mesmo apagar
        iRespMsgBox = MsgBox("Confirme se quer apagar o Absentismo ?", vbYesNo, "Apagar Absentismo")
        If iRespMsgBox = vbYes Then
On Error GoTo TrataErro
            mWSAbsentismo.BeginTrans
            ' apagar Inscrição
            Set qryApagarAbsent = mBDAbsentismo.QueryDefs("Absentismo Apagar")
    
            qryApagarAbsent.Parameters("Contador") = CLng(sgrdAbsentismo.Columns(5).Value)

            ' executa a query
            qryApagarAbsent.Execute dbFailOnError
            mWSAbsentismo.CommitTrans
On Error GoTo 0
            ' faz o refresh
            datAbsentismo.Refresh
           sgrdAbsentismo.Refresh
        End If
    End If
    GoTo SairDoProcedimento
    
TrataErro:
    mWSAbsentismo.Rollback
    Call ErrosGerais("Apagar Absentismo", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdNovo_Click()
    If Trim$(cboNumFunc.Text) = vbNullString Then
        MsgBox "Tem de seleccionar um Funcionário!", vbInformation, "Absentismo"
        Exit Sub
    End If
    ' cria nova instancia da Form
    Set fFichaAbsent = New frmFichaAbsent
    ' Carrega a Variavel para passar o Botao de onde vai
    cBotaoOrigem = "Novo"
    ' faz o load da form
    fFichaAbsent.Show
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridAbsentismo(ByRef Grid As SSDBGrid, ByVal lNumFunc)
    Dim recABSENTISMO As Recordset
    Dim cLinha
    Dim cSql
    
    Grid.Redraw = False
    
    cSql = "SELECT DATA_FALTA,HORA_INI & ':' & MIN_INI AS INICIO," & _
                "HORA_FIM & ':' & MIN_FIM AS FIM," & _
                " MIN_TOTAL AS TOTAL,MOTIVO_FALTA,CONTADOR FROM ABSENTISMO"
    
    cSql = cSql & " WHERE NUM_FUNCIONARIO=" & lNumFunc & " ORDER BY DATA_FALTA ASC"
   
    Set recABSENTISMO = mBDAbsentismo.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    
    Set datAbsentismo.Recordset = recABSENTISMO
    
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
    
    Call CarregaGridAbsentismo(sgrdAbsentismo, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSAbsentismo.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSAbsentismo = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmAbsentismo = Nothing
End Sub


Private Sub sgrdAbsentismo_InitColumnProps()
    With sgrdAbsentismo
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
        .Caption = "Lista de Absentismo"
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10
        .ForeColorEven = &H0&
        .FieldSeparator = vbTab
        .HeadFont.Name = "MS Sans Serif"
        .HeadFont.Size = 10
        .HeadFont.Bold = True
        .RowSelectionStyle = ssRowSelectionStyle3D
        .ScrollBars = ssScrollBarsBoth
        .SelectByCell = False
        .SelectTypeCol = ssSelectionTypeNone
        .SelectTypeRow = ssSelectionTypeSingleSelect
       
        'Data
        .Columns(0).Alignment = ssCaptionAlignmentCenter
        .Columns(0).Caption = "Data"
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 7 'Date
        .Columns(0).Width = 1200
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).Visible = True
        
        ' Inicio
        .Columns(1).Alignment = ssCaptionAlignmentCenter
        .Columns(1).Caption = "Inicio"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 8 ' Text
        .Columns(1).Width = 750
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Visible = True
        
        ' Fim
        .Columns(2).Alignment = ssCaptionAlignmentCenter
        .Columns(2).Caption = "Fim"
        .Columns(2).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(2).DataType = 8 ' Text
        .Columns(2).Width = 750
        .Columns(2).HeadStyleSet = "Cabecalho"
        .Columns(2).Visible = True
    
        ' Total
        .Columns(3).Alignment = ssCaptionAlignmentRight
        .Columns(3).Caption = "Total"
        .Columns(3).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(3).DataType = 8 ' Text
        .Columns(3).Width = 750
        .Columns(3).HeadStyleSet = "Cabecalho"
        .Columns(3).Visible = True
        
        ' Motivo
        .Columns(4).Alignment = ssCaptionAlignmentLeft
        .Columns(4).Caption = "Motivo"
        .Columns(4).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(4).DataType = 8 ' Text
        .Columns(4).Width = 4500
        .Columns(4).HeadStyleSet = "Cabecalho"
        .Columns(4).Visible = True
    
        ' Contador
        .Columns(5).Visible = False
    End With
End Sub


