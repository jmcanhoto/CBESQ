VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAlterarRecibosSala 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestao de Recibos da Sala"
   ClientHeight    =   5565
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
   ScaleHeight     =   5565
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   5745
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "2003"
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "2002"
      Top             =   4560
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdAlteracoes 
      Height          =   2520
      Left            =   120
      TabIndex        =   5
      Top             =   1785
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
      _ExtentY        =   4445
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
      TabIndex        =   3
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
   Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
      Height          =   330
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   4
      Top             =   1455
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
      TabIndex        =   0
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
      TabIndex        =   2
      Top             =   720
      Width           =   3300
   End
End
Attribute VB_Name = "frmAlterarRecibosSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSAlteraRecibosSala As Workspace
Dim mBDAlteraRecibosSala As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgBox

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSAlteraRecibosSala = DBEngine.CreateWorkspace("AlteracoesSala", gUtilizador.Nome, gUtilizador.Password)
    Set mBDAlteraRecibosSala = mWSAlteraRecibosSala.OpenDatabase(cBD_Path & cNomeBD)
    
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

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryAlteraRecibosSala As QueryDef

    Set mProcessamento = New Processamento
    
    ' pede confirmação se quer continuar
    iRespMsgBox = MsgBox("Confirma que quer Alterar os Dados dos Recibos, " & vbCrLf & _
                " da Sala : " & cboSalas.Text & ", da Instituição : " & cboInstituicao.Text & " !", _
                        vbQuestion + vbYesNo, "Alterar Recibos")
    ' se resposta não sai
    If iRespMsgBox = vbNo Then
        GoTo SairDoProcedimento
    End If

On Error GoTo TrataErro
    mWSAlteraRecibosSala.BeginTrans
    ' Actualiza Valores
    Set qryAlteraRecibosSala = mBDAlteraRecibosSala.QueryDefs("RECIBOS Altera Recibos Sala")
    
    sgrdAlteracoes.Row = 0
    qryAlteraRecibosSala.Parameters("Valor1") = sgrdAlteracoes.Columns(1).Value
    sgrdAlteracoes.Row = 1
    qryAlteraRecibosSala.Parameters("Valor2") = sgrdAlteracoes.Columns(1).Value
    sgrdAlteracoes.Row = 2
    qryAlteraRecibosSala.Parameters("Valor3") = sgrdAlteracoes.Columns(1).Value
    sgrdAlteracoes.Row = 3
    qryAlteraRecibosSala.Parameters("Valor4") = sgrdAlteracoes.Columns(1).Value
    sgrdAlteracoes.Row = 4
    qryAlteraRecibosSala.Parameters("Valor5") = sgrdAlteracoes.Columns(1).Value
    qryAlteraRecibosSala.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
    qryAlteraRecibosSala.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)

    ' executa a inserção
    qryAlteraRecibosSala.Execute dbFailOnError
    
    ' Actualiza Total Mensal
    Set qryAlteraRecibosSala = mBDAlteraRecibosSala.QueryDefs("RECIBOS Altera Recibos Sala Actualiza Total")
    
    qryAlteraRecibosSala.Parameters("Utiliz") = gUtilizador.Nome
    qryAlteraRecibosSala.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
    qryAlteraRecibosSala.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)

    ' executa a inserção
    qryAlteraRecibosSala.Execute dbFailOnError
    mWSAlteraRecibosSala.CommitTrans
    
    ' carrega a grid
    Call CarregaGridAlteracoes(sgrdAlteracoes)
    
    MsgBox "Alterou os dados relativos à Sala : " & _
            cboSalas.Text & ", da Instituição : " & _
            cboInstituicao.Text & " !", vbInformation + vbOKOnly, "Alterar Recibos"
    
        
    GoTo SairDoProcedimento
    
TrataErro:
    mWSAlteraRecibosSala.Rollback
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
    
    Call CarregaGridAlteracoes(sgrdAlteracoes)
    
    tBDAberta = tAbreBD
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSAlteraRecibosSala.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSAlteraRecibosSala = Nothing
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
    With Grid
        .Redraw = False
        .RemoveAll
        .Caption = "Alterações"
         
        .AddItem "Desconto por Ausência" & vbTab & 0
        .AddItem "Refeições (Almoço)" & vbTab & 0
        .AddItem "Multa por Atraso" & vbTab & 0
        .AddItem "Seguro" & vbTab & 0
        .AddItem "Outras" & vbTab & 0
        
        .Redraw = True
    End With
End Sub

