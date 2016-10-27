VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmPagaQuotaSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagar Quotas"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
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
   ScaleHeight     =   5880
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraQuotasDivida 
      Caption         =   " Quotas em Dívida "
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
      Height          =   2880
      Left            =   90
      TabIndex        =   7
      Top             =   1950
      Width           =   7230
      Begin SSDataWidgets_B.SSDBGrid sgrdQuotasDivida 
         Height          =   2565
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   6915
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
         _ExtentX        =   12197
         _ExtentY        =   4524
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
   End
   Begin VB.Frame fraSocio 
      Caption         =   " Sócio "
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
      Height          =   1800
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   7230
      Begin VB.TextBox txtNumSocio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   6900
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   225
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2003"
      Top             =   4905
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "2002"
      Top             =   4920
      Width           =   1200
   End
End
Attribute VB_Name = "frmPagaQuotaSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSPagaQuotasSocios As Workspace
Dim mBDPagaQuotasSocios As Database
    
Dim tBDAberta

Dim cSql
Dim lNum_Socio
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSPagaQuotasSocios = DBEngine.CreateWorkspace("PagaQuotasSocios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDPagaQuotasSocios = mWSPagaQuotasSocios.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Paga Quotas Sócios-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iNumLinhas, iContador
    Dim qryPagaQuotas As QueryDef
    
On Error GoTo TrataErro
    mWSPagaQuotasSocios.BeginTrans
    ' carrega a variavel com o numero de linhas -1 por causa do index
    iNumLinhas = sgrdQuotasDivida.Rows - 1
    'ciclo para precorrer todas as linhas da grid
    For iContador = 0 To iNumLinhas
        If iContador = 0 Then
            sgrdQuotasDivida.MoveFirst
        Else
            sgrdQuotasDivida.MoveNext
        End If
        If sgrdQuotasDivida.Columns(3).Value Then
            
            Set qryPagaQuotas = mBDPagaQuotasSocios.QueryDefs("QUOTAS Pagar")
                
            qryPagaQuotas.Parameters("Num_Socio") = lNum_Socio
            qryPagaQuotas.Parameters("Ano") = sgrdQuotasDivida.Columns(0).Text
            qryPagaQuotas.Parameters("Trimestre") = sgrdQuotasDivida.Columns(1).Text
            qryPagaQuotas.Parameters("Utiliz") = gUtilizador.Nome
            
            qryPagaQuotas.Execute dbFailOnError
        End If
    Next iContador
    mWSPagaQuotasSocios.CommitTrans
    ' limpa a grid para carregar de novo
    sgrdQuotasDivida.RemoveAll
    ' carrega a grid
    Call CarregaGridQuotas(sgrdQuotasDivida)
    
    GoTo SairDoProcedimento

TrataErro:
    mWSPagaQuotasSocios.Rollback
    Call ErrosGerais("Pagar Quotas", Err.Number, Err.Description)
    Resume SairDoProcedimento

SairDoProcedimento:
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridQuotas(ByRef Grid As SSDBGrid)
    Dim recTabela As Recordset, lNumRegistos, lCiclo, cLinha
    Dim cSql
    
    Grid.Redraw = False
    Grid.RemoveAll
    
    cSql = "SELECT ANO,TRIMESTRE,VALOR FROM QUOTAS WHERE NUM_SOCIO=" & lNum_Socio & " AND ISNULL(DATA_PAG)"
    cSql = cSql & " ORDER BY ANO ASC, TRIMESTRE ASC;"
   
    Set recTabela = mBDPagaQuotasSocios.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        
    If Not (recTabela.EOF And recTabela.BOF) Then
      recTabela.MoveLast
      lNumRegistos = recTabela.RecordCount
      recTabela.MoveFirst
      For lCiclo = 1 To lNumRegistos
        cLinha = vFiltraCamposNulos(recTabela.Fields("ANO")) & vbTab & _
                vFiltraCamposNulos(recTabela.Fields("TRIMESTRE")) & vbTab & _
                vFiltraCamposNulos(recTabela.Fields("VALOR")) & vbTab & _
                False
        Grid.AddItem cLinha
        recTabela.MoveNext
      Next lCiclo
    End If
        
SairDoProcedimento:
    recTabela.Close
    Set recTabela = Nothing
    Grid.Redraw = True
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    lNum_Socio = CLng(frmGestaoSocios.sgrdGestaoSocios.Columns(0).Text)
            
    Me.Caption = Me.Caption & " do Sócio Nº " & lNum_Socio
        
    Call AlteraWindowList(Me.Caption)
    
    tBDAberta = tAbreBD
        
    Call CarregaGridQuotas(sgrdQuotasDivida)
    
    txtNumSocio.Text = lNum_Socio
    txtNome.Text = frmGestaoSocios.sgrdGestaoSocios.Columns(1).Text
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSPagaQuotasSocios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSPagaQuotasSocios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub


Private Sub sgrdQuotasDivida_InitColumnProps()
    With sgrdQuotasDivida
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
        .Caption = "Quotas em Dívida"
        .DataMode = ssDataModeAddItem
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
       
        'Ano
        .Columns.Add 0
        .Columns(0).Alignment = ssCaptionAlignmentRight
        .Columns(0).Caption = "Ano"
        .Columns(0).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(0).DataType = 8 'Text
        .Columns(0).Width = 1200
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).Locked = True
        .Columns(0).Visible = True

        'Trimestre
        .Columns.Add 1
        .Columns(1).Alignment = ssCaptionAlignmentLeft
        .Columns(1).Caption = "Trim."
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).DataType = 8 ' Text
        .Columns(1).Width = 900
        .Columns(1).HeadStyleSet = "Cabecalho"
        .Columns(1).Locked = True
        .Columns(1).Visible = True
        
        'Valor
        .Columns.Add 2
        .Columns(2).Alignment = ssCaptionAlignmentRight
        .Columns(2).Caption = "Valor"
        .Columns(2).CaptionAlignment = ssColCapAlignCenter
        .Columns(2).DataType = 5 ' Double
        .Columns(2).Width = 2000
        .Columns(2).HeadStyleSet = "Cabecalho"
        .Columns(2).NumberFormat = "#####0.00"
        .Columns(2).Locked = True
        .Columns(2).Visible = True
        
        'Pagar
        .Columns.Add 3
        .Columns(3).Alignment = ssCaptionAlignmentCenter
        .Columns(3).Caption = "Pagar"
        .Columns(3).DataType = 11   'boolean
        .Columns(3).Width = 900
        .Columns(3).Style = ssStyleCheckBox
        .Columns(3).HeadStyleSet = "Cabecalho"
        .Columns(3).Locked = False
        .Columns(3).Visible = True
    End With
End Sub


