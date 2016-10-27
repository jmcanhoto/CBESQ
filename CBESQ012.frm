VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmTabelaAlteracoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabela de Alterações"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
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
   ScaleHeight     =   4125
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "2002"
      Top             =   3120
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "2003"
      Top             =   3105
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid sgrdTabAlteracoes 
      Height          =   2100
      Left            =   135
      TabIndex        =   2
      Top             =   915
      Width           =   6825
      ScrollBars      =   0
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   0
      RowHeight       =   503
      Columns(0).Width=   3200
      _ExtentX        =   12039
      _ExtentY        =   3704
      _StockProps     =   79
   End
   Begin SSDataWidgets_B.SSDBCombo cboMes 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   1950
      _Version        =   196616
      DataMode        =   2
      Columns(0).Width=   3200
      _ExtentX        =   3440
      _ExtentY        =   635
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione o Mês"
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
      Width           =   1875
   End
End
Attribute VB_Name = "frmTabelaAlteracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSTabelaAlteracoes As Workspace
Dim mBDTabelaAlteracoes As Database

Dim tBDAberta

Dim cSql

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSTabelaAlteracoes = DBEngine.CreateWorkspace("Alteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDTabelaAlteracoes = mWSTabelaAlteracoes.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Tabela de Alterações Mensais-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub CarregacboMes(ByRef Combo As SSDBCombo)
    Dim recTABMESES As Recordset
    Combo.Redraw = False
        
    Combo.RemoveAll
    cSql = "SELECT NOME,COD_MES FROM TABMESES ORDER BY COD_MES"
    
    Set recTABMESES = mBDTabelaAlteracoes.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    
    If Not recTABMESES.EOF And Not recTABMESES.BOF Then
        While Not recTABMESES.EOF
            Combo.AddItem recTABMESES!Nome & vbTab & _
                            recTABMESES!COD_MES
            recTABMESES.MoveNext
        Wend
    End If
    recTABMESES.Close
    Set recTABMESES = Nothing
    Combo.Redraw = True
End Sub

Private Sub cboMes_Click()
    ' Carrega Grid
    Call CarregaGridTabAlteracoes(sgrdTabAlteracoes)
End Sub

Private Sub cboMes_InitColumnProps()
    With cboMes
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
        .Columns(0).Caption = "Mês"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Visible = False
    End With
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryTABALTERACOES As QueryDef

    Set mProcessamento = New Processamento

On Error GoTo TrataErro
    mWSTabelaAlteracoes.BeginTrans
    
    Set qryTABALTERACOES = mBDTabelaAlteracoes.QueryDefs("TABALTERACOES Altera")
    
    sgrdTabAlteracoes.Row = 0
    qryTABALTERACOES.Parameters("Valor1") = sgrdTabAlteracoes.Columns(1).Value
    sgrdTabAlteracoes.Row = 1
    qryTABALTERACOES.Parameters("Valor2") = sgrdTabAlteracoes.Columns(1).Value
    sgrdTabAlteracoes.Row = 2
    qryTABALTERACOES.Parameters("Valor3") = sgrdTabAlteracoes.Columns(1).Value
    sgrdTabAlteracoes.Row = 3
    qryTABALTERACOES.Parameters("Valor4") = sgrdTabAlteracoes.Columns(1).Value
    sgrdTabAlteracoes.Row = 4
    qryTABALTERACOES.Parameters("Valor5") = sgrdTabAlteracoes.Columns(1).Value
    qryTABALTERACOES.Parameters("Utiliz") = gUtilizador.Nome
    qryTABALTERACOES.Parameters("Mes") = cboMes.Columns(1).Text

    ' executa a inserção
    qryTABALTERACOES.Execute dbFailOnError
    mWSTabelaAlteracoes.CommitTrans
    
    MsgBox "Alterou os dados relativos ao Mês de " & cboMes.Text, vbInformation + vbOKOnly, "Tabela de Alterações"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSTabelaAlteracoes.Rollback
    Call ErrosGerais("Tabela de Alterações", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    sgrdTabAlteracoes.Row = 0
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub CarregaGridTabAlteracoes(ByRef Grid As SSDBGrid)
    Dim recTABALTERACOES As Recordset
    With Grid
        .Redraw = False
        .RemoveAll
        .Caption = "Alterações do Mês - [ " & cboMes.Text & " ]"
         
        cSql = "SELECT DESCRICAO1,DESCRICAO2,DESCRICAO3,DESCRICAO4,DESCRICAO5," & _
            "VALOR1,VALOR2,VALOR3,VALOR4,VALOR5 FROM TABALTERACOES WHERE COD_MES='" & _
            cboMes.Columns(1).Text & "'"
        
        Set recTABALTERACOES = mBDTabelaAlteracoes.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
         
        .AddItem recTABALTERACOES!DESCRICAO1 & vbTab & recTABALTERACOES!VALOR1
        .AddItem recTABALTERACOES!DESCRICAO2 & vbTab & recTABALTERACOES!VALOR2
        .AddItem recTABALTERACOES!DESCRICAO3 & vbTab & recTABALTERACOES!VALOR3
        .AddItem recTABALTERACOES!DESCRICAO4 & vbTab & recTABALTERACOES!VALOR4
        .AddItem recTABALTERACOES!DESCRICAO5 & vbTab & recTABALTERACOES!VALOR5
         
        recTABALTERACOES.Close
        Set recTABALTERACOES = Nothing
        .Redraw = True
    End With
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Me.Show
    DoEvents
    
    tBDAberta = tAbreBD
    
    ' carrega a combo
    Call CarregacboMes(cboMes)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSTabelaAlteracoes.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSTabelaAlteracoes = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmTabelaAlteracoes = Nothing
End Sub


Private Sub sgrdTabAlteracoes_BeforeRowColChange(Cancel As Integer)
    If IsEmpty(sgrdTabAlteracoes.Columns(1).Value) Then
        sgrdTabAlteracoes.Columns(1).Value = 0
    End If
End Sub

Private Sub sgrdTabAlteracoes_InitColumnProps()
    With sgrdTabAlteracoes
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
        .Caption = "Alterações do Mês - [ " & cboMes.Text & " ]"
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


Private Sub sgrdTabAlteracoes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 46 Then
        KeyAscii = Asc(cSeparadorDecimal)
    End If
End Sub


