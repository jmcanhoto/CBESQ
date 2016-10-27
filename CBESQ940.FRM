VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmPerfil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perfil de Acesso"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
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
   ScaleHeight     =   4425
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPerfil 
      Appearance      =   0  'Flat
      Caption         =   "Perfil "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   6765
      TabIndex        =   3
      Top             =   45
      Width           =   2175
      Begin SSDataWidgets_B.SSDBCombo scboPerfil 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1950
         _Version        =   196616
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
         _ExtentX        =   3440
         _ExtentY        =   529
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
   Begin SSDataWidgets_B.SSDBGrid sgrdMenu 
      Height          =   4170
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   6420
      _Version        =   196616
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
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   11324
      _ExtentY        =   7355
      _StockProps     =   79
      Caption         =   "Opções do Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6645
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   3390
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "2003"
      Top             =   3390
      Width           =   1200
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSTabPer As Workspace
Dim mBDTabPer As Database
Dim tBDAberta
Dim cToolID()
Dim iLarguraColuna()
Dim iContador


'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    
    Set mWSTabPer = DBEngine.CreateWorkspace("Perfil", gUtilizador.Nome, gUtilizador.Password)
    Set mBDTabPer = mWSTabPer.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Perfis-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
    
End Function

Private Function cFiltra(ByVal sParse As String, ByVal sDirtyChr As String)

    ' This function accepts a string and a character to remove from that string,
    ' and returns the new string, which will not contain the character

    Dim lDirty As Long
    Dim sLeft As String
    Dim sRight As String
    
    lDirty = InStr(sParse, sDirtyChr)
    
    Do While lDirty
    
        sParse = Left$(sParse, lDirty - 1) & Right$(sParse, Len(sParse) - lDirty - Len(sDirtyChr) + 1)
        lDirty = InStr(sParse, sDirtyChr)
    
    Loop
    
    cFiltra = sParse

End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call GravaDadosNaBD
End Sub

Private Sub Form_Activate()
        fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    'Carrega para a grelha o menu
    Call CarregaMenu
    
    tBDAberta = tAbreBD
    
    Me.Show
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tBDAberta Then
        mBDTabPer.Close
        mWSTabPer.Close
    End If
    Set frmPerfil = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub

Private Sub scboPerfil_Click()
    Call ProcuraPerfil(scboPerfil.Text)
End Sub

Private Sub scboPerfil_DropDown()
    Call CarregaPerfis
End Sub

Private Sub scboPerfil_GotFocus()
    Call CorDeFundo(scboPerfil, True)
    Call LimitaTexto(scboPerfil, 15)
End Sub

Private Sub scboPerfil_InitColumnProps()
    With scboPerfil
        .StyleSets.Add "Cabecalho"
        .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
        .StyleSets("Cabecalho").ForeColor = vbTitleBarText
        .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
        .StyleSets("Cabecalho").Font.Size = 9
        .StyleSets("Cabecalho").Font.Bold = True
    
        .BackColorOdd = dCorAmarelo
        .ListAutoPosition = True
        .ListAutoValidate = False
        .Font.Name = "MS Sans Serif"
        .Font.Size = 9
        .FieldSeparator = vbTab
        .DataMode = 2
        .MaxDropDownItems = 12
        .MinDropDownItems = 1
        .Text = "< Novo Perfil >"
        .DataFieldList = "Column 0"
        
        .Columns.Add 0    'PERFIL
        .Columns(0).Caption = "Perfil"
        .Columns(0).Width = .Width
        .Columns(0).HeadStyleSet = "Cabecalho"
        .Columns(0).Visible = True
    End With
End Sub

Private Sub scboPerfil_LostFocus()
    Call CorDeFundo(scboPerfil, False)
    If Len(Trim$(scboPerfil.Text)) = 0 Then scboPerfil.Text = "< Novo Perfil >"
End Sub

Private Sub sgrdMenu_Change()
    With sgrdMenu
        If .Redraw Then
            Call SelecionaSubOpcoes(CBool(.Columns(.Columns.Count - 1).Value))
        End If
    End With
End Sub

Private Sub sgrdMenu_InitColumnProps()
    With sgrdMenu
        .StyleSets.Add "Cabecalho"
        .StyleSets("Cabecalho").BackColor = vbActiveTitleBar
        .StyleSets("Cabecalho").ForeColor = vbTitleBarText
        .StyleSets("Cabecalho").Font.Name = "MS Sans Serif"
        .StyleSets("Cabecalho").Font.Size = 9
        .StyleSets("Cabecalho").Font.Bold = True
              
        .AllowDelete = False
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowColumnMoving = ssRelocateNotAllowed
        .AllowColumnShrinking = False
        .AllowColumnSizing = False
        .AllowColumnSwapping = ssRelocateNotAllowed
        .AllowGroupMoving = False
        .AllowGroupShrinking = False
        .AllowGroupSizing = False
        .BackColorOdd = dCorAmarelo
        .Caption = "Opções do Menu"
        .DefColWidth = 1000
        .Enabled = True
        .Font.Name = "MS Sans Serif"
        .Font.Size = 9
        .FieldSeparator = vbTab
        .ScrollBars = 3
        .SelectByCell = False
        .SelectTypeCol = 0
        .SelectTypeRow = 0
        .MaxSelectedRows = 100
    End With
End Sub

Private Sub CarregaSubMenu(ByVal cMenuID, ByVal iNivel)
    Dim iX, iNOpcoes, cMenuSubOpcao, cMenuSubDescricao
    
    iNivel = iNivel + 1
    iNOpcoes = fFrmMDIPrincipal.atbOpcoes.Tools(cMenuID).Menu.Tools.Count
    
    Call AcrescentaColuna(iNivel, "Nivel " & CStr(iNivel))
    
    'So redimensiona o tamanho da variavel para maior
    If iNivel > UBound(iLarguraColuna) Then ReDim Preserve iLarguraColuna(iNivel)
    
    For iX = 1 To iNOpcoes
        If fFrmMDIPrincipal.atbOpcoes.Tools(cMenuID).Menu.Tools(iX).Visible Then
            cMenuSubOpcao = fFrmMDIPrincipal.atbOpcoes.Tools(cMenuID).Menu.Tools(iX).ID
            If cMenuSubOpcao <> "separator" Then
                cMenuSubDescricao = cFiltra(fFrmMDIPrincipal.atbOpcoes.Tools(cMenuID).Menu.Tools(iX).Name, "&")
                sgrdMenu.AddItem String(iNivel, vbTab) & cMenuSubDescricao
                If TextWidth(cMenuSubDescricao) > iLarguraColuna(iNivel) Then iLarguraColuna(iNivel) = TextWidth(cMenuSubDescricao)
                iContador = iContador + 1
                ReDim Preserve cToolID(iContador)
                cToolID(iContador) = cMenuSubOpcao
                If fFrmMDIPrincipal.atbOpcoes.Tools(cMenuID).Menu.Tools(cMenuSubOpcao).Type = ssTypeMenu Then Call CarregaSubMenu(cMenuSubOpcao, iNivel)
            End If
        End If
    Next iX
End Sub

Private Sub AcrescentaColuna(ByVal iNColuna, ByRef cTitulo)
    With sgrdMenu
        .Redraw = False
        'Verifica se o numero de colunas existentes e menor que necessario
        If (.Columns.Count - 1) < iNColuna Then
            .Columns.Add iNColuna   'Opcão
            .Columns(iNColuna).Alignment = ssCaptionAlignmentLeft
            .Columns(iNColuna).Caption = cTitulo
            .Columns(iNColuna).DataType = 8 'Text
            .Columns(iNColuna).HeadStyleSet = "Cabecalho"
            .Columns(iNColuna).Locked = True
            .Columns(iNColuna).Visible = True
            .Columns(iNColuna).Width = 900
        End If
        .Redraw = True
    End With
End Sub

Private Sub CarregaMenu()
    Dim iX, iTotalOpcoes, iNivel, cMenuOpcao, cMenuDescricao
    
    iNivel = 0
    iContador = -1
    iTotalOpcoes = fFrmMDIPrincipal.atbOpcoes.ToolBars("Menu").Tools.Count
    
    Call AcrescentaColuna(iNivel, "Menu")
    
    
    ReDim Preserve iLarguraColuna(iNivel)
    iLarguraColuna(iNivel) = 1000
    
    sgrdMenu.Redraw = False
    For iX = 1 To iTotalOpcoes
        'Verifica se as opções estão visiveis
        If fFrmMDIPrincipal.atbOpcoes.ToolBars("Menu").Tools(iX).Visible Then
            cMenuOpcao = fFrmMDIPrincipal.atbOpcoes.ToolBars("Menu").Tools(iX).ID
            If cMenuOpcao <> "separator" Then
                cMenuDescricao = cFiltra(fFrmMDIPrincipal.atbOpcoes.ToolBars("Menu").Tools(iX).Name, "&")
                sgrdMenu.AddItem String(iNivel, vbTab) & cMenuDescricao
                If TextWidth(cMenuDescricao) > iLarguraColuna(iNivel) Then iLarguraColuna(iNivel) = TextWidth(cMenuDescricao)
                iContador = iContador + 1
                ReDim Preserve cToolID(iContador)
                cToolID(iContador) = cMenuOpcao
                If fFrmMDIPrincipal.atbOpcoes.ToolBars("Menu").Tools(cMenuOpcao).Type = ssTypeMenu Then Call CarregaSubMenu(cMenuOpcao, iNivel)
            End If
        End If
    Next iX
    'Vai actualizar a largura das coluna d grelha
    For iX = LBound(iLarguraColuna) To UBound(iLarguraColuna)
        sgrdMenu.Columns(iX).Width = iLarguraColuna(iX)
    Next iX
    sgrdMenu.Redraw = True
    'Acrescenta a grelha a coluna com a caixa para se selecionar
    Call AcrescentaUltimaColuna
End Sub

Private Sub AcrescentaUltimaColuna()
    Dim iNColuna
    With sgrdMenu
        .Redraw = False
        iNColuna = .Columns.Count
        .Columns.Add iNColuna    'Disponivel
        .Columns(iNColuna).Alignment = ssCaptionAlignmentCenter
        .Columns(iNColuna).Caption = "Visivel"
        .Columns(iNColuna).DataType = 11   'boolean
        .Columns(iNColuna).Style = ssStyleCheckBox
        .Columns(iNColuna).HeadStyleSet = "Cabecalho"
        .Columns(iNColuna).Visible = True
        .Columns(iNColuna).Width = 700
        .Redraw = True
    End With
End Sub

Private Sub SelecionaSubOpcoes(ByVal tPropriedade)
    Dim iNLinhas, vPonteiro, iLinha, iNivel, iColuna, bNivelAnterior
    'Se for a ultima coluna da grelha
    With sgrdMenu
        .Redraw = False
        vPonteiro = .Bookmark
        'Procura em que nivel esta a opcao
        iNivel = -1
        For iColuna = 0 To .Columns.Count - 2
            If Len(Trim$(.Columns(iColuna).Text)) > 0 Then iNivel = iColuna: Exit For
        Next iColuna
        'Sai do procedimento se não encontar nenhuma opção
        If iNivel < 0 Then Exit Sub
        
        iNLinhas = .Rows - 1
        For iLinha = (.AddItemRowIndex(.Bookmark)) + 1 To iNLinhas
            .Bookmark = .AddItemBookmark(iLinha)
            'Verifica se é uma opção do menu e do mesmo nivel
            If Len(Trim$(.Columns(iNivel).Text)) = 0 Then
                If iNivel > 0 Then
                    bNivelAnterior = False
                    For iColuna = 0 To iNivel
                        If Len(Trim$(.Columns(iColuna).Text)) > 0 Then bNivelAnterior = True: Exit For
                    Next iColuna
                    If bNivelAnterior Then Exit For
                End If
                .Columns(.Columns.Count - 1).Value = tPropriedade
            Else
                Exit For
            End If
        Next iLinha
        
        .Bookmark = vPonteiro
        
        'Anda para cima se a propriedade do true
        If tPropriedade Then
            For iColuna = iNivel - 1 To 0 Step -1
                For iLinha = (.AddItemRowIndex(.Bookmark)) - 1 To 0 Step -1
                    .Bookmark = .AddItemBookmark(iLinha)
                    If Len(Trim$(.Columns(iColuna).Text)) > 0 Then
                        .Columns(.Columns.Count - 1).Value = tPropriedade
                        Exit For
                    End If
                Next iLinha
            Next iColuna
        End If
        
        .Bookmark = vPonteiro
        .Redraw = True
    End With
End Sub

Private Sub GravaDadosNaBD()
    Dim iNTotallinhas, iNlinha, bSelecionado, iColuna, vPonteiro, recTABPERFIS As Recordset, cPerfil
    Dim mProcessamento As Processamento
    
    cPerfil = Trim$(scboPerfil.Text)
    
    If cPerfil = "< Novo Perfil >" Then GoTo SairDoProcedimento
    
    Set mProcessamento = New Processamento
    
    On Error GoTo TrataErro
    
    mWSTabPer.BeginTrans
    
    Set recTABPERFIS = mBDTabPer.OpenRecordset("TABPERFIS", dbOpenDynaset, dbSeeChanges, dbPessimistic)
    With recTABPERFIS
        If Not (.EOF And .BOF) Then
            .FindFirst "PERFIL='" & cPerfil & "'"
            If .NoMatch Then GoTo CriaNovoPerfil
        Else
CriaNovoPerfil:
            .AddNew
            .Fields("PERFIL") = cPerfil
            .Update
        End If
    End With
    
    mWSTabPer.CommitTrans
    
    recTABPERFIS.Close
    Set recTABPERFIS = Nothing
    
    iColuna = sgrdMenu.Columns.Count - 1
    mWSTabPer.BeginTrans
    'Apaga todas a opções existentes na tabela
    mBDTabPer.Execute "DELETE * FROM PERFIS WHERE PERFIL='" & cPerfil & "'", dbFailOnError
    iNTotallinhas = sgrdMenu.Rows - 1
    For iNlinha = 0 To iNTotallinhas
        vPonteiro = sgrdMenu.AddItemBookmark(iNlinha)
        bSelecionado = sgrdMenu.Columns(iColuna).CellValue(vPonteiro)
        If bSelecionado Then mBDTabPer.Execute "INSERT INTO PERFIS (PERFIL,NOMEMENU) VALUES('" & cPerfil & "','" & cToolID(iNlinha) & "')", dbFailOnError
    Next iNlinha
    mWSTabPer.CommitTrans
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSTabPer.Rollback
    Call ErrosGerais("Grava Perfil", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub CarregaPerfis()
    Dim recTABPERFIS As Recordset
    
    scboPerfil.Redraw = False
    scboPerfil.RemoveAll
    
    Set recTABPERFIS = mBDTabPer.OpenRecordset("SELECT * FROM TABPERFIS WHERE NIVEL >=" & gUtilizador.Nivel & " ORDER BY PERFIL ASC", dbOpenSnapshot, dbSeeChanges + dbReadOnly)
    With recTABPERFIS
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                scboPerfil.AddItem Trim$(.Fields("PERFIL"))
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recTABPERFIS = Nothing
    scboPerfil.Redraw = True
End Sub

Private Sub ProcuraPerfil(ByVal cPerfil)
    Dim recTABPERFIS As Recordset, cOpcao, iNOpcao, iColuna, vPonteiro
    Dim mProcessamento As Processamento
    
    cPerfil = Trim$(cPerfil)
   
    If cPerfil = "< Novo Perfil >" Then Exit Sub
    
    Set mProcessamento = New Processamento
    
    With sgrdMenu
        vPonteiro = .Bookmark
        iColuna = (.Columns.Count) - 1
        'Desseleciona todas as opções
        .Redraw = False
        For iNOpcao = 0 To sgrdMenu.Rows - 1
            .Bookmark = .AddItemBookmark(iNOpcao)
            .Columns(iColuna).Value = False
        Next iNOpcao
    End With
    
    Set recTABPERFIS = mBDTabPer.OpenRecordset("SELECT NOMEMENU FROM PERFIS WHERE PERFIL ='" & cPerfil & "' ORDER BY NOMEMENU ASC", dbOpenSnapshot, dbSeeChanges + dbReadOnly)
    With recTABPERFIS
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                cOpcao = CStr(Trim$(vFiltraCamposNulos(.Fields("NOMEMENU"))))
                For iNOpcao = LBound(cToolID) To UBound(cToolID)
                    If cToolID(iNOpcao) = cOpcao Then
                        sgrdMenu.Bookmark = sgrdMenu.AddItemBookmark(iNOpcao)
                        sgrdMenu.Columns(iColuna).Value = True
                        Exit For
                    End If
                Next iNOpcao
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set recTABPERFIS = Nothing
    
    sgrdMenu.Bookmark = vPonteiro
    sgrdMenu.Redraw = True
End Sub
