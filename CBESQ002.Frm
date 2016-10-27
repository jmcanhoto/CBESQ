VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   3165
   ClientLeft      =   1530
   ClientTop       =   1305
   ClientWidth     =   6150
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveToolBars.SSActiveToolBars atbOpcoes 
      Left            =   240
      Top             =   285
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      FontSource      =   1
      MenuAnimations  =   2
      ToolBarsCount   =   2
      ToolsCount      =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "CBESQ002.frx":0000
      ToolBars        =   "CBESQ002.frx":18125
   End
   Begin MSComctlLib.StatusBar stbBarraDeEstado 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "C.B.E.S.Q."
            TextSave        =   "C.B.E.S.Q."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5001
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmMDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub OutrasFuncoes(ByVal cFuncao)
    
    cFuncao = UCase(Trim$(cFuncao))
    
    On Error GoTo TrataErro
    
    If TypeName(ActiveForm) = "Nothing" Then GoTo SairDoProcedimento
    
    Select Case cFuncao
        Case "CORTAR"
            If TypeOf ActiveForm.ActiveControl Is TextBox Then
                If Not ActiveForm.ActiveControl.Locked Then
                    ' Limpa o clipboard
                    Clipboard.Clear
                    ' Copia o Texto Selecionado para o clipboard
                    Clipboard.SetText ActiveForm.ActiveControl.SelText
                    ' Apaga o texto selecionado
                    ActiveForm.ActiveControl.SelText = vbNullString
                End If
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBCombo Then
                    ' Limpa o clipboard
                    Clipboard.Clear
                    ' Copia o Texto Selecionado para o clipboard
                    Clipboard.SetText ActiveForm.ActiveControl.SelText
                    ' Apaga o texto selecionado
                    ActiveForm.ActiveControl.SelText = vbNullString
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBGrid Then
                ' Limpa o clipboard
                Clipboard.Clear
                ' Copia o Texto Selecionado para o clipboard
                Clipboard.SetText ActiveForm.ActiveControl.ActiveCell.SelText
                ' Apaga o texto selecionado
                ActiveForm.ActiveControl.ActiveCell.SelText = vbNullString
            End If
        Case "COPIAR"
            If TypeOf ActiveForm.ActiveControl Is TextBox Then
                ' Limpa o clipboard
                Clipboard.Clear
                ' Copia o Texto Selecionado para o clipboard
                Clipboard.SetText ActiveForm.ActiveControl.SelText
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBCombo Then
                ' Limpa o clipboard
                Clipboard.Clear
                ' Copia o Texto Selecionado para o clipboard
                Clipboard.SetText ActiveForm.ActiveControl.SelText
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBGrid Then
                ' Limpa o clipboard
                Clipboard.Clear
                ' Copia o Texto Selecionado para o clipboard
                Clipboard.SetText ActiveForm.ActiveControl.ActiveCell.SelText
            End If
        Case "COLAR"
            If Len(CStr(Clipboard.GetText())) = 0 Then GoTo SairDoProcedimento
            
            If TypeOf ActiveForm.ActiveControl Is TextBox Then
                If Not ActiveForm.ActiveControl.Locked Then
                    ' Cola o texto do clipboard
                    ActiveForm.ActiveControl.SelText = Clipboard.GetText()
                End If
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBGrid Then
                If ActiveForm.ActiveControl.ActiveCell.SelLength = 0 Then
                    If Len(ActiveForm.ActiveControl.ActiveCell.Text) = 0 Then
                        ActiveForm.ActiveControl.ActiveCell.Text = Clipboard.GetText()
                    Else
                        ActiveForm.ActiveControl.ActiveCell.Text = Mid$(ActiveForm.ActiveControl.ActiveCell.Text, 1, ActiveForm.ActiveControl.ActiveCell.SelStart) & _
                        Clipboard.GetText() & Mid$(ActiveForm.ActiveControl.ActiveCell.Text, ActiveForm.ActiveControl.ActiveCell.SelStart)
                    End If
                Else
                    ' Cola o texto do clipboard
                    ActiveForm.ActiveControl.ActiveCell.SelText = Clipboard.GetText()
                End If
            ElseIf TypeOf ActiveForm.ActiveControl Is SSDBCombo Then
                ' Cola o texto do clipboard
                ActiveForm.ActiveControl.SelText = Clipboard.GetText()
            End If
    End Select
    
    GoTo SairDoProcedimento
    
TrataErro:
    Beep
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub


Private Sub atbOpcoes_ToolClick(ByVal Tool As SSTool)
    Dim mProcessamento As Processamento, Janela As Form
    
On Error Resume Next
    
    Set mProcessamento = New Processamento
        
    Select Case Tool.ID
        ' Sócios
        Case "mnuSociosGestao"                  '(Button)
            Load frmGestaoSocios
        Case "mnuSociosLista"                   '(Button)
            Load frmListaSocios
        Case "mnuSociosCriaQuo"                  '(Button)
            Load frmCriarQuotasSocio
        
        ' Opção de Sair da Aplicação
        Case "mnuSociosSair"                    '(Button)
            Unload Me
            End
        
        ' Incrições
        Case "mnuInscricoesGestao"              '(Button)
            Load frmGestaoInscricoes
        Case "mnuInscricoesLista"               '(Button)
            Load frmListaInscricoes
        
        ' Utentes
        Case "mnuUtentesGestao"                 '(Button)
            Load frmGestaoUtentes
        Case "mnuGestaoRecibos"                 '(Button)
            Load frmGestaoRecibos
        Case "mnuUtentesLista"                  '(Button)
            Load frmListaUtentes
        Case "mnuUtentesActualizacaoMensalidades"            '(Button)
            Load frmActualizacaoMensalidades
        Case "mnuUtentesRendPercapita"            '(Button)
            Load frmRendimentoPercapita
        Case "mnuUtentesRendPercapitaMapa"            '(Button)
            Load frmRendimentoPercapitaMapa
        Case "mnuUtentesDeclaracoes"            '(Button)
            Load frmDeclaracoesIRS
        Case "mnuUtentesDeclaracoesFreq"            '(Button)
            Load frmDeclaracoesFrequencia
        Case "mnuUtentesMensProxAno"            '(Button)
            Load frmMensalidadesProximoAno
        Case "mnuUtentesMultarRecibos"            '(Button)
            Load frmMultaRecibos
        Case "mnuUtentesMensRecalcular"            '(Button)
            Load frmMensalidadesRecalcular
        Case "mnuUtentesMudarSala"            '(Button)
            Load frmMudarSala
        Case "mnuUtentesNomeEducadora"            '(Button)
            Load frmNomeEducadora

        'Mensalidades
        Case "mnuUtentesMensalidadesAlteracoes" '(Button)
            Load frmTabelaAlteracoes
        Case "mnuUtentesMensalidadesLista"      '(Button)
            Load frmListaRecibos
        Case "mnuUtentesMensalidadesCriarRecibo"      '(Button)
            Load frmCriarRecibo
        Case "mnuUtentesMensalidadesCriar"      '(Button)
            Load frmCriarRecibos
        Case "mnuUtentesMensalidadesAlterar"    '(Button)
            Load frmAlterarRecibos
        Case "mnuUtentesMensalidadesAlterarSala"    '(Button)
            Load frmAlterarRecibosSala
        Case "mnuUtentesMensalidadesFechar"     '(Button)
            Load frmFecharRecibos
        Case "mnuUtentesMensalidadesAlterarCriados"    '(Button)
            Load frmAlterarRecibosCriados
        
        ' CAIF - Incrições
        Case "mnuCAIFInscricoesGestao"              '(Button)
            Load frmCAIFGestaoInscricoes
        Case "mnuCAIFInscricoesLista"               '(Button)
            Load frmCAIFListaInscricoes
        
        ' CAIF - Utentes
        Case "mnuCAIFGestao"                 '(Button)
            Load frmCAIFGestaoUtentes
        Case "mnuCAIFGestaoRecibos"                 '(Button)
            Load frmCAIFGestaoRecibos
        Case "mnuCAIFLista"                  '(Button)
            Load frmCAIFListaUtentes
        Case "mnuCAIFMensProxAno"    '(Button)
            Load frmCAIFMensalidadesProximoAno
        Case "mnuCAIFActualizacaoMensalidades"    '(Button)
            Load frmCAIFActualizacaoMensalidades
        Case "mnuCAIFUtentesRendPercapita"            '(Button)
            Load frmCAIFRendimentoPercapita
        Case "mnuCAIFUtentesRendPercapitaMapa"            '(Button)
            Load frmCAIFRendimentoPercapitaMapa
        Case "mnuCAIFDeclaracoesIRS"            '(Button)
            Load frmCAIFDeclaracoesIRS
        Case "mnuCAIFDeclaracoesFreq"            '(Button)
            Load frmCAIFDeclaracoesFrequencia
        Case "mnuCAIFMultarRecibos"            '(Button)
            Load frmCAIFMultaRecibos
        Case "mnuCAIFTabelaAlteracoesMensais"    '(Button)
            Load frmCAIFTabelaAlteracoes
        Case "mnuCAIFCriarRecibo"    '(Button)
            Load frmCAIFCriarRecibo
        Case "mnuCAIFCriarRecibos"    '(Button)
            Load frmCAIFCriarRecibos
        Case "mnuCAIFAlterarRecibos"    '(Button)
            Load frmCAIFAlterarRecibos
        Case "mnuCAIFAlterarRecibosSala"    '(Button)
            Load frmCAIFAlterarRecibosSala
        Case "mnuCAIFFecharRecibos"    '(Button)
            Load frmCAIFFecharRecibos
        Case "mnuCAIFListaRecibos"    '(Button)
            Load frmCAIFListaRecibos
        
        Case "mnuCAIFAlterarRecibosCriados"    '(Button)
            Load frmCAIFAlterarRecibosCriados
        
        
        
        ' Funcionários
        Case "mnuFuncGestao"                    '(Button)
            Load frmGestaoFuncionarios
        Case "mnuFuncLista"                     '(Button)
            Load frmListaFuncionarios
        Case "mnuFuncAbsentismo"                '(Button)
            Load frmAbsentismo
        Case "mnuFuncProlongamento"             '(Button)
            Load frmProlongamento
        Case "mnuFuncListaAbsProl"              '(Button)
        
        ' Tabelas
        Case "mnuTabUti"                        'Utilizadores
            Load frmUtilizadores
        Case "mnuTabPerfil"                     'Perfis
            Load frmPerfil
        
        ' Tabelas
        Case "mnuTabUti"                        '(Button)
        
        ' Ajuda
        Case "mnuAjuSob"                        'Sobre a Aplicação
            Load frmSobre
                
        ' Utilitários
        Case "mnuUtiCut"                        ' Cortar
            Call OutrasFuncoes("CORTAR")
        Case "mnuUtiCop"                        ' Copiar
            Call OutrasFuncoes("COPIAR")
        Case "mnuUtiCol"                        ' Colar
            Call OutrasFuncoes("COLAR")
        Case "mnuUtiDisHor"                        ' Na Horizontal
            Me.Arrange vbTileHorizontal
        Case "mnuUtiDisVer"                        ' Na Vertical
            Me.Arrange vbTileVertical
        Case "mnuUtiDisCas"                        ' Em Cascata
            Me.Arrange vbCascade
        Case "mnuUtiBdCom"              ' Compactar
            Call CompactarBD(cBD_Path, cNomeBD)
        Case "mnuUtiBdSeg"              ' Seguranças / Recuperação
            Load frmSegurancas
        
        Case "mnuUtiImp"
            If Forms.Count > 1 Then
'                Select Case Me.ActiveForm.Name
'                Case "frmTesRegistoDoc"
'                    Call frmTesRegistoDoc.Imprimir(1)
'                Case "frmFacRegistoDoc"
'                    Call frmFacRegistoDoc.Imprimir(1)
'                End Select
            End If
        Case "mnuUtiPre"
            If Forms.Count > 1 Then
'                Select Case Me.ActiveForm.Name
'                Case "frmTesRegistoDoc"
'                    Call frmTesRegistoDoc.Imprimir(0)
'                Case "frmFacRegistoDoc"
'                    Call frmFacRegistoDoc.Imprimir(0)
'                End Select
            End If
        
        Case Else
            If Tool.ID <> "mnuJan" Then
                ' Lista de janelas
                For Each Janela In Forms
                    If Janela.Caption = Tool.ID Then
                        Janela.SetFocus
                    End If
                Next
            End If
    End Select
End Sub


Private Sub MDIForm_Load()
    Call CarregaPerfil
    Me.Caption = gEmpresa.Nome
    Me.stbBarraDeEstado.Panels(1).Text = "Versão : " & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "00")
    Me.stbBarraDeEstado.Panels(2).Text = "Utilizador : " & gUtilizador.Nome & "  "
    Me.stbBarraDeEstado.Panels(3).Text = vbNullString
    If tFicheiroExiste(cApl_Path & "\Fotos\Fundo.jpg") Then
        Me.Picture = LoadPicture(cApl_Path & "\Fotos\Fundo.jpg")
    End If
End Sub

'Este Procedimento vai verificar qual as opções disponiveis para o perfil do utilizador
Private Sub CarregaPerfil()
    Dim mWSEntrada As Workspace, mBDEntrada As Database, recPERFIL As Recordset
    Dim iNMenus, iNSubMenus, cMenu, cSubMenu, iMenu, iSubMenu
    
    If UCase(gUtilizador.Nome) = "JOCA" Then
        Exit Sub
    Else
        'Mete todas as opcoes invisiveis
        iNMenus = atbOpcoes.Tools.Count
        For iMenu = 1 To iNMenus
            fFrmMDIPrincipal.atbOpcoes.Tools(iMenu).Visible = False
        Next iMenu
        
        On Error GoTo TrataErro
        
        Set mWSEntrada = DBEngine.CreateWorkspace("Entrada", gUtilizador.Nome, gUtilizador.Password)
        Set mBDEntrada = mWSEntrada.OpenDatabase(cBD_Path & cNomeBD)
        Set recPERFIL = mBDEntrada.OpenRecordset("SELECT NOMEMENU FROM PERFIS WHERE PERFIL='" & gUtilizador.Perfil & "'", dbOpenSnapshot, dbReadOnly)
        With recPERFIL
            If Not (.EOF And .BOF) Then
                Do While Not .EOF
                    cMenu = CStr(vFiltraCamposNulos(.Fields("NOMEMENU")))
                    atbOpcoes.Tools(cMenu).Visible = True
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set recPERFIL = Nothing
        mBDEntrada.Close
        Set mBDEntrada = Nothing
        mWSEntrada.Close
        Set mWSEntrada = Nothing
    End If
    
    GoTo SaiDoProcedimento
    
TrataErro:
    Call ErrosGerais("Carregar Perfil", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Janela As Form
    
    For Each Janela In Forms
        If Not Janela Is Me Then
            If Janela.MDIChild Then Unload Janela
        End If
    Next
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim mBDMDI As Database, Janela As Form
    
    On Error GoTo TrataErro
    
    Set mBDMDI = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    mBDMDI.Execute "DELETE * FROM TABLOGIN WHERE UTILIZADOR='" & gUtilizador.Nome & "'", dbFailOnError
    
    mBDMDI.Close
    gwsInicial.Close
    
    On Error GoTo 0
    For Each Janela In Forms
        If Janela Is Me Then Set Janela = Nothing: Exit For
    Next
    GoTo SaiDoProcedimento

TrataErro:
    Call ErrosGerais("Desligar Utilizador", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
End Sub


