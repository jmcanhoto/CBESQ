VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Begin VB.Form frmFichaAbsent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registo do Absentismo"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
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
   ScaleHeight     =   4740
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "2003"
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "2002"
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Frame fraProlHora 
      Caption         =   " Absentismo "
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
      Height          =   3435
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5205
      Begin SSDataWidgets_B.SSDBCombo cboHi 
         Height          =   345
         Left            =   1650
         TabIndex        =   4
         Top             =   540
         Width           =   630
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   1111
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboSector 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   1860
         Width           =   4890
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   8625
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtMotivo 
         Height          =   360
         Left            =   150
         MaxLength       =   30
         TabIndex        =   15
         Top             =   2595
         Width           =   4890
      End
      Begin GTMaskDate.GTMaskDate dcboData_Falta 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   540
         Width           =   1350
         _Version        =   65537
         _ExtentX        =   2381
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GTMaskNum.GTMaskNum txtMT 
         Height          =   330
         Left            =   4155
         TabIndex        =   11
         Top             =   1200
         Width           =   870
         _Version        =   65536
         _ExtentX        =   1535
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Format          =   "0000"
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         DisplayAlignment=   2
         Text            =   "0000"
         CalcDropDown    =   0   'False
         BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskAllowLeadZeros=   -1  'True
         MaskType        =   0
         DataType        =   2
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBCombo cboMi 
         Height          =   345
         Left            =   2475
         TabIndex        =   6
         Top             =   540
         Width           =   630
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   1111
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboHf 
         Height          =   345
         Left            =   3600
         TabIndex        =   8
         Top             =   540
         Width           =   630
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   1111
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B.SSDBCombo cboMf 
         Height          =   345
         Left            =   4425
         TabIndex        =   10
         Top             =   540
         Width           =   630
         _Version        =   196617
         DataMode        =   2
         Columns(0).Width=   3200
         _ExtentX        =   1111
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         Height          =   240
         Index           =   8
         Left            =   150
         TabIndex        =   14
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Sector"
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   12
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   240
         Index           =   5
         Left            =   4335
         TabIndex        =   9
         Top             =   540
         Width           =   45
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "às"
         Height          =   240
         Index           =   4
         Left            =   3600
         TabIndex        =   7
         Top             =   300
         Width           =   225
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   ":"
         Height          =   240
         Index           =   3
         Left            =   2355
         TabIndex        =   5
         Top             =   540
         Width           =   45
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "das"
         Height          =   240
         Index           =   1
         Left            =   1650
         TabIndex        =   3
         Top             =   300
         Width           =   345
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmFichaAbsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSAbsent As Workspace
Dim mBDAbsent As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lContador
Dim lNumFunc
Dim cInstituicao



Public Sub CarregacboHoras(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABHORA As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT COD_HORA,DESCRICAO FROM TABHORA ORDER BY COD_HORA"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABHORA = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABHORA.EOF And Not recTABHORA.BOF Then
        While Not recTABHORA.EOF
            Combo.AddItem recTABHORA!COD_HORA & vbTab & recTABHORA!DESCRICAO
            recTABHORA.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABHORA.Close
    Set recTABHORA = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub
Public Sub CarregacboMinutos(ByRef Combo As SSDBCombo)
    ' defenição de variaveis
    Dim mBDBaseDados As Database
    Dim recTABMINUTO As Recordset
    Dim cSql
    
    Combo.Redraw = False
    ' remove tudo o que a combo tiver
    Combo.RemoveAll
    ' atribui o SQL a variavel de seleção de Dados
    cSql = "SELECT COD_MINUTO,DESCRICAO FROM TABMINUTO ORDER BY COD_MINUTO"
    ' abre a base de Dados
    Set mBDBaseDados = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    ' abre a tabela
    Set recTABMINUTO = mBDBaseDados.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    'carrega a combo
    If Not recTABMINUTO.EOF And Not recTABMINUTO.BOF Then
        While Not recTABMINUTO.EOF
            Combo.AddItem recTABMINUTO!COD_MINUTO & vbTab & recTABMINUTO!DESCRICAO
            recTABMINUTO.MoveNext
        Wend
    End If
    ' fecha a tabela
    recTABMINUTO.Close
    Set recTABMINUTO = Nothing
    ' fecha a base de dados
    mBDBaseDados.Close
    Set mBDBaseDados = Nothing
    
    Combo.Redraw = True
End Sub

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSAbsent = DBEngine.CreateWorkspace(cBotao & "Absent", gUtilizador.Nome, gUtilizador.Password)
    Set mBDAbsent = mWSAbsent.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Ficha Absent-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cboHf_InitColumnProps()
    With cboHf
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
        .Columns(0).Caption = "H."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Caption = "Descrição"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).Width = 3000
    End With
End Sub

Private Sub cboHf_LostFocus()
    txtMT.Text = DateDiff("n", cboHi.Text & ":" & cboMi.Text, cboHf.Text & ":" & cboMf.Text)
End Sub


Private Sub cboHi_InitColumnProps()
    With cboHi
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
        .Columns(0).Caption = "H."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Caption = "Descrição"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).Width = 3000
    End With
End Sub

Private Sub cboHi_LostFocus()
    txtMT.Text = DateDiff("n", cboHi.Text & ":" & cboMi.Text, cboHf.Text & ":" & cboMf.Text)
End Sub


Private Sub cboMf_InitColumnProps()
    With cboMf
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
        .Columns(0).Caption = "M."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Caption = "Descrição"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).Width = 3000
    End With
End Sub

Private Sub cboMf_LostFocus()
    txtMT.Text = DateDiff("n", cboHi.Text & ":" & cboMi.Text, cboHf.Text & ":" & cboMf.Text)
End Sub


Private Sub cboMi_InitColumnProps()
    With cboMi
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
        .Columns(0).Caption = "M."
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
        ' coluna 1
        .Columns.Add 1
        .Columns(1).Caption = "Descrição"
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
        .Columns(1).Width = 3000
    End With
End Sub


Private Sub cboMi_LostFocus()
    txtMT.Text = DateDiff("n", cboHi.Text & ":" & cboMi.Text, cboHf.Text & ":" & cboMf.Text)
End Sub


Private Sub cboSector_DropDown()
    ' carrega a combo
    Call CarregacboSector(cboSector)
End Sub

Private Sub cboSector_InitColumnProps()
    With cboSector
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
        .Columns(0).Caption = "Sector"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryAbsentismo As QueryDef
    
    Set mProcessamento = New Processamento
    
On Error GoTo TrataErro
    ' começa a transação
    mWSAbsent.BeginTrans
    If cBotao = "Novo" Then
        Set qryAbsentismo = mBDAbsent.QueryDefs("ABSENTISMO Insere")
        ' parametros de input
        qryAbsentismo.Parameters("NumFunc") = lNumFunc
        qryAbsentismo.Parameters("Cod_Inst") = cCodificaInstituicao(cInstituicao)
        qryAbsentismo.Parameters("Data_Falta") = dcboData_Falta.DateValue
        qryAbsentismo.Parameters("Hora_Ini") = cboHi.Text
        qryAbsentismo.Parameters("Min_Ini") = cboMi.Text
        qryAbsentismo.Parameters("Hora_Fim") = cboHf.Text
        qryAbsentismo.Parameters("Min_Fim") = cboMf.Text
        qryAbsentismo.Parameters("Min_Total") = txtMT.Text
        qryAbsentismo.Parameters("Sector") = cCodificaSector(cboSector.Text)
        qryAbsentismo.Parameters("Motivo") = txtMotivo.Text
        qryAbsentismo.Parameters("Utiliz") = gUtilizador.Nome
      ElseIf cBotao = "Altera" Then
        Set qryAbsentismo = mBDAbsent.QueryDefs("ABSENTISMO Altera")
        ' parametros de input
        qryAbsentismo.Parameters("Contador") = lContador
        qryAbsentismo.Parameters("Data_Falta") = dcboData_Falta.DateValue
        qryAbsentismo.Parameters("Hora_Ini") = cboHi.Text
        qryAbsentismo.Parameters("Min_Ini") = cboMi.Text
        qryAbsentismo.Parameters("Hora_Fim") = cboHf.Text
        qryAbsentismo.Parameters("Min_Fim") = cboMf.Text
        qryAbsentismo.Parameters("Min_Total") = txtMT.Text
        qryAbsentismo.Parameters("Sector") = cCodificaSector(cboSector.Text)
        qryAbsentismo.Parameters("Motivo") = txtMotivo.Text
        qryAbsentismo.Parameters("Utiliz") = gUtilizador.Nome
    End If
    ' executa a query
    qryAbsentismo.Execute dbFailOnError
    
    mWSAbsent.CommitTrans
    ' faz o refresh da frmAbsentismo
    frmAbsentismo.datAbsentismo.Refresh
    frmAbsentismo.sgrdAbsentismo.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSAbsent.Rollback
    Call ErrosGerais(cBotao & " Absentismo", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Unload Me
End Sub

Private Sub dcboData_Falta_DropDown()
    If Not IsDate(dcboData_Falta.Text) Then
        dcboData_Falta.DateValue = Date
    End If
End Sub


Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    cBotao = cBotaoOrigem
    If cBotao <> "Novo" Then
        lContador = CLng(frmAbsentismo.sgrdAbsentismo.Columns(5).Value)
    ElseIf cBotao = "Novo" Then
        lNumFunc = CLng(frmAbsentismo.cboNumFunc.Text)
        cInstituicao = frmAbsentismo.cboInstituicao.Text
    End If

    If cBotao = "Novo" Then
        Me.Caption = "Novo " & Me.Caption
    ElseIf cBotao = "Altera" Then
        Me.Caption = "Alteração do " & Me.Caption
    End If

    CenterMe Me
    LoadResStrings Me
        
    Call AlteraWindowList(Me.Caption)
    
    tBDAberta = tAbreBD
    
    Call CarregacboHoras(cboHi)
    Call CarregacboHoras(cboHf)
    Call CarregacboMinutos(cboMi)
    Call CarregacboMinutos(cboMf)
    
    Call CamposLimpaCarrega
End Sub
Public Sub CamposLimpaCarrega()
    Dim recABSENTISMO As Recordset
    ' Ficha
    If cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM ABSENTISMO WHERE CONTADOR=" & lContador
        Set recABSENTISMO = mBDAbsent.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        'Poe os campos com os dados do Absentismo
        dcboData_Falta.Text = vFiltraCamposNulos(recABSENTISMO!DATA_FALTA)
        cboHi.Text = vFiltraCamposNulos(recABSENTISMO!HORA_INI)
        cboMi.Text = vFiltraCamposNulos(recABSENTISMO!MIN_INI)
        cboHf.Text = vFiltraCamposNulos(recABSENTISMO!HORA_FIM)
        cboMf.Text = vFiltraCamposNulos(recABSENTISMO!MIN_FIM)
        txtMT.Text = vFiltraCamposNulos(recABSENTISMO!MIN_TOTAL)
        cboSector.Text = cDescodificaSector(recABSENTISMO!SECTOR)
        txtMotivo.Text = vFiltraCamposNulos(recABSENTISMO!MOTIVO_FALTA)
        ' fecha o recordset
        recABSENTISMO.Close
        Set recABSENTISMO = Nothing
    ElseIf cBotao = "Novo" Then
        'Poe os campos preparados para nova ficha
        dcboData_Falta.Text = Date
        cboHi.Text = "00"
        cboMi.Text = "00"
        cboHf.Text = "00"
        cboMf.Text = "00"
        txtMT.Text = "0"
        cboSector.Text = vbNullString
        txtMotivo.Text = vbNullString
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSAbsent.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSAbsent = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub

