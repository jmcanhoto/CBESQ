VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCAIFCriarRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Criar Recibos das Mensalidades"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
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
   ScaleHeight     =   3300
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAno 
      Height          =   360
      Left            =   450
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1770
      Width           =   555
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2003"
      Top             =   2340
      Width           =   1200
   End
   Begin VB.CommandButton cmdCriar 
      Caption         =   "&Criar"
      Height          =   900
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   2340
      Width           =   1200
   End
   Begin VB.TextBox txtAtencao 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1200
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "CBESQ040.frx":0000
      Top             =   120
      Width           =   4200
   End
   Begin SSDataWidgets_B.SSDBCombo cboMes 
      Height          =   360
      Left            =   2385
      TabIndex        =   3
      Top             =   1770
      Width           =   1950
      _Version        =   196617
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
      Caption         =   "Indique o Ano"
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
      TabIndex        =   5
      Top             =   1500
      Width           =   1440
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
      Index           =   1
      Left            =   2385
      TabIndex        =   4
      Top             =   1500
      Width           =   1875
   End
End
Attribute VB_Name = "frmCAIFCriarRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSCriarRecibos As Workspace
Dim mBDCriarRecibos As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSCriarRecibos = DBEngine.CreateWorkspace("CAIFAlteracoes", gUtilizador.Nome, gUtilizador.Password)
    Set mBDCriarRecibos = mWSCriarRecibos.OpenDatabase(cBD_Path & cNomeBD)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Criar Recibos-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

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

Private Sub cmdCriar_Click()
    Dim mProcessamento As Processamento
    Dim qryCriarRecibos As QueryDef
    
    Set mProcessamento = New Processamento
    
    If bVerificaseExistemRecibos Then
        MsgBox " Já existem recibos do ano de " & txtAno.Text & vbCrLf & _
               " relativos ao mês de " & cboMes.Text, vbInformation + vbOKOnly, _
               "CAIF - Criar Recibos"
        GoTo SairDoProcedimento
    End If
    
    ' pede confirmação se quer continuar
    iRespMsgbox = MsgBox("Confirma que quer criar os recibos do ano de " & txtAno.Text & vbCrLf & _
                        "relativos ao mês de " & cboMes.Text, vbQuestion + vbYesNo, _
                        "Criar Recibos")
    ' se resposta não sai
    If iRespMsgbox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSCriarRecibos.BeginTrans
    
    Set qryCriarRecibos = mBDCriarRecibos.QueryDefs("CAIF RECIBOS Criar")
    
    qryCriarRecibos.Parameters("Ano") = txtAno.Text
    qryCriarRecibos.Parameters("Mes") = cboMes.Columns(1).Text
    qryCriarRecibos.Parameters("Utiliz") = gUtilizador.Nome
    
    qryCriarRecibos.Execute dbFailOnError
    mWSCriarRecibos.CommitTrans
    
    MsgBox "Criou os recibos relativos ao mês de " & cboMes.Text & _
        " do ano de " & txtAno.Text, vbInformation + vbOKOnly, "Criar Recibos"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSCriarRecibos.Rollback
    Call ErrosGerais("CAIF - Criar Recibos", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Call VerificaCriacao
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
    
    tBDAberta = tAbreBD
    
    txtAtencao.Text = " Atenção" & vbCrLf & _
                    " Vai criar recibos de Mensalidades." & vbCrLf & _
                    " Antes de continuar verifique se a Tabela de" & vbCrLf & _
                    " Alterações, está correctamente preenchida."
    
    ' verifica se existem recibos para tratar
    Call VerificaCriacao
    
    txtAno.Text = Year(Date$)
    
    ' carrega a combo
    Call CarregacboMes(cboMes)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSCriarRecibos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSCriarRecibos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCAIFCriarRecibos = Nothing
End Sub



Public Sub VerificaCriacao()
    Dim recRECIBOS As Recordset
    
    cSql = "SELECT * FROM RECIBOS_IDOSOS WHERE ISNULL(ESTADO_REC)"
    
    Set recRECIBOS = mBDCriarRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)

    ' se existirem recibos para fechar não deixa criar mais
    If Not (recRECIBOS.EOF And recRECIBOS.BOF) Then
        cmdCriar.Visible = False
    End If
    recRECIBOS.Close
    Set recRECIBOS = Nothing
End Sub
Private Sub CarregacboMes(ByRef Combo As SSDBCombo)
    Dim recTABMESES As Recordset
    Combo.Redraw = False
        
    Combo.RemoveAll
    cSql = "SELECT NOME,COD_MES FROM TABMESES ORDER BY COD_MES"
    
    Set recTABMESES = mBDCriarRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)
    
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



Public Function bVerificaseExistemRecibos()
    Dim recRECIBOSCRIADOS As Recordset
    
    cSql = "SELECT DISTINCT ANO,COD_MES FROM RECIBOS_IDOSOS"
    
    Set recRECIBOSCRIADOS = mBDCriarRecibos.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)

    recRECIBOSCRIADOS.FindFirst "ANO='" & txtAno.Text & "' AND COD_MES='" & cboMes.Columns(1).Text & "'"
    
    bVerificaseExistemRecibos = Not recRECIBOSCRIADOS.NoMatch
    
    recRECIBOSCRIADOS.Close
    Set recRECIBOSCRIADOS = Nothing
End Function
