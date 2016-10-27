VERSION 5.00
Begin VB.Form frmCriarQuotasSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Criar Quotas de Sócios"
   ClientHeight    =   2505
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
   ScaleHeight     =   2505
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAno 
      Height          =   360
      Left            =   450
      MaxLength       =   4
      TabIndex        =   4
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
      Top             =   1485
      Width           =   1200
   End
   Begin VB.CommandButton cmdCriar 
      Caption         =   "&Criar"
      Height          =   900
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2002"
      Top             =   1485
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
      Text            =   "CBESQ021.frx":0000
      Top             =   120
      Width           =   4200
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
      TabIndex        =   3
      Top             =   1500
      Width           =   1440
   End
End
Attribute VB_Name = "frmCriarQuotasSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSCriarQuotas As Workspace
Dim mBDCriarQuotas As Database
Dim mBDCriarQuotasTemp As Database

Dim tBDAberta

Dim cSql
Dim iRespMsgbox
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSCriarQuotas = DBEngine.CreateWorkspace("CriarQuotas", gUtilizador.Nome, gUtilizador.Password)
    Set mBDCriarQuotas = mWSCriarQuotas.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDCriarQuotasTemp = mWSCriarQuotas.OpenDatabase(cBDComNomeUtilizador)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Criar Quotas-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCriar_Click()
    Dim mProcessamento As Processamento
    Dim qryCriarQuotas As QueryDef
    Dim cNomeMapa
    
    Set mProcessamento = New Processamento
    
    If bVerificaseExistemQuotas Then
        MsgBox " Já existem Quotas do ano de " & txtAno.Text, vbInformation + vbOKOnly, _
               "Criar Quotas"
        GoTo SairDoProcedimento
    End If
    
    ' pede confirmação se quer continuar
    iRespMsgbox = MsgBox("Confirma que quer criar os Quotas do ano de " & txtAno.Text, vbQuestion + vbYesNo, _
                        "Criar Quotas")
    ' se resposta não sai
    If iRespMsgbox = vbNo Then
        GoTo SairDoProcedimento
    End If
    
On Error GoTo TrataErro
    mWSCriarQuotas.BeginTrans
    
    Set qryCriarQuotas = mBDCriarQuotas.QueryDefs("QUOTAS Criar")
        
    qryCriarQuotas.Parameters("Ano") = txtAno.Text
    qryCriarQuotas.Parameters("Trimestre") = "1"
    qryCriarQuotas.Parameters("Utiliz") = gUtilizador.Nome
    
    qryCriarQuotas.Execute dbFailOnError
    
    qryCriarQuotas.Parameters("Ano") = txtAno.Text
    qryCriarQuotas.Parameters("Trimestre") = "2"
    qryCriarQuotas.Parameters("Utiliz") = gUtilizador.Nome
    
    qryCriarQuotas.Execute dbFailOnError
    
'    qryCriarQuotas.Parameters("Ano") = txtAno.Text
'    qryCriarQuotas.Parameters("Trimestre") = "3"
'    qryCriarQuotas.Parameters("Utiliz") = gUtilizador.Nome
'
'    qryCriarQuotas.Execute dbFailOnError
'
'    qryCriarQuotas.Parameters("Ano") = txtAno.Text
'    qryCriarQuotas.Parameters("Trimestre") = "4"
'    qryCriarQuotas.Parameters("Utiliz") = gUtilizador.Nome
'
'    qryCriarQuotas.Execute dbFailOnError
    
    mWSCriarQuotas.CommitTrans
    
    ' emitiu o mapa como deve ser
    MsgBox "Criou os Quotas relativas ao ano de " & txtAno.Text, vbInformation + vbOKOnly, "Criar Quotas"
    
    GoTo SairDoProcedimento

TrataErro:
    mWSCriarQuotas.Rollback
    Call ErrosGerais("Criar Quotas", Err.Number, Err.Description)
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
    
    tBDAberta = tAbreBD
    
    txtAtencao.Text = " Atenção" & vbCrLf & _
                    " Vai criar Quotas dos Sócios."
    
    
    txtAno.Text = Year(Date$)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSCriarQuotas.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSCriarQuotas = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmCriarQuotasSocio = Nothing
End Sub





Public Function bVerificaseExistemQuotas()
    Dim recQUOTASCRIADAS As Recordset
    
    cSql = "SELECT DISTINCT ANO FROM QUOTAS"
    
    Set recQUOTASCRIADAS = mBDCriarQuotas.OpenRecordset(cSql, dbOpenDynaset, dbReadOnly)

    recQUOTASCRIADAS.FindFirst "ANO='" & txtAno.Text & "'"
    
    bVerificaseExistemQuotas = Not recQUOTASCRIADAS.NoMatch
    
    recQUOTASCRIADAS.Close
    Set recQUOTASCRIADAS = Nothing
End Function
