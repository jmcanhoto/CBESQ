VERSION 5.00
Begin VB.Form frmSenha 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha de Acesso"
   ClientHeight    =   2115
   ClientLeft      =   3495
   ClientTop       =   2880
   ClientWidth     =   3900
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2115
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Tag             =   "1001"
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "2003"
      Top             =   1140
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "2002"
      Top             =   1140
      Width           =   1200
   End
   Begin VB.TextBox txtLogin 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1200
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   2550
   End
   Begin VB.TextBox txtLogin 
      Height          =   360
      Index           =   0
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   2550
   End
   Begin VB.Label lblAcesso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   240
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   675
      Width           =   900
   End
   Begin VB.Label lblAcesso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizador"
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   255
      Width           =   840
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    txtLogin(0).Text = NomeDoUtilizador
    Me.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim Form As Form
    For Each Form In Forms
        If Form Is Me Then Set Form = Nothing: Exit For
    Next
End Sub

Private Sub cmdOK_Click()
    If Len(Trim$(txtLogin(0).Text)) > 0 Then
        Call VerificaLogin(Trim$(txtLogin(0).Text), Trim$(txtLogin(1).Text), Trim$(NomeDoComputador))
    Else
        txtLogin(0).SetFocus
    End If
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub txtLogin_GotFocus(Index As Integer)
    Call CorDeFundo(txtLogin(Index), True)
    SelecionaTexto txtLogin(Index)
    cmdOK.Default = True
End Sub


Private Sub txtLogin_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtLogin_LostFocus(Index As Integer)
    Call CorDeFundo(txtLogin(Index), False)
    cmdOK.Default = False
End Sub



'Este procedimento vai verificar o utilizador e a respectiva password
'para aceder ao sistema
Private Sub VerificaLogin(ByVal cUtilizador, ByVal cPassword, ByVal cComputador)
    Dim mBD As Database, mrecTabLogin As Recordset, tReparou
    
    tReparou = False
    
    On Error Resume Next
    Set gwsInicial = DBEngine.CreateWorkspace(cUtilizador, cUtilizador, cPassword)
'    Set gwsInicial = DBEngine.CreateWorkspace(cUtilizador, "Admin", "")
    If Err = 3028 Then
        Call MsgBox("A Base de Dados encontra-se Bloqueada por outro utilizador", vbExclamation)
        End
    ElseIf Err = 3029 Then
        Call MsgBox("Utilizador ou Password invalida.", vbExclamation)
        txtLogin(0).SetFocus
        Exit Sub
    ElseIf Err = 3044 Then  'O path para o system.mda não e valido
        Call MsgBox("Não foi possivel encontrar o ficheiro de arranque.", vbExclamation)
        End
    ElseIf Err <> 0 Then
        Beep
    End If
    
    'Passa os valores para a variavel utilizador
    gUtilizador.Nome = cUtilizador
    gUtilizador.Password = cPassword
    
    On Error GoTo TrataErro
AbreBaseDeDados:
    'Workspaces.Append gwsInicial
    Set mBD = gwsInicial.OpenDatabase(cBD_Path & cNomeBD, False)
    Set mrecTabLogin = mBD.OpenRecordset("SELECT UTILIZADOR,COMPUTADOR FROM TABLOGIN", dbOpenDynaset, dbReadOnly)
    If Not (mrecTabLogin.EOF And mrecTabLogin.BOF) Then
        mrecTabLogin.FindFirst "UTILIZADOR='" & cUtilizador & "'"
        If mrecTabLogin.NoMatch Then
            mBD.Execute "INSERT INTO TABLOGIN(UTILIZADOR,COMPUTADOR,DATALTERA,HORALTERA) VALUES('" & cUtilizador & "','" & cComputador & "',#" & Date & "#,#" & Time & "#)", dbFailOnError
        Else
            If mrecTabLogin.Fields("COMPUTADOR") <> cComputador Then
                Call MsgBox("Este Utilizador já se encontra no sistema", vbExclamation)
                txtLogin(0).SetFocus
                Exit Sub
            Else
                mBD.Execute "UPDATE TABLOGIN SET DATALTERA=#" & Date & "#,HORALTERA=#" & Time & "#", dbFailOnError
            End If
        End If
    Else
        mBD.Execute "INSERT INTO TABLOGIN(UTILIZADOR,COMPUTADOR,DATALTERA,HORALTERA) VALUES('" & cUtilizador & "','" & cComputador & "',#" & Date & "#,#" & Time & "#)", dbFailOnError
    End If
    
    mrecTabLogin.Close
    Set mrecTabLogin = Nothing
    
    'Carrega os dados da empresa
    Call CarregaDadosEmpresa(mBD)
    
    'Verifica se o utilizador pertence á Gabingest
    If UCase(gUtilizador.Nome) = "JOCA" Then
        gUtilizador.Nivel = 0
        gUtilizador.Perfil = vbNullString
    Else
        'Carrega as dados do utilizador
        Call CarregaDadosUtilizador(mBD, cUtilizador)
    End If
    'Carrega os Titulos para os Mapas
    Call CarregaTitulosDoINI
        
    'Faz a copia da base dados TEMP para o nome do utilizador
    cBDComNomeUtilizador = cBD_Path & "\" & Trim$(gUtilizador.Nome) & ".MDB"
    'If Dir(cBDComNomeUtilizador) <> "" Then
    '    Call ApagaFicheiro(cBDComNomeUtilizador)
    'End If
    FileCopy cBD_Path & cNomeBDTemp, cBDComNomeUtilizador
    
    mBD.Close
    Set mBD = Nothing
    
    GoTo SairDoProcedimento
    
TrataErro:
    Select Case Err.Number
        Case 3043       'Falha de Disco ou de rede
            MsgBox "Problemas com o Disco ou com a Rede. Reinicie o Windows e esta Aplicação.", vbCritical + vbOKOnly
            End
        Case 3049       'Base de Dados está Corrupta
            'Tenta reparar a Base de Dados
            If Not tReparou Then
                DBEngine.RepairDatabase cBD_Path & cNomeBD
                tReparou = True
                GoTo AbreBaseDeDados
            Else
                Call ErrosGerais("Verifica Login", Err.Number, Err.Description)
                End
            End If
        Case Else       'Outros Erros que possam acontecer
            Call ErrosGerais("Verifica Login", Err.Number, Err.Description)
            End
        End Select
    
    
SairDoProcedimento:
    Me.Hide
    Showsplash
    Unload Me
End Sub

