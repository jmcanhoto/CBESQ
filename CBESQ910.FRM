VERSION 5.00
Begin VB.Form frmSegurancas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segurança / Reposição dos Dados do Sistema"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ClipControls    =   0   'False
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
   ScaleHeight     =   4290
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "2003"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Frame fraSeleccionar 
      Caption         =   " Selecione o Destino para a Segurança "
      ClipControls    =   0   'False
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
      Height          =   2985
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6840
      Begin VB.TextBox txtPath 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   6300
      End
      Begin VB.FileListBox filPath 
         Height          =   2010
         Left            =   3855
         TabIndex        =   4
         Top             =   750
         Width           =   2805
      End
      Begin VB.DriveListBox drvPath 
         Height          =   360
         Left            =   330
         TabIndex        =   3
         Top             =   2400
         Width           =   3510
      End
      Begin VB.DirListBox dirPath 
         Height          =   1440
         Left            =   330
         TabIndex        =   2
         Top             =   750
         Width           =   3510
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "2002"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Frame fraImpExp 
      ClipControls    =   0   'False
      Height          =   795
      Left            =   150
      TabIndex        =   5
      Top             =   3165
      Width           =   3855
      Begin VB.OptionButton optSegRec 
         Caption         =   "Segurança"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   330
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.OptionButton optSegRec 
         Caption         =   "Reposição"
         Height          =   300
         Index           =   1
         Left            =   2100
         TabIndex        =   7
         Top             =   330
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmSegurancas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cCaminho, cNomeDoFicheiro, iCiclo, mProcessamento As Processamento
    
    Set mProcessamento = New Processamento
    
    On Error Resume Next
    
    cCaminho = dirPath
    If CStr(Right(cCaminho, 1)) <> "\" Then cCaminho = cCaminho & "\"
    
    'Se for para fazer uma segurança
    If optSegRec(0).Value Then
        cNomeDoFicheiro = "SG_" & CStr(Format$(Date, "yyyymmdd")) & "_00" & ".mdb"
        If tFicheiroExiste((cBD_Path & cNomeBD)) Then
            Do While True
                iCiclo = iCiclo + 1
                If Not tFicheiroExiste((cCaminho & cNomeDoFicheiro)) Then Exit Do
                cNomeDoFicheiro = "SG_" & CStr(Format$(Date, "yyyymmdd")) & "_" & Format$(iCiclo, "00") & ".mdb"
            Loop
            DBEngine.CompactDatabase (cBD_Path & cNomeBD), (cCaminho & cNomeDoFicheiro), dbLangGeneral
            MsgBox "Segurança da Base de Dados concluida com Sucesso.", vbInformation + vbOKOnly, "Segurança da Base de Dados"
            filPath.Refresh
        Else
            MsgBox "Não foi encontrada a Base de Dados.", vbInformation + vbOKOnly, "Segurança da Base de Dados"
        End If
    Else    'Se for uma recuperação da Base de Dados
        'Verifica se existe algum utilizador no sistema
        If DBEngine.Workspaces.Count > 1 Then
            MsgBox "Só pode efectuar esta operação se não houver mais nenhum utilizador no sistema.", vbInformation + vbOKOnly, "Recuperação da Base de Dados"
        Else
            If MsgBox("A Base de Dados actual vai ser subtituida." & vbCrLf & _
                        "Tem a certeza que deseja continuar ?", vbYesNo + vbQuestion + vbDefaultButton2, _
                        "Recuperação da Base de Dados") = vbNo Then Exit Sub
            
            cNomeDoFicheiro = "\SG" & CStr(Format$(Date, "yyyymmdd")) & ".mdb"
            iCiclo = 0
            'Renomeia a BD para SG+Data_Do_Sistema
            Do While True
                iCiclo = iCiclo + 1
                If Not tFicheiroExiste((cBD_Path & cNomeDoFicheiro)) Then Exit Do
                cNomeDoFicheiro = "\SG" & CStr(Format$(Date, "yyyymmdd")) & "_" & CStr(iCiclo) & ".mdb"
            Loop
            Name (cBD_Path & cNomeBD) As (cBD_Path & cNomeDoFicheiro)
            'Copia a Base de Dados
            cCaminho = cCaminho & filPath.FileName
            FileCopy cCaminho, (cBD_Path & cNomeBD)
            MsgBox "Recuperação concluida com sucesso.", vbInformation + vbOKOnly, "Recuperação da Base de Dados"
        End If
    End If
End Sub

Private Sub dirPath_Change()
    txtPath.Text = UCase(dirPath.Path)
    filPath.Path = dirPath.Path
End Sub

Private Sub drvPath_Change()
    On Error GoTo DriveHandler
    dirPath.Path = drvPath.Drive
    Exit Sub

DriveHandler:
    drvPath.Drive = dirPath.Path
    Exit Sub
End Sub

Private Sub filPath_DblClick()
    Dim cCaminho
    cCaminho = UCase(dirPath.Path)
    If CStr(Right(cCaminho, 1)) <> "\" Then cCaminho = cCaminho & "\"
    txtPath.Text = cCaminho & UCase(filPath.FileName)
End Sub

Private Sub optImpExp_Click(Index As Integer)
    Select Case Index
        Case 0  'Exportações
            filPath.Enabled = False
            txtPath.Text = UCase(dirPath.Path)
        Case 1  'Importações
            filPath.Enabled = True
    End Select
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    If Dir(cBD_Path & "\Seguranças", vbDirectory) = vbNullString Then
        MkDir (cBD_Path & "\Seguranças")
    End If
    
    'Coloca-se na Directoria da Base de Dados
    dirPath.Path = cBD_Path & "\Seguranças"
    
    'Seleciona as extensões que vão ser mostradas
    filPath.Pattern = "SG_*.MDB"
    filPath.Enabled = False
    
    'Carrega o Destino
    txtPath.Text = UCase(dirPath.Path)
   
    Me.Show
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
    Set frmSegurancas = Nothing
End Sub

Private Sub optSegRec_Click(Index As Integer)
    If optSegRec(0).Value Then
        fraSeleccionar.Caption = " Selecione o Destino para a Segurança "
        filPath.Enabled = False
    Else
        fraSeleccionar.Caption = " Selecione a Origem da Segurança "
        filPath.Enabled = True
    End If
End Sub
