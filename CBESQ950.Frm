VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmUtilizadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilizadores"
   ClientHeight    =   3345
   ClientLeft      =   1590
   ClientTop       =   1590
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   3900
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "2003"
      Top             =   2385
      Width           =   1200
   End
   Begin VB.CommandButton cmdApagar 
      Caption         =   "&Apagar"
      Height          =   900
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "2004"
      Top             =   2385
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Confirmar"
      Height          =   900
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "2002"
      Top             =   2385
      Width           =   1200
   End
   Begin VB.TextBox txtNivel 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1995
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1860
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1995
      MaxLength       =   14
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   540
      Width           =   1770
   End
   Begin VB.TextBox txtNome 
      Height          =   360
      Left            =   150
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1200
      Width           =   3660
   End
   Begin SSDataWidgets_B.SSDBCombo scboPerfil 
      Height          =   360
      Left            =   105
      TabIndex        =   7
      Top             =   1860
      Width           =   1800
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
      Columns(0).Width=   3200
      _ExtentX        =   3175
      _ExtentY        =   635
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B.SSDBCombo scboUtilizador 
      Height          =   360
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   1830
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
      Columns(0).Width=   3200
      _ExtentX        =   3228
      _ExtentY        =   635
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.Label lblUtilizadores 
      AutoSize        =   -1  'True
      Caption         =   "Nivel"
      Height          =   240
      Index           =   5
      Left            =   1995
      TabIndex        =   8
      Top             =   1620
      Width           =   465
   End
   Begin VB.Label lblUtilizadores 
      AutoSize        =   -1  'True
      Caption         =   "Perfil"
      Height          =   240
      Index           =   3
      Left            =   105
      TabIndex        =   6
      Top             =   1620
      Width           =   450
   End
   Begin VB.Label lblUtilizadores 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   240
      Index           =   2
      Left            =   150
      TabIndex        =   4
      Top             =   960
      Width           =   555
   End
   Begin VB.Label lblUtilizadores 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   240
      Index           =   1
      Left            =   1995
      TabIndex        =   2
      Top             =   300
      Width           =   900
   End
   Begin VB.Label lblUtilizadores 
      AutoSize        =   -1  'True
      Caption         =   "Utilizador"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "frmUtilizadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWorkspace As Workspace
Dim mBDBaseDados As Database
Dim mBDTemp32 As Database
Dim tUtilizadorNovo
Dim tAlteraDados

Function EnumerateDocuments() As Integer
    Dim MyDatabase As Database
    Dim MyContainer As Container, MyDocument As Document
    Dim i As Integer, J As Integer
    Set MyDatabase = gwsInicial.OpenDatabase(cBD_Path & "\Gestao32.mdb")
    For J = 0 To MyDatabase.Containers.Count - 1
        Set MyContainer = MyDatabase.Containers(J)
        Debug.Print ">> Container: "; MyContainer.Name;
        Debug.Print "  Owner: "; MyContainer.Owner
        Debug.Print "  UserName: "; MyContainer.UserName;
        Debug.Print "  Permissions: "; MyContainer.Permissions
        For i = 0 To MyContainer.Documents.Count - 1
            Set MyDocument = MyContainer.Documents(i)
            Debug.Print " > Document: "; MyDocument.Name;
            Debug.Print "  Owner: "; MyDocument.Owner;
            Debug.Print "  Container: "; MyDocument.Container
            Debug.Print "  UserName: "; MyDocument.UserName;
            Debug.Print "  Permissions: "; MyDocument.Permissions
        Next i
    Next J
    EnumerateDocuments = True
End Function
'Este procedimento carrega para a combo perfil os perfis existentes na bd
Private Sub CarregaComboPerfil()
    Dim mrecTabPerfis As Recordset
    
    On Error Resume Next
    
    Set mrecTabPerfis = mBDBaseDados.OpenRecordset("SELECT PERFIL FROM TABPERFIS WHERE NIVEL>=" & gUtilizador.Nivel & " ORDER BY PERFIL ASC", dbOpenSnapshot, dbForwardOnly + dbReadOnly)
    
    scboPerfil.Redraw = False
    scboPerfil.RemoveAll
    With mrecTabPerfis
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                scboPerfil.AddItem Trim$(vFiltraCamposNulos(.Fields("PERFIL")))
                .MoveNext
            Loop
        End If
    End With
    scboPerfil.Redraw = True
End Sub


'Este procedimento insere um novo utilizador no Grupo Clientes no System.mda
'e atribui permissão as BD Orca32.mdb e Temp32.mdb
Private Sub NovoUtilizador(ByVal cUtilizador As String, ByVal cPassword As String)
    Dim NovoUtilizador As User, cPid, lPermissao
    Dim mContainer As Container, mDocument As Document
    Dim i, cSql
    
    cPid = "CBESQ"
    
    On Error GoTo TrataErro
   
    'Cria um novo utilizador
    Set NovoUtilizador = mWorkspace.CreateUser(cUtilizador, cPid, cPassword)
    'Grava o utilizador acrescentando ao outros utilizadores
    mWorkspace.Users.Append NovoUtilizador
    mWorkspace.Groups.Refresh
    
    'Acrescenta o utilizador ao grupo
    NovoUtilizador.Groups.Append NovoUtilizador.CreateGroup("Clientes")
    NovoUtilizador.Groups.Append NovoUtilizador.CreateGroup("Admins")
               
    'Constant Description
    'dbSecNoAccess          No access to the object.
    'dbSecFullAccess        Full access to the object.
    'dbSecDelete            Can delete the object.
    'dbSecReadSec           Can read the object's security-related information.
    'dbSecWriteSec          Can alter access permissions.
    'dbSecWriteOwner        Can change the Owner property setting.
    
    'For the Tables Container or any Document in a Documents collection, the settings for value are:
    'Constant Description
    'dbSecCreate            Can create new documents (valid only with a Container object).
    'dbSecReadDef           Can read the table definition including column and index information.
    'dbSecWriteDef          Can modify or delete the table definition including column and index information.
    'dbSecRetrieveData      Can retrieve data from the Document.
    'dbSecInsertData        Can add records.
    'dbSecReplaceData       Can modify records.
    'dbSecDeleteData        Can delete records.
    
    'For the Databases Container object or any Document object in a Documents collection, the settings for value are:
    'Constant User Or Group
    'dbSecDBCreate          Can create new databases (valid only on the databases Container object in the system database [SYSTEM.MDW]) .
    'dbSecDBExclusive       Exclusive access.
    'dbSecDBOpen            Can open the database.
    
    On Error Resume Next
    
    'Trata a BD CBESQ32
    lPermissao = dbSecReadSec
    
    Set mContainer = mBDBaseDados.Containers("Tables")
    mContainer.UserName = cUtilizador
    mContainer.Permissions = mContainer.Permissions Or lPermissao
    
    lPermissao = dbSecRetrieveData Or dbSecInsertData Or dbSecReplaceData Or _
                dbSecDeleteData
    
    For i = 0 To mContainer.Documents.Count - 1
        Set mDocument = mContainer.Documents(i)
        mDocument.UserName = cUtilizador
        mDocument.Permissions = mDocument.Permissions Or lPermissao
    Next i
    
    'Trata a BD Temp32
    lPermissao = dbSecNoAccess
    
    Set mContainer = mBDTemp32.Containers("Tables")
    mContainer.UserName = cUtilizador
    mContainer.Permissions = mContainer.Permissions Or lPermissao
    
    lPermissao = dbSecReadDef Or dbSecRetrieveData Or dbSecInsertData Or dbSecReplaceData _
                Or dbSecDeleteData
    
    For i = 0 To mContainer.Documents.Count - 1
        Set mDocument = mContainer.Documents(i)
        mDocument.UserName = cUtilizador
        mDocument.Permissions = mDocument.Permissions Or lPermissao
    Next i
    
    On Error GoTo ErroNaBD
    
    mWorkspace.BeginTrans
    cSql = "INSERT INTO TABUTILIZADORES(UTILIZADOR,NOME,PERFIL,NIVEL,DATACRIACAO) VALUES('" & cUtilizador & _
            "','" & Trim$(txtNome.Text) & "','" & Trim$(scboPerfil.Text) & "'," & Val(txtNivel.Text) & ",#" & Date & "#)"
    
    mBDBaseDados.Execute cSql, dbFailOnError
    mWorkspace.CommitTrans
    
    GoTo SaiDoProcedimento
    
ErroNaBD:
    mWorkspace.Rollback
TrataErro:
    Call ErrosGerais("Utilizadores-Novo Utilizador", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
    
End Sub

Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

Private Sub Form_Load()
    CenterMe Me
    LoadResStrings Me
    
    Call AlteraWindowList(Me.Caption)
    
    Call SetNumber(txtNivel, True)
    
    tUtilizadorNovo = True
    tAlteraDados = False
    
    Me.Show
    DoEvents
    
    If Not tAbreBD Then
        Unload Me
    End If
    'EnumerateDocuments
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWorkspace.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set frmUtilizadores = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub






Private Sub scboPerfil_DropDown()
    Call CarregaComboPerfil
End Sub

Private Sub scboPerfil_GotFocus()
     scboPerfil.BackColor = dCorSelecionado
End Sub

Private Sub scboPerfil_InitColumnProps()
    With scboPerfil
        .AllowInput = True
        .AllowNull = True
        .AutoRestore = True
        .BackColorOdd = dCorAmarelo
        .DataFieldList = "Column 0"
        .DataMode = 2
        .FieldSeparator = vbTab
        .Font.Name = "MS Sans Serif"
        .Font.Size = 9
        .ListAutoPosition = True
        .ListAutoValidate = True
        .MaxDropDownItems = 12
        .MinDropDownItems = 1
        .Columns.Add (0)
        .Columns(0).Caption = "Perfil"
        .Columns(0).Width = .Width
    End With
End Sub

Private Sub scboPerfil_LostFocus()
    scboPerfil.BackColor = dCorNormal
End Sub

Private Sub scboUtilizador_Change()
    tAlteraDados = False
    tUtilizadorNovo = False
End Sub

Private Sub scboUtilizador_Click()
    tAlteraDados = False
    tUtilizadorNovo = False
    If UCase(scboUtilizador.Text) <> UCase(gUtilizador.Nome) And gUtilizador.Nivel > 1 Then Call EstadoDosCampos(False)
    Call cmdOK_Click
End Sub

Private Sub scboUtilizador_DropDown()
    CarregaComboUtilizadores
End Sub

Private Sub scboUtilizador_GotFocus()
    scboUtilizador.BackColor = dCorSelecionado
    cmdOK.Default = True
    Call EstadoDosCampos(True)
End Sub


Private Sub scboUtilizador_InitColumnProps()
    With scboUtilizador
        .AllowInput = True
        .AllowNull = True
        .AutoRestore = True
        .BackColorOdd = dCorAmarelo
        .DataFieldList = "Column 0"
        .DataMode = 2
        .FieldSeparator = vbTab
        .Font.Name = "MS Sans Serif"
        .Font.Size = 9
        .ListAutoPosition = True
        .ListAutoValidate = True
        .MaxDropDownItems = 12
        .MinDropDownItems = 1
        
        .Columns.Add (0)
        .Columns(0).Caption = "Utilizador"
        .Columns(0).Width = .Width
    End With
End Sub

Private Sub scboUtilizador_LostFocus()
    With scboUtilizador
        .BackColor = dCorNormal
        'Se for um utilizador já registado
        If .IsItemInList Then
            tUtilizadorNovo = False
            If UCase(scboUtilizador.Text) <> UCase(gUtilizador.Nome) And gUtilizador.Nivel > 1 Then Call EstadoDosCampos(False)
        Else
            tUtilizadorNovo = True
        End If
    End With
    cmdOK.Default = False
End Sub

Private Sub cmdApagar_Click()
    Call ApagaUtilizador(Trim$(scboUtilizador.Text))
End Sub

Private Sub cmdOK_Click()
    Dim cUtilizador, cPasswordAcesso
    
    On Error Resume Next
    
    cUtilizador = Trim$(scboUtilizador.Text)
    
    'If UCase(cUtilizador) <> UCase(gUtilizador.Nome) And gUtilizador.Nivel <= 1 Then Call EstadoDosCampos(False)
    
    'Verifica se os campos estão todos preenchidos
    If Len(cUtilizador) > 0 Then
        If tUtilizadorNovo Then
            cPasswordAcesso = Trim$(txtPassword.Text)
            Call NovoUtilizador(cUtilizador, cPasswordAcesso)
        Else
            If tAlteraDados Then
                Call AlteraDadosUtilizador(cUtilizador)
            Else
                'Edita o utilizador
                Call EditaUtilizador(cUtilizador)
            End If
        End If
    Else
        scboUtilizador.SetFocus
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

'Este procedimento vai carregar para a combo todos os utilizadores
'que pertencem ao grupo Clientes
Private Sub CarregaComboUtilizadores()
    Dim mUtilizador As User, mGrupo As Group, i, iJ
    
    On Error Resume Next
    
    scboUtilizador.Redraw = False
    scboUtilizador.RemoveAll
    ' Enumerate all users.
    For iJ = 0 To mWorkspace.Users.Count - 1
        Set mUtilizador = mWorkspace.Users(iJ)
        ' Enumerate groups.
        For i = 0 To mUtilizador.Groups.Count - 1
            If Trim$(mUtilizador.Groups(i).Name) = "Clientes" Then
                scboUtilizador.AddItem Trim$(mUtilizador.Name)
            End If
        Next i
    Next iJ
    'For iJ = 0 To mWorkspace.Groups("Clientes").Users.Count
    '    Set mUtilizador = mWorkspace.Groups("Clientes").Users(iJ)
    '    scboUtilizador.AddItem Trim$(mUtilizador.Name)
    'Next iJ
    scboUtilizador.Redraw = True
End Sub

Private Sub txtNivel_GotFocus()
     txtNivel.BackColor = dCorSelecionado
End Sub

Private Sub txtNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txtNivel_LostFocus()
    txtNivel.BackColor = dCorNormal
    If CByte(0 & txtNivel.Text) < gUtilizador.Nivel Then txtNivel.Text = gUtilizador.Nivel
End Sub


Private Sub txtNome_GotFocus()
     txtNome.BackColor = dCorSelecionado
End Sub


Private Sub txtNome_LostFocus()
    txtNome.BackColor = dCorNormal
End Sub


Private Sub txtPassword_GotFocus()
     txtPassword.BackColor = dCorSelecionado
End Sub


Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = dCorNormal
End Sub



'Este procedimento edita para o ecran o utilizador
Private Sub EditaUtilizador(ByVal cUtilizador As String)
    Dim mrecTabUtilizador As Recordset
    
    On Error GoTo TrataErro
    
    Set mrecTabUtilizador = mBDBaseDados.OpenRecordset("Select * from TABUTILIZADORES", dbOpenDynaset, dbSeeChanges + dbReadOnly)
    With mrecTabUtilizador
        'Se houverem registos
        If Not (.EOF And .BOF) Then
            .FindFirst "UTILIZADOR='" & cUtilizador & "'"
            'se encontrou o utilizador
            If Not .NoMatch Then
                txtNome.Text = vFiltraCamposNulos(.Fields("NOME"))
                scboPerfil.Text = vFiltraCamposNulos(.Fields("PERFIL"))
                If UCase(cUtilizador) = UCase(gUtilizador.Nome) Then
                    txtPassword.Text = gUtilizador.Password
                Else
                    txtPassword.Text = ""
                End If
                txtNivel.Text = vFiltraCamposNulos(.Fields("NIVEL"))
            End If
        End If
        .Close
    End With
    GoTo SaiDoProcedimento
    
TrataErro:
    Call ErrosGerais("Utilizadores-Edita Utilizador", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
    Set mrecTabUtilizador = Nothing
    tAlteraDados = True
    
End Sub

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWorkspace = DBEngine.CreateWorkspace("Utilizadores", gUtilizador.Nome, gUtilizador.Password)
    Set mBDBaseDados = mWorkspace.OpenDatabase(cBD_Path & cNomeBD)
    Set mBDTemp32 = mWorkspace.OpenDatabase(cBD_Path & cNomeBDTemp)
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Utilizadores-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
    
End Function



'Este procedimento vai apagar um utilizador
Private Sub ApagaUtilizador(cUtilizador As String)
    Dim cSql
    
    On Error GoTo TrataErro
    
    gwsInicial.BeginTrans
    
    cSql = "DELETE * FROM TABUTILIZADORES WHERE UTILIZADOR='" & cUtilizador & "'"
    mBDBaseDados.Execute cSql, dbFailOnError
    'Apaga o utilizador
    mWorkspace.Users.Delete cUtilizador
    mWorkspace.Users.Refresh
    
    gwsInicial.CommitTrans
    GoTo SaiDoProcedimento
    
TrataErro:
    mWorkspace.Rollback
    Call ErrosGerais("Utilizadores-Apaga Utilizador", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
    
End Sub

Private Sub AlteraDadosUtilizador(ByVal cUtilizador)
    Dim Utilizador As User, cPassWordAntiga, cPassWordNova, cSql
       
    On Error GoTo TrataErro
   
    'Define o utilizador
    Set Utilizador = mWorkspace.Users(cUtilizador)
    'Só o Proprio utilizador pode alterar a password e todos os restantes dados á excepção do perfil e do nivel
    If UCase(cUtilizador) = UCase(gUtilizador.Nome) Then
        cPassWordAntiga = gUtilizador.Password
        cPassWordNova = Trim$(txtPassword.Text)
        Utilizador.NewPassword cPassWordAntiga, cPassWordNova
        
        On Error GoTo ErroNaBD
    
        mWorkspace.BeginTrans
        cSql = "UPDATE TABUTILIZADORES SET NOME='" & Trim$(txtNome.Text) & "' WHERE UTILIZADOR='" & cUtilizador & "'"
        
        mBDBaseDados.Execute cSql, dbFailOnError
        mWorkspace.CommitTrans
    ElseIf gUtilizador.Nivel <= 1 Then  'Só os utilizadores com nivel 1 ou 0 podem alterar os dados de um utilizador
        On Error GoTo ErroNaBD
    
        mWorkspace.BeginTrans
        cSql = "UPDATE TABUTILIZADORES SET NOME='" & Trim$(txtNome.Text) & "',PERFIL='" & Trim$(scboPerfil.Text) & "',NIVEL=" & Val(txtNivel.Text) & " WHERE UTILIZADOR='" & cUtilizador & "'"
        
        mBDBaseDados.Execute cSql, dbFailOnError
        mWorkspace.CommitTrans
    Else
        MsgBox "Não tem permissão para alterar os dados, contacte o Administrador do Sistema.", vbInformation + vbOKOnly
    End If
    
    
    
    GoTo SaiDoProcedimento
    
ErroNaBD:
    mWorkspace.Rollback
TrataErro:
    Call ErrosGerais("Utilizadores-Novo Utilizador", Err.Number, Err.Description)
    Resume SaiDoProcedimento

SaiDoProcedimento:
    
End Sub

Private Sub EstadoDosCampos(ByRef tPropriedade)
    scboPerfil.Enabled = tPropriedade
    txtNivel.Enabled = tPropriedade
    txtPassword.Enabled = tPropriedade
End Sub


