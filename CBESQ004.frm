VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFichaSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Sócio"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
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
   ScaleHeight     =   5940
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   46
      Tag             =   "2010"
      Top             =   4935
      Width           =   1200
   End
   Begin TabDlg.SSTab tabSocio 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 - Sócio"
      TabPicture(0)   =   "CBESQ004.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picFoto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSocio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2 - Dados Pessoais"
      TabPicture(1)   =   "CBESQ004.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDadosPessois"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Morada para Cobrança"
      TabPicture(2)   =   "CBESQ004.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCobranca"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Observações"
      TabPicture(3)   =   "CBESQ004.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraObs"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraObs 
         Caption         =   " Observações "
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
         Height          =   2025
         Left            =   -74850
         TabIndex        =   43
         Top             =   660
         Width           =   8805
         Begin VB.TextBox txtObs 
            Height          =   1500
            Left            =   150
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   300
            Width           =   7980
         End
      End
      Begin VB.Frame fraCobranca 
         Caption         =   "Morada para Cobrança "
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
         Height          =   1890
         Left            =   -74850
         TabIndex        =   36
         Top             =   600
         Width           =   8805
         Begin VB.TextBox txtLocal_Cob 
            Height          =   360
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1200
            Width           =   6300
         End
         Begin VB.TextBox txtCod_Postal_Cob 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   40
            Top             =   1200
            Width           =   990
         End
         Begin VB.TextBox txtMorada_Cob 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   38
            Top             =   540
            Width           =   7500
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   20
            Left            =   1350
            TabIndex        =   41
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   19
            Left            =   150
            TabIndex        =   39
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   18
            Left            =   150
            TabIndex        =   37
            Top             =   300
            Width           =   705
         End
      End
      Begin VB.Frame fraDadosPessois 
         Caption         =   " Dados Pessoais "
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
         Height          =   3780
         Left            =   -74850
         TabIndex        =   19
         Top             =   600
         Width           =   8805
         Begin VB.TextBox txtProfissao 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   21
            Top             =   540
            Width           =   6300
         End
         Begin VB.TextBox txtEmpresa 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1200
            Width           =   6300
         End
         Begin VB.TextBox txtTel_Empresa 
            Height          =   360
            Left            =   6720
            MaxLength       =   12
            TabIndex        =   25
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtLocal_Emi_BI 
            Height          =   360
            Left            =   150
            MaxLength       =   40
            TabIndex        =   35
            Top             =   3180
            Width           =   6300
         End
         Begin VB.TextBox txtNum_Contribuinte 
            Height          =   360
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   29
            Top             =   1860
            Width           =   1500
         End
         Begin VB.TextBox txtNum_BI 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   31
            Top             =   2520
            Width           =   1500
         End
         Begin GTMaskDate.GTMaskDate dcboData_Nasc 
            Height          =   375
            Left            =   150
            TabIndex        =   27
            Top             =   1860
            Width           =   1935
            _Version        =   65537
            _ExtentX        =   3413
            _ExtentY        =   661
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
            MaskCentury     =   2
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
         Begin GTMaskDate.GTMaskDate dcboData_BI 
            Height          =   375
            Left            =   2700
            TabIndex        =   33
            Top             =   2520
            Width           =   1935
            _Version        =   65537
            _ExtentX        =   3413
            _ExtentY        =   661
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
            MaskCentury     =   2
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Profissão"
            Height          =   240
            Index           =   10
            Left            =   150
            TabIndex        =   20
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   240
            Index           =   11
            Left            =   150
            TabIndex        =   22
            Top             =   960
            Width           =   825
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Tel. Empresa"
            Height          =   240
            Index           =   12
            Left            =   6720
            TabIndex        =   24
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nascimento"
            Height          =   240
            Index           =   13
            Left            =   150
            TabIndex        =   26
            Top             =   1620
            Width           =   1845
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Contribuinte"
            Height          =   240
            Index           =   14
            Left            =   2700
            TabIndex        =   28
            Top             =   1620
            Width           =   1605
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº Bilhete de Identidade"
            Height          =   240
            Index           =   15
            Left            =   150
            TabIndex        =   30
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão B. I."
            Height          =   240
            Index           =   16
            Left            =   2700
            TabIndex        =   32
            Top             =   2280
            Width           =   1920
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Local de Emissão do B. I."
            Height          =   240
            Index           =   17
            Left            =   150
            TabIndex        =   34
            Top             =   2940
            Width           =   2265
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
         Height          =   3780
         Left            =   150
         TabIndex        =   1
         Top             =   600
         Width           =   7200
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   3
            Top             =   540
            Width           =   6900
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1860
            Width           =   5655
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   7
            Top             =   1860
            Width           =   990
         End
         Begin VB.TextBox txtMorada 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   5
            Top             =   1200
            Width           =   6900
         End
         Begin VB.TextBox txtTel_Morada 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   11
            Top             =   2520
            Width           =   1500
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   17
            Top             =   3180
            Width           =   1500
         End
         Begin GTMaskNum.GTMaskNum txtQuota 
            Height          =   360
            Left            =   5175
            TabIndex        =   15
            Top             =   2520
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   635
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
            MaskErrorForeColor=   16777217
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
            MaskDecimalFixed=   -1  'True
            MaskType        =   0
            DataType        =   4
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
         Begin GTMaskDate.GTMaskDate dcboData_Admissao 
            Height          =   375
            Left            =   2475
            TabIndex        =   13
            Top             =   2520
            Width           =   1935
            _Version        =   65537
            _ExtentX        =   3413
            _ExtentY        =   661
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
            MaskCentury     =   2
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Admissão"
            Height          =   240
            Index           =   7
            Left            =   2490
            TabIndex        =   12
            Top             =   2280
            Width           =   1680
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Quota"
            Height          =   240
            Index           =   9
            Left            =   5190
            TabIndex        =   14
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   4
            Left            =   1305
            TabIndex        =   8
            Top             =   1620
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   6
            Top             =   1620
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   4
            Top             =   960
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   10
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemovel"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   16
            Top             =   2940
            Width           =   975
         End
      End
      Begin VB.PictureBox picFoto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   7455
         ScaleHeight     =   2040
         ScaleWidth      =   1560
         TabIndex        =   18
         Top             =   720
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   45
      Tag             =   "2002"
      Top             =   4935
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   47
      Tag             =   "2003"
      Top             =   4920
      Width           =   1200
   End
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   6195
      Top             =   5190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmFichaSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSSocios As Workspace
Dim mBDSocios As Database
Dim mBDSociosTemp As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lNum_Socio
Dim cNomeMapa

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    
    Set mProcessamento = New Processamento

    cNomeMapa = "CBESQ014.RPT"
On Error GoTo TrataErro
   
   ' apaga os registos da Temp32
    cSql = "DELETE * FROM FICHA_SOCIOS"
    ' apaga o registo em Temp32
    mBDSociosTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_SOCIOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM SOCIOS WHERE NUM_SOCIO=" & lNum_Socio
    ' insere o registo em Temp32
    mBDSocios.Execute cSql
                        
    With rptFichaIndividual
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Sócio"
            GoTo SairDoProcedimento
        End If
'        .WindowParentHandle = fFrmMDIPrincipal.hwnd
'        .WindowTitle = "Impressão de Ficha Inscrição"
'        .WindowState = crptMaximized
'        .WindowAllowDrillDown = False
'        .WindowBorderStyle = 2
'        .WindowControlBox = True
'        .WindowControls = True
'        .WindowMaxButton = True
'        .WindowMinButton = False
'        .WindowShowCloseBtn = False
'        .WindowShowExportBtn = False
'        .WindowShowGroupTree = False
'        .WindowShowNavigationCtls = True
'        .WindowShowPrintBtn = True
'        .WindowShowPrintSetupBtn = True
'        .WindowShowProgressCtls = False
'        .WindowShowZoomCtl = True
'        .WindowShowSearchBtn = False
'        .WindowShowRefreshBtn = False
        'Configura o destino e o numero de copias e de linhas para o Mapa
'        .Destination = crptToWindow
        .Destination = crptToPrinter
        .DataFiles(0) = cBDComNomeUtilizador
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                      
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Impressão Ficha de Socios", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qrySocio As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' Campos Obrigatórios
    If Trim$(txtNome.Text) = vbNullString Then
        MsgBox "Nome é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Sócio"
        txtNome.SetFocus
        Exit Sub
    End If
On Error GoTo TrataErro
    ' começa a transação
    mWSSocios.BeginTrans
    If cBotao = "Novo" Then
        Set qrySocio = mBDSocios.QueryDefs("SOCIOS Insere")
        ' parametros de input
        qrySocio.Parameters("Num_Socio") = lNovoNumSocio
        qrySocio.Parameters("Nome") = txtNome.Text
        qrySocio.Parameters("Morada") = txtMorada.Text
        qrySocio.Parameters("Local") = txtLocal.Text
        qrySocio.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qrySocio.Parameters("Tel_Morada") = txtTel_Morada.Text
        qrySocio.Parameters("Telemovel") = txtTelemovel.Text
        qrySocio.Parameters("Data_Admissao") = dcboData_Admissao.DateValue
        qrySocio.Parameters("Quota") = txtQuota.Value
        qrySocio.Parameters("Profissao") = txtProfissao.Text
        qrySocio.Parameters("Empresa") = txtEmpresa.Text
        qrySocio.Parameters("Tel_Empresa") = txtTel_Empresa.Text
        qrySocio.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        qrySocio.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qrySocio.Parameters("Num_BI") = txtNum_BI.Text
        qrySocio.Parameters("Data_BI") = dcboData_BI.DateValue
        qrySocio.Parameters("Local_BI") = txtLocal_Emi_BI.Text
        qrySocio.Parameters("Morada_Cob") = txtMorada_Cob.Text
        qrySocio.Parameters("Cod_Postal_Cob") = txtCod_Postal_Cob.Text
        qrySocio.Parameters("Local_Cob") = txtLocal_Cob.Text
        qrySocio.Parameters("Obs") = txtObs.Text
        qrySocio.Parameters("Utiliz") = gUtilizador.Nome
      ElseIf cBotao = "Altera" Then
        Set qrySocio = mBDSocios.QueryDefs("SOCIOS Altera")
        ' parametros de input
        qrySocio.Parameters("Num_Socio") = lNum_Socio
        qrySocio.Parameters("Nome") = txtNome.Text
        qrySocio.Parameters("Morada") = txtMorada.Text
        qrySocio.Parameters("Local") = txtLocal.Text
        qrySocio.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qrySocio.Parameters("Tel_Morada") = txtTel_Morada.Text
        qrySocio.Parameters("Telemovel") = txtTelemovel.Text
        qrySocio.Parameters("Data_Admissao") = dcboData_Admissao.DateValue
        qrySocio.Parameters("Quota") = txtQuota.Value
        qrySocio.Parameters("Profissao") = txtProfissao.Text
        qrySocio.Parameters("Empresa") = txtEmpresa.Text
        qrySocio.Parameters("Tel_Empresa") = txtTel_Empresa.Text
        qrySocio.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        qrySocio.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qrySocio.Parameters("Num_BI") = txtNum_BI.Text
        qrySocio.Parameters("Data_BI") = dcboData_BI.DateValue
        qrySocio.Parameters("Local_BI") = txtLocal_Emi_BI.Text
        qrySocio.Parameters("Morada_Cob") = txtMorada_Cob.Text
        qrySocio.Parameters("Cod_Postal_Cob") = txtCod_Postal_Cob.Text
        qrySocio.Parameters("Local_Cob") = txtLocal_Cob.Text
        qrySocio.Parameters("Obs") = txtObs.Text
        qrySocio.Parameters("Utiliz") = gUtilizador.Nome
    End If
    ' executa a query
    qrySocio.Execute dbFailOnError
    
    mWSSocios.CommitTrans
    ' faz o refresh da frmGestaoSocios
    frmGestaoSocios.datSocios.Refresh
    frmGestaoSocios.sgrdGestaoSocios.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSSocios.Rollback
    Call ErrosGerais(cBotao & " Sócio", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Unload Me
End Sub

Private Sub dcboData_Admissao_DropDown()
    If Not IsDate(dcboData_Admissao.Text) Then
        dcboData_Admissao.DateValue = Date
    End If
End Sub


Private Sub dcboData_BI_DropDown()
    If Not IsDate(dcboData_BI.Text) Then
        dcboData_BI.DateValue = Date
    End If
End Sub


Private Sub dcboData_Nasc_DropDown()
    If Not IsDate(dcboData_Nasc.Text) Then
        dcboData_Nasc.DateValue = Date
    End If
End Sub


Private Sub Form_Activate()
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools(Me.Caption).State = ssChecked
End Sub

'Este procedimento activa ou desactiva o default do botao OK
Private Sub BotaoOKDefault(ByVal Propriedade As Boolean)
  cmdOK.Default = Propriedade
End Sub

Private Sub Form_Load()
    cBotao = cBotaoOrigem
    If cBotao <> "Novo" Then
        lNum_Socio = CLng(frmGestaoSocios.sgrdGestaoSocios.Columns(0).Text)
    End If
    
    If cBotao = "Ficha" Then
        Me.Caption = Me.Caption & " Nº " & lNum_Socio
        cmdOK.Visible = False
        cmdImprimir.Visible = True
    ElseIf cBotao = "Novo" Then
        Me.Caption = "Nova " & Me.Caption
        cmdImprimir.Visible = False
   ElseIf cBotao = "Altera" Then
        Me.Caption = "Alteração na " & Me.Caption & " Nº " & lNum_Socio
        cmdImprimir.Visible = False
    End If

    CenterMe Me
    LoadResStrings Me
        
    Call AlteraWindowList(Me.Caption)
    
    tBDAberta = tAbreBD
    
    Call CamposEnabledDisabled
    Call CamposLimpaCarrega
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSSocios.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSSocios = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSSocios = DBEngine.CreateWorkspace("Socios", gUtilizador.Nome, gUtilizador.Password)
    Set mBDSocios = mWSSocios.OpenDatabase(cBD_Path & cNomeBD)
    If cBotao = "Ficha" Then
        Set mBDSociosTemp = mWSSocios.OpenDatabase(cBDComNomeUtilizador)
    End If
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Ficha Sócios-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function



Private Sub CamposEnabledDisabled()
    If cBotaoOrigem = "Ficha" Then
        'Poe os campos locked só para consulta da ficha
        txtNome.Locked = True
        txtNome.TabStop = False
        txtNome.BackColor = &H8000000F
        
        txtMorada.Locked = True
        txtMorada.TabStop = False
        txtMorada.BackColor = &H8000000F
        
        txtCodigoPostal.Locked = True
        txtCodigoPostal.TabStop = False
        txtCodigoPostal.BackColor = &H8000000F
        
        txtLocal.Locked = True
        txtLocal.TabStop = False
        txtLocal.BackColor = &H8000000F
        
        txtTel_Morada.Locked = True
        txtTel_Morada.TabStop = False
        txtTel_Morada.BackColor = &H8000000F
        
        txtTelemovel.Locked = True
        txtTelemovel.TabStop = False
        txtTelemovel.BackColor = &H8000000F
        
        dcboData_Admissao.Locked = True
        dcboData_Admissao.CalDropDown = False
        dcboData_Admissao.TabStop = False
        
        txtQuota.Locked = True
        txtQuota.TabStop = False
        
        txtProfissao.Locked = True
        txtProfissao.TabStop = False
        txtProfissao.BackColor = &H8000000F
        
        txtEmpresa.Locked = True
        txtEmpresa.TabStop = False
        txtEmpresa.BackColor = &H8000000F
        
        txtTel_Empresa.Locked = True
        txtTel_Empresa.TabStop = False
        txtTel_Empresa.BackColor = &H8000000F
        
        dcboData_Nasc.Locked = True
        dcboData_Nasc.CalDropDown = False
        dcboData_Nasc.TabStop = False
        
        txtNum_Contribuinte.Locked = True
        txtNum_Contribuinte.TabStop = False
        txtNum_Contribuinte.BackColor = &H8000000F
        
        txtNum_BI.Locked = True
        txtNum_BI.TabStop = False
        txtNum_BI.BackColor = &H8000000F
        
        dcboData_BI.Locked = True
        dcboData_BI.CalDropDown = False
        dcboData_BI.TabStop = False
        
        txtLocal_Emi_BI.Locked = True
        txtLocal_Emi_BI.TabStop = False
        txtLocal_Emi_BI.BackColor = &H8000000F
        
        txtMorada_Cob.Locked = True
        txtMorada_Cob.TabStop = False
        txtMorada_Cob.BackColor = &H8000000F
        
        txtCod_Postal_Cob.Locked = True
        txtCod_Postal_Cob.TabStop = False
        txtCod_Postal_Cob.BackColor = &H8000000F
        
        txtLocal_Cob.Locked = True
        txtLocal_Cob.TabStop = False
        txtLocal_Cob.BackColor = &H8000000F
        
        txtObs.Locked = True
        txtObs.TabStop = False
        txtObs.BackColor = &H8000000F
    End If
End Sub

Public Sub CamposLimpaCarrega()
    Dim recSocio As Recordset
    
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM SOCIOS WHERE NUM_SOCIO=" & CLng(frmGestaoSocios.sgrdGestaoSocios.Columns(0).Value)
        Set recSocio = mBDSocios.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        'Poe os campos com os dados do sócio
        If Dir(cApl_Path & "\FOTOS\" & recSocio!NUM_SOCIO & ".BMP") = "" Then
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        Else
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\" & recSocio!NUM_SOCIO & ".BMP")
        End If
        txtNome.Text = vFiltraCamposNulos(recSocio!Nome)
        txtMorada.Text = vFiltraCamposNulos(recSocio!MORADA)
        txtCodigoPostal.Text = vFiltraCamposNulos(recSocio!COD_POSTAL)
        txtLocal.Text = vFiltraCamposNulos(recSocio!LOCAL)
        txtTel_Morada.Text = vFiltraCamposNulos(recSocio!TELEFONE)
        txtTelemovel.Text = vFiltraCamposNulos(recSocio!TELEMOVEL)
        dcboData_Admissao.Text = vFiltraCamposNulos(recSocio!DATA_ADMISSAO)
        txtQuota.Text = vFiltraCamposNulos(recSocio!QUOTA)
        txtProfissao.Text = vFiltraCamposNulos(recSocio!PROFISSAO)
        txtEmpresa.Text = vFiltraCamposNulos(recSocio!Empresa)
        txtTel_Empresa.Text = vFiltraCamposNulos(recSocio!TEL_EMPRESA)
        dcboData_Nasc.Text = vFiltraCamposNulos(recSocio!DATA_NASC)
        txtNum_Contribuinte.Text = vFiltraCamposNulos(recSocio!NUM_CONTRIBUINTE)
        txtNum_BI.Text = vFiltraCamposNulos(recSocio!NUM_BI)
        dcboData_BI.Text = vFiltraCamposNulos(recSocio!DATA_BI)
        txtLocal_Emi_BI.Text = vFiltraCamposNulos(recSocio!LOCAL_EMI_BI)
        txtMorada_Cob.Text = vFiltraCamposNulos(recSocio!MORADA_COB)
        txtCod_Postal_Cob.Text = vFiltraCamposNulos(recSocio!COD_POSTAL_COB)
        txtLocal_Cob.Text = vFiltraCamposNulos(recSocio!LOCAL_COB)
        txtObs.Text = vFiltraCamposNulos(recSocio!OBS)
        ' fecha o recordset
        recSocio.Close
        Set recSocio = Nothing
    ElseIf cBotao = "Novo" Then
        'Poe os campos preparados para nova ficha
        picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        txtNome.Text = vbNullString
        txtMorada.Text = vbNullString
        txtCodigoPostal.Text = vbNullString
        txtLocal.Text = vbNullString
        txtTel_Morada.Text = vbNullString
        txtTelemovel.Text = vbNullString
        dcboData_Admissao.DateValue = Date
        txtQuota.Text = 0
        txtProfissao.Text = vbNullString
        txtEmpresa.Text = vbNullString
        txtTel_Empresa.Text = vbNullString
        dcboData_Nasc.Text = vbNullString
        txtNum_Contribuinte.Text = vbNullString
        txtNum_BI.Text = vbNullString
        dcboData_BI.Text = vbNullString
        txtLocal_Emi_BI.Text = vbNullString
        txtMorada_Cob.Text = vbNullString
        txtCod_Postal_Cob.Text = vbNullString
        txtLocal_Cob.Text = vbNullString
        txtObs.Text = vbNullString
    End If
End Sub


