VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFichaFuncionario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Funcionário"
   ClientHeight    =   6825
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
   ScaleHeight     =   6825
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   45
      Tag             =   "2010"
      Top             =   5865
      Width           =   1200
   End
   Begin TabDlg.SSTab tabFuncionario 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "&1 - Funcionário"
      TabPicture(0)   =   "CBESQ019.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picFoto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFuncionario"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2 - Dados Pessoais"
      TabPicture(1)   =   "CBESQ019.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCategoria"
      Tab(1).Control(1)=   "fraNumeros"
      Tab(1).Control(2)=   "fraBI"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraCategoria 
         Caption         =   " Categoria "
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
         Height          =   1200
         Left            =   -74850
         TabIndex        =   40
         Top             =   2475
         Width           =   6600
         Begin VB.TextBox txtCategoria 
            Height          =   360
            Left            =   150
            MaxLength       =   40
            TabIndex        =   43
            Top             =   600
            Width           =   6300
         End
         Begin SSDataWidgets_B.SSDBCombo cboCategoria 
            Height          =   360
            Left            =   150
            TabIndex        =   42
            Top             =   600
            Width           =   6300
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   11112
            _ExtentY        =   635
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   41
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame fraNumeros 
         Caption         =   " Nº de "
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
         Height          =   1800
         Left            =   -68190
         TabIndex        =   35
         Top             =   600
         Width           =   2250
         Begin VB.TextBox txtSeguranca_Social 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   39
            Top             =   1200
            Width           =   1950
         End
         Begin VB.TextBox txtNum_Contribuinte 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   37
            Top             =   540
            Width           =   1950
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Seg. Social"
            Height          =   240
            Index           =   10
            Left            =   150
            TabIndex        =   38
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   36
            Top             =   300
            Width           =   1050
         End
      End
      Begin VB.Frame fraDatas 
         Caption         =   " Datas "
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
         Height          =   1740
         Left            =   6870
         TabIndex        =   23
         Top             =   3720
         Width           =   2235
         Begin GTMaskDate.GTMaskDate dcboData_Admissao 
            Height          =   375
            Left            =   150
            TabIndex        =   25
            Top             =   540
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
         Begin GTMaskDate.GTMaskDate dcboData_Demissao 
            Height          =   375
            Left            =   135
            TabIndex        =   27
            Top             =   1200
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
            Caption         =   "Demissão"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   26
            Top             =   960
            Width           =   930
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Admissão"
            Height          =   240
            Index           =   7
            Left            =   150
            TabIndex        =   24
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.Frame fraBI 
         Caption         =   " Bilhete de Identidade "
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
         Height          =   1800
         Left            =   -74850
         TabIndex        =   28
         Top             =   600
         Width           =   6600
         Begin VB.TextBox txtLocal_Emi_BI 
            Height          =   360
            Left            =   120
            MaxLength       =   40
            TabIndex        =   34
            Top             =   1200
            Width           =   6300
         End
         Begin VB.TextBox txtNum_BI 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   30
            Top             =   540
            Width           =   1500
         End
         Begin GTMaskDate.GTMaskDate dcboData_BI 
            Height          =   375
            Left            =   2100
            TabIndex        =   32
            Top             =   540
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
            Caption         =   "Nº"
            Height          =   240
            Index           =   15
            Left            =   150
            TabIndex        =   29
            Top             =   300
            Width           =   225
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão"
            Height          =   240
            Index           =   16
            Left            =   2100
            TabIndex        =   31
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Local de Emissão"
            Height          =   240
            Index           =   17
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1620
         End
      End
      Begin VB.Frame fraFuncionario 
         Caption         =   " Funcionário "
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
         Height          =   4860
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   6600
         Begin GTMaskNum.GTMaskNum txtNumFunc 
            Height          =   360
            Left            =   150
            TabIndex        =   48
            Top             =   540
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
            ValidateRangeHigh=   "5000"
            ValidateRangeLow=   "1"
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
            MaskType        =   0
            DataType        =   1
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
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   14
            Top             =   3060
            Width           =   1500
         End
         Begin VB.TextBox txtEstado_Civil 
            Height          =   360
            Left            =   150
            TabIndex        =   19
            Top             =   3690
            Width           =   4200
         End
         Begin SSDataWidgets_B.SSDBCombo cboEstado_Civil 
            Height          =   360
            Left            =   150
            TabIndex        =   18
            Top             =   3690
            Width           =   4200
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   7408
            _ExtentY        =   635
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.TextBox txtInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   22
            Top             =   4320
            Width           =   5250
         End
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1170
            Width           =   6300
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1305
            MaxLength       =   45
            TabIndex        =   10
            Top             =   2430
            Width           =   5160
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   8
            Top             =   2430
            Width           =   990
         End
         Begin VB.TextBox txtMorada 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   6
            Top             =   1800
            Width           =   6300
         End
         Begin VB.TextBox txtTel_Morada 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   12
            Top             =   3060
            Width           =   1500
         End
         Begin GTMaskDate.GTMaskDate dcboData_Nasc 
            Height          =   375
            Left            =   3435
            TabIndex        =   16
            Top             =   3060
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
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   21
            Top             =   4320
            Width           =   5250
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   9260
            _ExtentY        =   635
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   240
            Index           =   14
            Left            =   150
            TabIndex        =   47
            Top             =   300
            Width           =   225
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemovel"
            Height          =   240
            Index           =   8
            Left            =   1800
            TabIndex        =   13
            Top             =   2820
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Instituição onde desempenha as funções"
            Height          =   240
            Index           =   12
            Left            =   150
            TabIndex        =   20
            Top             =   4080
            Width           =   3660
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil"
            Height          =   240
            Index           =   11
            Left            =   150
            TabIndex        =   17
            Top             =   3450
            Width           =   1065
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nascimento"
            Height          =   240
            Index           =   13
            Left            =   3435
            TabIndex        =   15
            Top             =   2820
            Width           =   1845
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   4
            Left            =   1305
            TabIndex        =   9
            Top             =   2190
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   7
            Top             =   2190
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   1560
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   11
            Top             =   2820
            Width           =   810
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
         Left            =   6900
         ScaleHeight     =   2040
         ScaleWidth      =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   44
      Tag             =   "2002"
      Top             =   5865
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   46
      Tag             =   "2003"
      Top             =   5850
      Width           =   1200
   End
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   6180
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmFichaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSFunc As Workspace
Dim mBDFunc As Database
Dim mBDFuncTemp As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lNum_Funcionario
Dim cNomeMapa


Private Sub cboCategoria_DropDown()
    ' carrega a combo
    Call CarregacboCategoria(cboCategoria)
End Sub

Private Sub cboCategoria_InitColumnProps()
    With cboCategoria
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
        .Columns(0).Caption = "Categoria"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
    End With
End Sub

Private Sub cboEstado_Civil_DropDown()
    ' carrega a combo
    Call CarregacboEstadoCivil(cboEstado_Civil)
End Sub

Private Sub cboEstado_Civil_InitColumnProps()
    With cboEstado_Civil
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
        .Columns(0).Caption = "Estado Civil"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
    End With
End Sub


Private Sub cboInstituicao_DropDown()
    ' carrega a combo
    Call CarregacboInstituicao(cboInstituicao)
End Sub

Private Sub cboInstituicao_InitColumnProps()
    With cboInstituicao
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
        .Columns(0).Caption = "Nome da Institução"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
    End With
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    
    Set mProcessamento = New Processamento

    cNomeMapa = "CBESQ016.RPT"
On Error GoTo TrataErro
   
    ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDFuncTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDFunc.Execute cSql, dbFailOnError
    
    ' apaga registos da TABESTADOCIVIL
    cSql = "DELETE * FROM TABESTADOCIVIL;"
    mBDFuncTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABESTADOCIVIL
    cSql = "INSERT INTO TABESTADOCIVIL IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABESTADOCIVIL"
    mBDFunc.Execute cSql, dbFailOnError
    
    ' apaga registos da TABCATEGORIA
    cSql = "DELETE * FROM TABCATEGORIA;"
    mBDFuncTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABCATEGORIA
    cSql = "INSERT INTO TABCATEGORIA IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABCATEGORIA"
    mBDFunc.Execute cSql, dbFailOnError
    
   ' apaga os registos da Temp32
    cSql = "DELETE * FROM FICHA_FUNCIONARIOS"
    ' apaga o registo em Temp32
    mBDFuncTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_FUNCIONARIOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM FUNCIONARIOS WHERE NUM_FUNCIONARIO=" & lNum_Funcionario
    ' insere o registo em Temp32
    mBDFunc.Execute cSql
                        
    With rptFichaIndividual
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Sócio"
            GoTo SairDoProcedimento
        End If
'        .WindowParentHandle = fFrmMDIPrincipal.hwnd
'        .WindowTitle = "Impressão de Ficha Funcionário"
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
        .DataFiles(1) = cBDComNomeUtilizador
        .DataFiles(2) = cBDComNomeUtilizador
        .DataFiles(3) = cBDComNomeUtilizador
        .PrintFileLinesPerPage = 60
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
    Call ErrosGerais("Impressão Ficha de Funcionário", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryFuncionario As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' Campos Obrigatórios
    If Trim$(txtNome.Text) = vbNullString Then
        MsgBox "Nome é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Sócio"
        txtNome.SetFocus
        Exit Sub
    End If
On Error GoTo TrataErro
    ' começa a transação
    mWSFunc.BeginTrans
    If cBotao = "Novo" Then
        Set qryFuncionario = mBDFunc.QueryDefs("FUNCIONARIOS Insere")
        ' parametros de input
        qryFuncionario.Parameters("Num_Funcionario") = txtNumFunc.Value
        qryFuncionario.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryFuncionario.Parameters("Nome") = txtNome.Text
        qryFuncionario.Parameters("Morada") = txtMorada.Text
        qryFuncionario.Parameters("Local") = txtLocal.Text
        qryFuncionario.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryFuncionario.Parameters("Telefone") = txtTel_Morada.Text
        qryFuncionario.Parameters("Telemovel") = txtTelemovel.Text
        qryFuncionario.Parameters("Cod_Estado_Civil") = cCodificaEstadoCivil(cboEstado_Civil.Text)
        qryFuncionario.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        qryFuncionario.Parameters("Data_Admissao") = dcboData_Admissao.DateValue
        qryFuncionario.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryFuncionario.Parameters("Num_BI") = txtNum_BI.Text
        qryFuncionario.Parameters("Data_BI") = dcboData_BI.DateValue
        qryFuncionario.Parameters("Local_Emi_BI") = txtLocal_Emi_BI.Text
        qryFuncionario.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        qryFuncionario.Parameters("Cod_Categoria") = cCodificaCategoria(cboCategoria.Text)
        qryFuncionario.Parameters("Utiliz") = gUtilizador.Nome
      ElseIf cBotao = "Altera" Then
        Set qryFuncionario = mBDFunc.QueryDefs("FUNCIONARIOS Altera")
        ' parametros de input
        qryFuncionario.Parameters("Num_Funcionario") = lNum_Funcionario
        qryFuncionario.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryFuncionario.Parameters("Nome") = txtNome.Text
        qryFuncionario.Parameters("Morada") = txtMorada.Text
        qryFuncionario.Parameters("Local") = txtLocal.Text
        qryFuncionario.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryFuncionario.Parameters("Telefone") = txtTel_Morada.Text
        qryFuncionario.Parameters("Telemovel") = txtTelemovel.Text
        qryFuncionario.Parameters("Cod_Estado_Civil") = cCodificaEstadoCivil(cboEstado_Civil.Text)
        qryFuncionario.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        qryFuncionario.Parameters("Data_Admissao") = dcboData_Admissao.DateValue
        qryFuncionario.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryFuncionario.Parameters("Num_BI") = txtNum_BI.Text
        qryFuncionario.Parameters("Data_BI") = dcboData_BI.DateValue
        qryFuncionario.Parameters("Local_Emi_BI") = txtLocal_Emi_BI.Text
        qryFuncionario.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        qryFuncionario.Parameters("Cod_Categoria") = cCodificaCategoria(cboCategoria.Text)
        qryFuncionario.Parameters("Utiliz") = gUtilizador.Nome
        qryFuncionario.Parameters("NovoNum_Funcionario") = txtNumFunc.Value
    End If
    ' executa a query
    qryFuncionario.Execute dbFailOnError
    
    mWSFunc.CommitTrans
    ' faz o refresh da frmGestaoFunc
    frmGestaoFuncionarios.datFuncionarios.Refresh
    frmGestaoFuncionarios.sgrdGestaoFuncionarios.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSFunc.Rollback
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


Private Sub dcboData_Demissao_DropDown()
    If Not IsDate(dcboData_Demissao.Text) Then
        dcboData_Demissao.DateValue = Date
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
        lNum_Funcionario = CLng(frmGestaoFuncionarios.sgrdGestaoFuncionarios.Columns(0).Text)
    End If
    
    If cBotao = "Ficha" Then
        Me.Caption = Me.Caption & " Nº " & lNum_Funcionario
        cmdOK.Visible = False
        cmdImprimir.Visible = True
    ElseIf cBotao = "Novo" Then
        Me.Caption = "Nova " & Me.Caption
        cmdImprimir.Visible = False
   ElseIf cBotao = "Altera" Then
        Me.Caption = "Alteração na " & Me.Caption & " Nº " & lNum_Funcionario
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
    For Each mBD In mWSFunc.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSFunc = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSFunc = DBEngine.CreateWorkspace("Func", gUtilizador.Nome, gUtilizador.Password)
    Set mBDFunc = mWSFunc.OpenDatabase(cBD_Path & cNomeBD)
    If cBotao = "Ficha" Then
        Set mBDFuncTemp = mWSFunc.OpenDatabase(cBDComNomeUtilizador)
    End If
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Ficha Func-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function



Private Sub CamposEnabledDisabled()
    If cBotaoOrigem = "Ficha" Then
        'Poe os campos locked só para consulta da ficha
        txtNumFunc.Locked = True
        txtNumFunc.TabStop = False
        txtNumFunc.BackColor = &H8000000F
        
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
        
        dcboData_Nasc.Locked = True
        dcboData_Nasc.CalDropDown = False
        dcboData_Nasc.TabStop = False
        
        cboEstado_Civil.Visible = False
        txtEstado_Civil.Visible = True
        txtEstado_Civil.Locked = True
        txtEstado_Civil.TabStop = False
        txtEstado_Civil.BackColor = &H8000000F
        
        cboInstituicao.Visible = False
        txtInstituicao.Visible = True
        txtInstituicao.Locked = True
        txtInstituicao.TabStop = False
        txtInstituicao.BackColor = &H8000000F

        dcboData_Admissao.Locked = True
        dcboData_Admissao.CalDropDown = False
        dcboData_Admissao.TabStop = False
        
        dcboData_Demissao.Locked = True
        dcboData_Demissao.CalDropDown = False
        dcboData_Demissao.TabStop = False
        
        txtNum_BI.Locked = True
        txtNum_BI.TabStop = False
        txtNum_BI.BackColor = &H8000000F
        
        dcboData_BI.Locked = True
        dcboData_BI.CalDropDown = False
        dcboData_BI.TabStop = False
        
        txtLocal_Emi_BI.Locked = True
        txtLocal_Emi_BI.TabStop = False
        txtLocal_Emi_BI.BackColor = &H8000000F
        
        
        
        
        
        txtNum_Contribuinte.Locked = True
        txtNum_Contribuinte.TabStop = False
        txtNum_Contribuinte.BackColor = &H8000000F
        
        txtSeguranca_Social.Locked = True
        txtSeguranca_Social.TabStop = False
        txtSeguranca_Social.BackColor = &H8000000F
        
        cboCategoria.Visible = False
        txtCategoria.Visible = True
        txtCategoria.Locked = True
        txtCategoria.TabStop = False
        txtCategoria.BackColor = &H8000000F
    Else
        cboEstado_Civil.Visible = True
        txtEstado_Civil.Visible = False
        
        cboInstituicao.Visible = True
        txtInstituicao.Visible = False
        
        cboCategoria.Visible = True
        txtCategoria.Visible = False
    End If
End Sub

Public Sub CamposLimpaCarrega()
    Dim recFuncionario As Recordset
    
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM FUNCIONARIOS WHERE NUM_FUNCIONARIO=" & CLng(frmGestaoFuncionarios.sgrdGestaoFuncionarios.Columns(0).Value)
        Set recFuncionario = mBDFunc.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        'Poe os campos com os dados do sócio
        If Dir(cApl_Path & "\FOTOS\" & recFuncionario!NUM_FUNCIONARIO & ".BMP") = "" Then
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        Else
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\" & recFuncionario!NUM_SOCIO & ".BMP")
        End If
        txtNumFunc.Text = vFiltraCamposNulos(recFuncionario!NUM_FUNCIONARIO)
        txtNome.Text = vFiltraCamposNulos(recFuncionario!Nome)
        txtMorada.Text = vFiltraCamposNulos(recFuncionario!MORADA)
        txtCodigoPostal.Text = vFiltraCamposNulos(recFuncionario!COD_POSTAL)
        txtLocal.Text = vFiltraCamposNulos(recFuncionario!LOCAL)
        txtTel_Morada.Text = vFiltraCamposNulos(recFuncionario!TELEFONE)
        txtTelemovel.Text = vFiltraCamposNulos(recFuncionario!TELEMOVEL)
        dcboData_Nasc.Text = vFiltraCamposNulos(recFuncionario!DATA_NASC)
        If cBotao = "Ficha" Then
            txtEstado_Civil.Text = cDescodificaEstadoCivil(recFuncionario!COD_ESTADO_CIVIL)
            txtInstituicao.Text = cDescodificaInstituicao(recFuncionario!COD_INST)
            txtCategoria.Text = cDescodificaCategoria(recFuncionario!COD_CATEGORIA)
        ElseIf cBotao = "Altera" Then
            cboEstado_Civil.Text = cDescodificaEstadoCivil(recFuncionario!COD_ESTADO_CIVIL)
            cboInstituicao.Text = cDescodificaInstituicao(recFuncionario!COD_INST)
            cboCategoria.Text = cDescodificaCategoria(recFuncionario!COD_CATEGORIA)
        End If
        dcboData_Admissao.Text = vFiltraCamposNulos(recFuncionario!DATA_ADMISSAO)
        dcboData_Demissao.Text = vFiltraCamposNulos(recFuncionario!DATA_DEMISSAO)
        txtNum_BI.Text = vFiltraCamposNulos(recFuncionario!NUM_BI)
        dcboData_BI.Text = vFiltraCamposNulos(recFuncionario!DATA_BI)
        txtLocal_Emi_BI.Text = vFiltraCamposNulos(recFuncionario!LOCAL_EMI_BI)
        txtNum_Contribuinte.Text = vFiltraCamposNulos(recFuncionario!NUM_CONTRIBUINTE)
        txtSeguranca_Social.Text = vFiltraCamposNulos(recFuncionario!NUM_SEG_SOCIAL)
        ' fecha o recordset
        recFuncionario.Close
        Set recFuncionario = Nothing
    ElseIf cBotao = "Novo" Then
        'Poe os campos preparados para Novo ficha
        picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        txtNumFunc.Text = lNovoNumFuncionario
        txtNome.Text = vbNullString
        txtMorada.Text = vbNullString
        txtCodigoPostal.Text = vbNullString
        txtLocal.Text = vbNullString
        txtTel_Morada.Text = vbNullString
        txtTelemovel.Text = vbNullString
        dcboData_Nasc.Text = vbNullString
        cboEstado_Civil.Text = vbNullString
        cboInstituicao.Text = vbNullString
        dcboData_Admissao.DateValue = Date
        dcboData_Demissao.Text = vbNullString
        txtNum_BI.Text = vbNullString
        dcboData_BI.Text = vbNullString
        txtLocal_Emi_BI.Text = vbNullString
        txtNum_Contribuinte.Text = vbNullString
        txtSeguranca_Social.Text = vbNullString
        cboCategoria.Text = vbNullString
    End If
End Sub


