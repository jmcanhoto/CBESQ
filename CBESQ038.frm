VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFFichaUtente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Ficha de Utente"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
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
   ScaleHeight     =   6510
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNome_Abrev 
      Height          =   360
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   139
      Top             =   6000
      Width           =   5115
   End
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   6165
      Top             =   5745
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   141
      Tag             =   "2010"
      Top             =   5565
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   142
      Tag             =   "2003"
      Top             =   5565
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   140
      Tag             =   "2002"
      Top             =   5565
      Width           =   1200
   End
   Begin TabDlg.SSTab tabFichaUtente 
      Height          =   5385
      Left            =   150
      TabIndex        =   143
      Top             =   120
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9499
      _Version        =   393216
      Style           =   1
      Tabs            =   8
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
      TabCaption(0)   =   "&1 - Utente"
      TabPicture(0)   =   "CBESQ038.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picFoto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraUtente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2 - Contactos"
      TabPicture(1)   =   "CBESQ038.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMae"
      Tab(1).Control(1)=   "fraPai"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3 - Outros Dados"
      TabPicture(2)   =   "CBESQ038.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBI"
      Tab(2).Control(1)=   "fraNaturalidade"
      Tab(2).Control(2)=   "fraMensalidade"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&4 - Outros Dados"
      TabPicture(3)   =   "CBESQ038.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraNumeros2"
      Tab(3).Control(1)=   "fraEnc_Edu"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "fraNumeros"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&5 - Outros Dados"
      TabPicture(4)   =   "CBESQ038.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraOutrosDados"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Observações"
      TabPicture(5)   =   "CBESQ038.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraInstituicao"
      Tab(5).Control(1)=   "fraObservacoes"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "&7 - Rendimentos"
      TabPicture(6)   =   "CBESQ038.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraRendimentos"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&8 - Despesas"
      TabPicture(7)   =   "CBESQ038.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame4"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   2280
         Left            =   -73920
         TabIndex        =   127
         Top             =   720
         Width           =   7170
         Begin GTMaskNum.GTMaskNum txtD_IRS 
            Height          =   360
            Left            =   5430
            TabIndex        =   129
            Top             =   270
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
         Begin GTMaskNum.GTMaskNum txtD_SS 
            Height          =   360
            Left            =   5430
            TabIndex        =   131
            Top             =   720
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
         Begin GTMaskNum.GTMaskNum txtD_IMI 
            Height          =   360
            Left            =   5400
            TabIndex        =   133
            Top             =   1200
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
         Begin GTMaskNum.GTMaskNum txtD_JEH 
            Height          =   360
            Left            =   5400
            TabIndex        =   135
            Top             =   1680
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I R S (Anual) ........................................................................................."
            Height          =   240
            Index           =   64
            Left            =   150
            TabIndex        =   128
            Top             =   300
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Segurança Social (Anual) .............................................................."
            Height          =   240
            Index           =   63
            Left            =   150
            TabIndex        =   130
            Top             =   750
            Width           =   5115
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I M I (Anual) ........................................................................................."
            Height          =   240
            Index           =   62
            Left            =   150
            TabIndex        =   132
            Top             =   1230
            Width           =   5070
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Juros / Amortização Empréstimo Habitação (Anual) ........."
            Height          =   240
            Index           =   61
            Left            =   150
            TabIndex        =   134
            Top             =   1710
            Width           =   5010
         End
      End
      Begin VB.Frame fraRendimentos 
         Height          =   4560
         Left            =   -73920
         TabIndex        =   108
         Top             =   720
         Width           =   7170
         Begin VB.TextBox txtR_O_DESC 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   125
            Top             =   4080
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtR_TD 
            Height          =   360
            Left            =   5430
            TabIndex        =   110
            Top             =   270
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
         Begin GTMaskNum.GTMaskNum txtR_P 
            Height          =   360
            Left            =   5430
            TabIndex        =   112
            Top             =   720
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
         Begin GTMaskNum.GTMaskNum txtR_PA 
            Height          =   360
            Left            =   5400
            TabIndex        =   114
            Top             =   1200
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
         Begin GTMaskNum.GTMaskNum txtR_TI 
            Height          =   360
            Left            =   5400
            TabIndex        =   116
            Top             =   1680
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
         Begin GTMaskNum.GTMaskNum txtR_R 
            Height          =   360
            Left            =   5400
            TabIndex        =   118
            Top             =   2160
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
         Begin GTMaskNum.GTMaskNum txtR_RCI 
            Height          =   360
            Left            =   5400
            TabIndex        =   120
            Top             =   2640
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
         Begin GTMaskNum.GTMaskNum txtR_SD 
            Height          =   360
            Left            =   5400
            TabIndex        =   122
            Top             =   3120
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
         Begin GTMaskNum.GTMaskNum txtR_AF 
            Height          =   360
            Left            =   5400
            TabIndex        =   124
            Top             =   3600
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
         Begin GTMaskNum.GTMaskNum txtR_O 
            Height          =   360
            Left            =   5400
            TabIndex        =   126
            Top             =   4080
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Abono (Anual) ...................................................................................."
            Height          =   240
            Index           =   60
            Left            =   150
            TabIndex        =   123
            Top             =   3630
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Subsidio Desemprego (Anual) ....................................................."
            Height          =   240
            Index           =   59
            Left            =   150
            TabIndex        =   121
            Top             =   3150
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "R C I (Anual) ........................................................................................"
            Height          =   240
            Index           =   58
            Left            =   150
            TabIndex        =   119
            Top             =   2670
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Rendas (Anual) .................................................................................."
            Height          =   240
            Index           =   57
            Left            =   150
            TabIndex        =   117
            Top             =   2190
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Independente (Anual) ...................................................."
            Height          =   240
            Index           =   56
            Left            =   150
            TabIndex        =   115
            Top             =   1710
            Width           =   5175
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões Alimentos (Anual) ............................................................"
            Height          =   240
            Index           =   55
            Left            =   150
            TabIndex        =   113
            Top             =   1230
            Width           =   5160
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões (Anual) ................................................................................"
            Height          =   240
            Index           =   54
            Left            =   150
            TabIndex        =   111
            Top             =   750
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Dependente (Anual) ......................................................"
            Height          =   240
            Index           =   50
            Left            =   150
            TabIndex        =   109
            Top             =   300
            Width           =   5145
         End
      End
      Begin VB.Frame fraNumeros2 
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
         Height          =   1770
         Left            =   -74790
         TabIndex        =   89
         Top             =   2805
         Width           =   2205
         Begin VB.TextBox txtSNS 
            Height          =   360
            Left            =   150
            MaxLength       =   20
            TabIndex        =   91
            Top             =   540
            Width           =   1950
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "S.N.S."
            Height          =   240
            Index           =   17
            Left            =   150
            TabIndex        =   90
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame fraObservacoes 
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
         Height          =   1545
         Left            =   -74760
         TabIndex        =   98
         Top             =   990
         Width           =   8790
         Begin VB.TextBox txtObservacoes 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   100
            Top             =   540
            Width           =   8460
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   240
            Index           =   14
            Left            =   150
            TabIndex        =   99
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Frame fraInstituicao 
         Caption         =   " Instituição e Valência"
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
         Height          =   1905
         Left            =   -74760
         TabIndex        =   101
         Top             =   2640
         Width           =   8790
         Begin VB.TextBox txtSalas 
            Height          =   360
            Left            =   150
            TabIndex        =   107
            Top             =   1185
            Width           =   5250
         End
         Begin VB.TextBox txtInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   104
            Top             =   525
            Width           =   5250
         End
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   330
            Left            =   150
            TabIndex        =   103
            Top             =   540
            Width           =   5265
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   9287
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin SSDataWidgets_B.SSDBCombo cboSalas 
            Height          =   330
            Left            =   150
            TabIndex        =   106
            Top             =   1200
            Width           =   5265
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   9287
            _ExtentY        =   582
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Valência"
            Height          =   240
            Index           =   28
            Left            =   150
            TabIndex        =   105
            Top             =   900
            Width           =   795
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Instituição"
            Height          =   240
            Index           =   27
            Left            =   150
            TabIndex        =   102
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame fraOutrosDados 
         Height          =   1125
         Left            =   -74730
         TabIndex        =   95
         Top             =   825
         Width           =   7170
         Begin GTMaskNum.GTMaskNum txtAgregado 
            Height          =   360
            Left            =   5775
            TabIndex        =   97
            Top             =   420
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Text            =   "1"
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
            Caption         =   "Agregado"
            Height          =   240
            Index           =   44
            Left            =   150
            TabIndex        =   96
            Top             =   480
            Width           =   915
         End
      End
      Begin VB.Frame fraEnc_Edu 
         Caption         =   " Grau de Dependência "
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
         Left            =   -72390
         TabIndex        =   92
         Top             =   2805
         Width           =   6390
         Begin VB.TextBox txtDependencia 
            Height          =   960
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   94
            Top             =   585
            Width           =   6075
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Dependência"
            Height          =   240
            Index           =   15
            Left            =   195
            TabIndex        =   93
            Top             =   345
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Dados Médicos"
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
         Left            =   -72390
         TabIndex        =   84
         Top             =   870
         Width           =   6390
         Begin VB.TextBox txtCentro_Saude 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   86
            Top             =   570
            Width           =   6075
         End
         Begin VB.TextBox txtMedico_Familia 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   88
            Top             =   1230
            Width           =   6075
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Saúde"
            Height          =   240
            Index           =   45
            Left            =   150
            TabIndex        =   85
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Médico de Família"
            Height          =   240
            Index           =   46
            Left            =   150
            TabIndex        =   87
            Top             =   990
            Width           =   1665
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
         Left            =   -74805
         TabIndex        =   79
         Top             =   870
         Width           =   2250
         Begin VB.TextBox txtSeguranca_Social 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   83
            Top             =   1200
            Width           =   1950
         End
         Begin VB.TextBox txtNum_Contribuinte 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   81
            Top             =   540
            Width           =   1950
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Seg. Social"
            Height          =   240
            Index           =   12
            Left            =   150
            TabIndex        =   82
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   240
            Index           =   13
            Left            =   150
            TabIndex        =   80
            Top             =   300
            Width           =   1050
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
         Height          =   1815
         Left            =   -70905
         TabIndex        =   59
         Top             =   1005
         Width           =   4890
         Begin VB.TextBox txtNum_BI 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   61
            Top             =   540
            Width           =   1500
         End
         Begin VB.TextBox txtLocal_Emi_BI 
            Height          =   360
            Left            =   120
            MaxLength       =   40
            TabIndex        =   65
            Top             =   1200
            Width           =   4575
         End
         Begin GTMaskDate.GTMaskDate dcboData_BI 
            Height          =   375
            Left            =   2100
            TabIndex        =   63
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
            Caption         =   "Local de Emissão"
            Height          =   240
            Index           =   41
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   1620
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão"
            Height          =   240
            Index           =   40
            Left            =   2100
            TabIndex        =   62
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   240
            Index           =   16
            Left            =   150
            TabIndex        =   60
            Top             =   300
            Width           =   225
         End
      End
      Begin VB.Frame fraNaturalidade 
         Caption         =   " Naturalidade "
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
         Left            =   -74820
         TabIndex        =   54
         Top             =   1020
         Width           =   3750
         Begin VB.TextBox txtConcelho 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   58
            Top             =   1230
            Width           =   3390
         End
         Begin VB.TextBox txtFreguesia 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   56
            Top             =   570
            Width           =   3390
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Concellho"
            Height          =   240
            Index           =   21
            Left            =   150
            TabIndex        =   57
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Freguesia"
            Height          =   240
            Index           =   20
            Left            =   150
            TabIndex        =   55
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Frame fraMensalidade 
         Caption         =   " Mensalidade "
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
         Left            =   -74835
         TabIndex        =   66
         Top             =   2880
         Width           =   8820
         Begin VB.OptionButton optTipo_Mensalidade 
            Caption         =   "Automático"
            Height          =   330
            Index           =   0
            Left            =   300
            TabIndex        =   67
            Top             =   360
            Width           =   1410
         End
         Begin VB.OptionButton optTipo_Mensalidade 
            Caption         =   "Manual"
            Height          =   330
            Index           =   1
            Left            =   2265
            TabIndex        =   68
            Top             =   360
            Value           =   -1  'True
            Width           =   1410
         End
         Begin GTMaskNum.GTMaskNum txtMensalidade 
            Height          =   360
            Left            =   5595
            TabIndex        =   76
            Top             =   1200
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
         Begin GTMaskNum.GTMaskNum txtComparticipacao 
            Height          =   360
            Left            =   7125
            TabIndex        =   78
            Top             =   1200
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
         Begin GTMaskNum.GTMaskNum txtProxMensalidade 
            Height          =   360
            Left            =   1710
            TabIndex        =   72
            Top             =   1170
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
         Begin GTMaskNum.GTMaskNum txtMensalidade_Base 
            Height          =   360
            Left            =   4035
            TabIndex        =   74
            Top             =   1200
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
         Begin GTMaskNum.GTMaskNum txtProxMensalidade_Base 
            Height          =   360
            Left            =   150
            TabIndex        =   70
            Top             =   1170
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mensalidade"
            Height          =   240
            Index           =   25
            Left            =   5595
            TabIndex        =   75
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Comparticipação"
            Height          =   240
            Index           =   26
            Left            =   7125
            TabIndex        =   77
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. Mensal."
            Height          =   240
            Index           =   32
            Left            =   1710
            TabIndex        =   71
            Top             =   930
            Width           =   1200
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mensal. Base"
            Height          =   240
            Index           =   33
            Left            =   4035
            TabIndex        =   73
            Top             =   960
            Width           =   1230
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. Mens. Base"
            Height          =   240
            Index           =   34
            Left            =   150
            TabIndex        =   69
            Top             =   930
            Width           =   1560
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
         TabIndex        =   27
         Top             =   3480
         Width           =   2235
         Begin GTMaskDate.GTMaskDate dcboData_Admissao 
            Height          =   375
            Left            =   150
            TabIndex        =   29
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
            TabIndex        =   31
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
            Caption         =   "Admissão"
            Height          =   240
            Index           =   30
            Left            =   150
            TabIndex        =   28
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Demissão"
            Height          =   240
            Index           =   29
            Left            =   135
            TabIndex        =   30
            Top             =   960
            Width           =   930
         End
      End
      Begin VB.Frame fraMae 
         Caption         =   " 2º Contacto Familiar "
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
         TabIndex        =   43
         Top             =   2820
         Width           =   8880
         Begin VB.TextBox txtParentesco_Contacto2 
            Height          =   360
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1230
            Width           =   3795
         End
         Begin VB.TextBox txtTelemovel_Contacto2 
            Height          =   360
            Left            =   3225
            MaxLength       =   12
            TabIndex        =   51
            Top             =   1230
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Contacto2 
            Height          =   360
            Left            =   105
            MaxLength       =   12
            TabIndex        =   47
            Top             =   1215
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Emp_Contacto2 
            Height          =   360
            Left            =   1665
            MaxLength       =   12
            TabIndex        =   49
            Top             =   1215
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Contacto2 
            Height          =   360
            Left            =   135
            MaxLength       =   50
            TabIndex        =   45
            Top             =   540
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   39
            Left            =   4800
            TabIndex        =   52
            Top             =   990
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   38
            Left            =   3225
            TabIndex        =   50
            Top             =   990
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   37
            Left            =   105
            TabIndex        =   46
            Top             =   975
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   10
            Left            =   1665
            TabIndex        =   48
            Top             =   975
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   44
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame fraPai 
         Caption         =   " 1º Contacto Familiar "
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
         TabIndex        =   32
         Top             =   900
         Width           =   8880
         Begin VB.TextBox txtParentesco_Contacto1 
            Height          =   360
            Left            =   4830
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1215
            Width           =   3795
         End
         Begin VB.TextBox txtTelemovel_Contacto1 
            Height          =   360
            Left            =   3255
            MaxLength       =   12
            TabIndex        =   40
            Top             =   1215
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Contacto1 
            Height          =   360
            Left            =   135
            MaxLength       =   12
            TabIndex        =   36
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Emp_Contacto1 
            Height          =   360
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   38
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Contacto1 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   34
            Top             =   540
            Width           =   8475
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   36
            Left            =   4830
            TabIndex        =   41
            Top             =   975
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   11
            Left            =   3255
            TabIndex        =   39
            Top             =   975
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   8
            Left            =   135
            TabIndex        =   35
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   33
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   7
            Left            =   1695
            TabIndex        =   37
            Top             =   960
            Width           =   1320
         End
      End
      Begin VB.Frame fraUtente 
         Caption         =   " Utente "
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
         Height          =   4320
         Left            =   150
         TabIndex        =   0
         Top             =   900
         Width           =   6600
         Begin VB.Frame Frame3 
            Caption         =   "Casado"
            Height          =   660
            Left            =   135
            TabIndex        =   20
            Top             =   3585
            Width           =   6345
            Begin VB.CheckBox chkEstaNoCBESQ 
               Caption         =   "Conjugue está no CAIF"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3675
               TabIndex        =   25
               Top             =   270
               Width           =   2490
            End
            Begin VB.OptionButton optCasado 
               Height          =   300
               Index           =   0
               Left            =   1125
               TabIndex        =   21
               Top             =   255
               Width           =   210
            End
            Begin VB.OptionButton optCasado 
               Height          =   300
               Index           =   1
               Left            =   2025
               TabIndex        =   23
               Top             =   255
               Value           =   -1  'True
               Width           =   210
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Sim"
               Height          =   240
               Index           =   51
               Left            =   1470
               TabIndex        =   22
               Top             =   255
               Width           =   345
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Não"
               Height          =   240
               Index           =   52
               Left            =   2385
               TabIndex        =   24
               Top             =   255
               Width           =   390
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Sexo"
            Height          =   735
            Left            =   2325
            TabIndex        =   5
            Top             =   1005
            Width           =   4110
            Begin VB.OptionButton optSexo 
               Height          =   300
               Index           =   1
               Left            =   2115
               TabIndex        =   8
               Top             =   300
               Width           =   210
            End
            Begin VB.OptionButton optSexo 
               Height          =   300
               Index           =   0
               Left            =   675
               TabIndex        =   6
               Top             =   300
               Width           =   210
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Feminino"
               Height          =   240
               Index           =   19
               Left            =   2385
               TabIndex        =   9
               Top             =   300
               Width           =   825
            End
            Begin VB.Label lblTexto 
               AutoSize        =   -1  'True
               Caption         =   "Masculino"
               Height          =   240
               Index           =   18
               Left            =   945
               TabIndex        =   7
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Left            =   1770
            MaxLength       =   12
            TabIndex        =   19
            Top             =   3210
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   17
            Top             =   3210
            Width           =   1500
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   13
            Top             =   2550
            Width           =   1110
         End
         Begin VB.TextBox txtMorada 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   11
            Top             =   1890
            Width           =   6300
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   15
            Top             =   2550
            Width           =   5160
         End
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   2
            Top             =   570
            Width           =   6300
         End
         Begin GTMaskDate.GTMaskDate dcboData_Nasc 
            Height          =   360
            Left            =   150
            TabIndex        =   4
            Top             =   1230
            Width           =   1935
            _Version        =   65537
            _ExtentX        =   3413
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
            CalDayCaption1  =   "Dom"
            CalDayCaption2  =   "Seg"
            CalDayCaption3  =   "Ter"
            CalDayCaption4  =   "Qua"
            CalDayCaption5  =   "Qui"
            CalDayCaption6  =   "Sex"
            CalDayCaption7  =   "Sáb"
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
            Caption         =   "Telemovel"
            Height          =   240
            Index           =   35
            Left            =   1770
            TabIndex        =   18
            Top             =   2970
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   16
            Top             =   2970
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   10
            Top             =   1650
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   12
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   4
            Left            =   1305
            TabIndex        =   14
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nascimento"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   990
            Width           =   1845
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   300
            Width           =   555
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
         Height          =   1800
         Left            =   7200
         ScaleHeight     =   1740
         ScaleWidth      =   1560
         TabIndex        =   26
         Top             =   1020
         Width           =   1620
      End
   End
   Begin GTMaskNum.GTMaskNum txtNum_Cama 
      Height          =   360
      Left            =   1500
      TabIndex        =   137
      Top             =   5565
      Width           =   900
      _Version        =   65536
      _ExtentX        =   1587
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
      Text            =   "0"
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
      Caption         =   "Nº de Cama"
      Height          =   240
      Index           =   49
      Left            =   240
      TabIndex        =   136
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nome Abrev."
      Height          =   240
      Index           =   48
      Left            =   240
      TabIndex        =   138
      Top             =   6060
      Width           =   1185
   End
End
Attribute VB_Name = "frmCAIFFichaUtente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSUtente As Workspace
Dim mBDUtente As Database
Dim mBDUtenteTemp As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lNum_Utente
Dim cNomeMapa

Dim dProx_Comparticipacao
Dim dProx_Mensalidade_Base
'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSUtente = DBEngine.CreateWorkspace(cBotao & "Utente", gUtilizador.Nome, gUtilizador.Password)
    Set mBDUtente = mWSUtente.OpenDatabase(cBD_Path & cNomeBD)
    If cBotao = "Ficha" Then
        Set mBDUtenteTemp = mWSUtente.OpenDatabase(cBDComNomeUtilizador)
    End If
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Ficha Utente-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function
Private Sub CamposEnabledDisabled()
    If cBotao = "Ficha" Then
        'Poe os campos Visiveis para a Ficha
        txtNome.Locked = True
        txtNome.TabStop = False
        txtNome.BackColor = &H8000000F
        
        dcboData_Nasc.Locked = True
        dcboData_Nasc.CalDropDown = False
        dcboData_Nasc.TabStop = False
        
        optSexo(0).Enabled = False
        optSexo(1).Enabled = False
        
        txtMorada.Locked = True
        txtMorada.TabStop = False
        txtMorada.BackColor = &H8000000F
        
        txtCodigoPostal.Locked = True
        txtCodigoPostal.TabStop = False
        txtCodigoPostal.BackColor = &H8000000F
        
        txtLocal.Locked = True
        txtLocal.TabStop = False
        txtLocal.BackColor = &H8000000F
        
        txtTelefone.Locked = True
        txtTelefone.TabStop = False
        txtTelefone.BackColor = &H8000000F
        
        txtTelemovel.Locked = True
        txtTelemovel.TabStop = False
        txtTelemovel.BackColor = &H8000000F
        
        dcboData_Admissao.Locked = True
        dcboData_Admissao.CalDropDown = False
        dcboData_Admissao.TabStop = False
        
        dcboData_Demissao.Locked = True
        dcboData_Demissao.CalDropDown = False
        dcboData_Demissao.TabStop = False
        
        txtNome_Contacto1.Locked = True
        txtNome_Contacto1.TabStop = False
        txtNome_Contacto1.BackColor = &H8000000F
        
        txtTelefone_Contacto1.Locked = True
        txtTelefone_Contacto1.TabStop = False
        txtTelefone_Contacto1.BackColor = &H8000000F
        
        txtTelefone_Emp_Contacto1.Locked = True
        txtTelefone_Emp_Contacto1.TabStop = False
        txtTelefone_Emp_Contacto1.BackColor = &H8000000F
        
        txtTelemovel_Contacto1.Locked = True
        txtTelemovel_Contacto1.TabStop = False
        txtTelemovel_Contacto1.BackColor = &H8000000F
        
        optCasado(0).Enabled = False
        optCasado(1).Enabled = False
       
        chkEstaNoCBESQ.Enabled = False
        
        txtParentesco_Contacto1.Locked = True
        txtParentesco_Contacto1.TabStop = False
        txtParentesco_Contacto1.BackColor = &H8000000F
       
        txtNome_Contacto2.Locked = True
        txtNome_Contacto2.TabStop = False
        txtNome_Contacto2.BackColor = &H8000000F
        
        txtTelefone_Contacto2.Locked = True
        txtTelefone_Contacto2.TabStop = False
        txtTelefone_Contacto2.BackColor = &H8000000F
        
        txtTelefone_Emp_Contacto2.Locked = True
        txtTelefone_Emp_Contacto2.TabStop = False
        txtTelefone_Emp_Contacto2.BackColor = &H8000000F
        
        txtTelemovel_Contacto2.Locked = True
        txtTelemovel_Contacto2.TabStop = False
        txtTelemovel_Contacto2.BackColor = &H8000000F
       
        txtParentesco_Contacto2.Locked = True
        txtParentesco_Contacto2.TabStop = False
        txtParentesco_Contacto2.BackColor = &H8000000F
       
        txtFreguesia.Locked = True
        txtFreguesia.TabStop = False
        txtFreguesia.BackColor = &H8000000F
        
        txtConcelho.Locked = True
        txtConcelho.TabStop = False
        txtConcelho.BackColor = &H8000000F
        
        txtNum_BI.Locked = True
        txtNum_BI.TabStop = False
        txtNum_BI.BackColor = &H8000000F
        
        dcboData_BI.Locked = True
        dcboData_BI.CalDropDown = False
        dcboData_BI.TabStop = False
        
        txtLocal_Emi_BI.Locked = True
        txtLocal_Emi_BI.TabStop = False
        txtLocal_Emi_BI.BackColor = &H8000000F
        
        optTipo_Mensalidade(0).Enabled = False
        optTipo_Mensalidade(1).Enabled = False
        
        txtProxMensalidade.Locked = True
        txtProxMensalidade.TabStop = False
        
        txtProxMensalidade_Base.Locked = True
        txtProxMensalidade_Base.TabStop = False
       
        txtMensalidade_Base.Locked = True
        txtMensalidade_Base.TabStop = False
        
        txtMensalidade.Locked = True
        txtMensalidade.TabStop = False
        
        txtComparticipacao.Locked = True
        txtComparticipacao.TabStop = False
        
        txtNum_Contribuinte.Locked = True
        txtNum_Contribuinte.TabStop = False
        txtNum_Contribuinte.BackColor = &H8000000F
        
        txtSeguranca_Social.Locked = True
        txtSeguranca_Social.TabStop = False
        txtSeguranca_Social.BackColor = &H8000000F
        
        txtCentro_Saude.Locked = True
        txtCentro_Saude.TabStop = False
        txtCentro_Saude.BackColor = &H8000000F
       
        txtMedico_Familia.Locked = True
        txtMedico_Familia.TabStop = False
        txtMedico_Familia.BackColor = &H8000000F
        
        txtDependencia.Locked = True
        txtDependencia.TabStop = False
        txtDependencia.BackColor = &H8000000F
        
        txtAgregado.Locked = True
        txtAgregado.TabStop = False
        
        txtObservacoes.Locked = True
        txtObservacoes.TabStop = False
        txtObservacoes.BackColor = &H8000000F
        
        cboInstituicao.Visible = False
        txtInstituicao.Visible = True
        txtInstituicao.Locked = True
        txtInstituicao.TabStop = False
        txtInstituicao.BackColor = &H8000000F
        
        cboSalas.Visible = False
        txtSalas.Visible = True
        txtSalas.Locked = True
        txtSalas.TabStop = False
        txtSalas.BackColor = &H8000000F
        
        txtNum_Cama.Locked = True
        txtNum_Cama.TabStop = False
        txtNum_Cama.BackColor = &H8000000F
    
        txtNome_Abrev.Locked = True
        txtNome_Abrev.TabStop = False
        txtNome_Abrev.BackColor = &H8000000F
    
        txtSNS.Locked = True
        txtSNS.TabStop = False
        txtSNS.BackColor = &H8000000F
    
        ' Novos Campos Rendimentos / Despesas
        txtR_TD.Locked = True
        txtR_TD.TabStop = False

        txtR_P.Locked = True
        txtR_P.TabStop = False

        txtR_PA.Locked = True
        txtR_PA.TabStop = False

        txtR_TI.Locked = True
        txtR_TI.TabStop = False

        txtR_R.Locked = True
        txtR_R.TabStop = False

        txtR_RCI.Locked = True
        txtR_RCI.TabStop = False

        txtR_SD.Locked = True
        txtR_SD.TabStop = False

        txtR_AF.Locked = True
        txtR_AF.TabStop = False

        txtR_O_DESC.Locked = True
        txtR_O_DESC.TabStop = False
        txtR_O_DESC.BackColor = &H8000000F
        
        txtR_O.Locked = True
        txtR_O.TabStop = False
        
        txtD_IRS.Locked = True
        txtD_IRS.TabStop = False
        
        txtD_SS.Locked = True
        txtD_SS.TabStop = False
        
        txtD_IMI.Locked = True
        txtD_IMI.TabStop = False
        
        txtD_JEH.Locked = True
        txtD_JEH.TabStop = False
    
    Else
        
        lblTexto(29).Visible = False
        dcboData_Demissao.Visible = False
        
        cboInstituicao.Visible = True
        txtInstituicao.Visible = False
        
        cboSalas.Visible = True
        txtSalas.Visible = False
    End If
End Sub


Public Sub CamposLimpaCarrega()
    Dim recUtente As Recordset
    ' Ficha
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM UTENTES_IDOSOS WHERE NUM_UTENTE=" & lNum_Utente
        Set recUtente = mBDUtente.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        'Poe os campos com os dados do Utente
        If Dir(cApl_Path & "\FOTOS\U" & recUtente!NUM_UTENTE & ".BMP") = "" Then
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        Else
            picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\U" & recUtente!NUM_UTENTE & ".BMP")
        End If
        txtNome.Text = vFiltraCamposNulos(recUtente!Nome)
        dcboData_Nasc.Text = vFiltraCamposNulos(recUtente!DATA_NASC)
        If vFiltraCamposNulos(recUtente!SEXO) = "M" Then
            optSexo(0).Value = True
        ElseIf vFiltraCamposNulos(recUtente!SEXO) = "F" Then
            optSexo(1).Value = True
        Else
            optSexo(0).Value = False
            optSexo(1).Value = False
        End If
        txtMorada.Text = vFiltraCamposNulos(recUtente!MORADA)
        txtCodigoPostal.Text = vFiltraCamposNulos(recUtente!COD_POSTAL)
        txtLocal.Text = vFiltraCamposNulos(recUtente!LOCAL)
        txtTelefone.Text = vFiltraCamposNulos(recUtente!TELEFONE1)
        txtTelemovel.Text = vFiltraCamposNulos(recUtente!TELEFONE2)
        dcboData_Admissao.Text = vFiltraCamposNulos(recUtente!DATA_ENTRADA)
        dcboData_Demissao.Text = vFiltraCamposNulos(recUtente!DATA_SAIDA)
        txtNome_Contacto1.Text = vFiltraCamposNulos(recUtente!NOME_CONTACTO_1)
        txtTelefone_Contacto1.Text = vFiltraCamposNulos(recUtente!TEL_CONTACTO_1)
        txtTelefone_Emp_Contacto1.Text = vFiltraCamposNulos(recUtente!TEL_EMP_CONTACTO_1)
        txtTelemovel_Contacto1.Text = vFiltraCamposNulos(recUtente!TLM_CONTACTO_1)
        txtParentesco_Contacto1.Text = vFiltraCamposNulos(recUtente!PARENTE_CONTACTO_1)
        txtNome_Contacto2.Text = vFiltraCamposNulos(recUtente!NOME_CONTACTO_2)
        txtTelefone_Contacto2.Text = vFiltraCamposNulos(recUtente!TEL_CONTACTO_2)
        txtTelefone_Emp_Contacto2.Text = vFiltraCamposNulos(recUtente!TEL_EMP_CONTACTO_2)
        txtTelemovel_Contacto2.Text = vFiltraCamposNulos(recUtente!TLM_CONTACTO_2)
        txtParentesco_Contacto2.Text = vFiltraCamposNulos(recUtente!PARENTE_CONTACTO_2)
        txtFreguesia.Text = vFiltraCamposNulos(recUtente!NAT_FREG)
        txtConcelho.Text = vFiltraCamposNulos(recUtente!NAT_CONC)
        txtNum_BI.Text = vFiltraCamposNulos(recUtente!NUM_BI)
        dcboData_BI.Text = vFiltraCamposNulos(recUtente!DATA_BI)
        txtLocal_Emi_BI.Text = vFiltraCamposNulos(recUtente!LOCAL_EMI_BI)
        txtNum_Contribuinte.Text = vFiltraCamposNulos(recUtente!NUM_CONTRIBUINTE)
        txtSeguranca_Social.Text = vFiltraCamposNulos(recUtente!NUM_SEG_SOCIAL)
        txtCentro_Saude.Text = vFiltraCamposNulos(recUtente!CENTRO_SAUDE)
        txtMedico_Familia.Text = vFiltraCamposNulos(recUtente!MEDICO_FAMILIA)
        txtDependencia.Text = vFiltraCamposNulos(recUtente!GRAU_DEPENDENCIA)
        txtAgregado.Text = vFiltraCamposNulos(recUtente!AGREGADO)
        txtMensalidade_Base.Text = vFiltraCamposNulos(recUtente!MENSALIDADE_BASE)
        txtMensalidade.Text = vFiltraCamposNulos(recUtente!MENSALIDADE)
        txtComparticipacao.Text = vFiltraCamposNulos(recUtente!COMPARTICIPACAO)
        If vFiltraCamposNulos(recUtente!TIPO_MENSALIDADE) Then
            optTipo_Mensalidade(0).Value = True
            ' Calcula a mensalidade com base na tabela de mensalidades
'            Call CalculaMensalidadeAutomaticaCAIF
        Else
            optTipo_Mensalidade(1).Value = True
        End If
        txtProxMensalidade.Text = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE)
        txtProxMensalidade_Base.Text = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE_BASE)
        dProx_Comparticipacao = vFiltraCamposNulos(recUtente!PROX_COMPARTICIPACAO)
        dProx_Mensalidade_Base = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE_BASE)
        txtObservacoes.Text = vFiltraCamposNulos(recUtente!OBSERVACOES)
'        If cBotao = "Ficha" Then
            txtInstituicao.Text = cDescodificaInstituicao(recUtente!COD_INST)
            txtSalas.Text = cDescodificaSala(recUtente!COD_INST, recUtente!COD_SALA)
'        ElseIf cBotao = "Altera" Then
            cboInstituicao.Text = cDescodificaInstituicao(recUtente!COD_INST)
            cboSalas.Text = cDescodificaSala(recUtente!COD_INST, recUtente!COD_SALA)
'        End If
        txtNum_Cama.Text = vFiltraCamposNulos(recUtente!NUM_CAMA)
        txtNome_Abrev.Text = vFiltraCamposNulos(recUtente!NOME_ABREV)
        If vFiltraCamposNulos(recUtente!CASADO) Then
            optCasado(0).Value = True
        Else
            optCasado(1).Value = True
        End If
        If vFiltraCamposNulos(recUtente!NOCBESQ) Then
            chkEstaNoCBESQ.Value = vbChecked
        Else
            chkEstaNoCBESQ.Value = vbUnchecked
        End If
        txtSNS.Text = vFiltraCamposNulos(recUtente!NUM_SNS)
        ' Novos Campos Rendimentos / Despesas
        txtR_TD.Text = vFiltraCamposNulos(recUtente!R_TD)
        txtR_P.Text = vFiltraCamposNulos(recUtente!R_P)
        txtR_PA.Text = vFiltraCamposNulos(recUtente!R_PA)
        txtR_TI.Text = vFiltraCamposNulos(recUtente!R_TI)
        txtR_R.Text = vFiltraCamposNulos(recUtente!R_R)
        txtR_RCI.Text = vFiltraCamposNulos(recUtente!R_RCI)
        txtR_SD.Text = vFiltraCamposNulos(recUtente!R_SD)
        txtR_AF.Text = vFiltraCamposNulos(recUtente!R_AF)
        txtR_O_DESC.Text = vFiltraCamposNulos(recUtente!R_O_DESC)
        txtR_O.Text = vFiltraCamposNulos(recUtente!R_O)
        txtD_IRS.Text = vFiltraCamposNulos(recUtente!D_IRS)
        txtD_SS.Text = vFiltraCamposNulos(recUtente!D_SS)
        txtD_IMI.Text = vFiltraCamposNulos(recUtente!D_IMI)
        txtD_JEH.Text = vFiltraCamposNulos(recUtente!D_JEH)
        
        ' fecha o recordset
        recUtente.Close
        Set recUtente = Nothing
    ElseIf cBotao = "Novo" Then
        'Poe os campos preparados para nova ficha
        picFoto.Picture = LoadPicture(cApl_Path & "\FOTOS\NOFOTO.BMP")
        txtNome.Text = vbNullString
        dcboData_Nasc.Text = vbNullString
        txtMorada.Text = vbNullString
        txtCodigoPostal.Text = vbNullString
        txtLocal.Text = vbNullString
        txtTelefone.Text = vbNullString
        txtTelemovel.Text = vbNullString
        dcboData_Admissao.Text = Date
        dcboData_Demissao.Text = vbNullString
        txtNome_Contacto1.Text = vbNullString
        txtTelefone_Contacto1.Text = vbNullString
        txtTelefone_Emp_Contacto1.Text = vbNullString
        txtTelemovel_Contacto1.Text = vbNullString
        txtParentesco_Contacto1.Text = vbNullString
        txtNome_Contacto2.Text = vbNullString
        txtTelefone_Contacto2.Text = vbNullString
        txtTelefone_Emp_Contacto2.Text = vbNullString
        txtTelemovel_Contacto2.Text = vbNullString
        txtParentesco_Contacto2.Text = vbNullString
        txtFreguesia.Text = vbNullString
        txtConcelho.Text = vbNullString
        txtNum_BI.Text = 0
        dcboData_BI.Text = vbNullString
        txtLocal_Emi_BI.Text = vbNullString
        optTipo_Mensalidade(1).Value = True
        txtMensalidade_Base.Text = 0
        txtMensalidade.Text = 0
        txtProxMensalidade.Text = 0
        txtProxMensalidade_Base.Text = 0
        dProx_Mensalidade_Base = 0
        dProx_Comparticipacao = 0
        txtComparticipacao.Text = 0
        txtNum_Contribuinte.Text = vbNullString
        txtSeguranca_Social.Text = vbNullString
        txtCentro_Saude.Text = vbNullString
        txtMedico_Familia.Text = vbNullString
        txtDependencia.Text = vbNullString
        txtAgregado.Text = 0
        txtObservacoes.Text = vbNullString
        cboInstituicao.Text = vbNullString
        cboSalas.Text = vbNullString
        txtNum_Cama.Text = 0
        txtNome_Abrev.Text = vbNullString
        ' Novos Campos Rendimentos / Despesas
        txtR_TD.Text = 0
        txtR_P.Text = 0
        txtR_PA.Text = 0
        txtR_TI.Text = 0
        txtR_R.Text = 0
        txtR_RCI.Text = 0
        txtR_SD.Text = 0
        txtR_AF.Text = 0
        txtR_O_DESC.Text = vbNullString
        txtR_O.Text = 0
        txtD_IRS.Text = 0
        txtD_SS.Text = 0
        txtD_IMI.Text = 0
        txtD_JEH.Text = 0
  
    End If
End Sub
'Este procedimento activa ou desactiva o default do botao OK
Private Sub BotaoOKDefault(ByVal Propriedade As Boolean)
  cmdOK.Default = Propriedade
End Sub



Private Sub cboInstituicao_Click()
    cboSalas.Text = vbNullString
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

Private Sub cboSalas_DropDown()
    ' carrega a combo
    Call CarregacboSalas(cboSalas, cboInstituicao.Text)
End Sub

Private Sub cboSalas_InitColumnProps()
    With cboSalas
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
        .Columns(0).Caption = "Nome da Sala"
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

'    cNomeMapa = "CBESQ036.RPT"
    cNomeMapa = "CBESQ036B.RPT"
On Error GoTo TrataErro
     ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDUtenteTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDUtente.Execute cSql, dbFailOnError
    
    ' apaga registos da TABSALAS
    cSql = "DELETE * FROM TABSALAS;"
    mBDUtenteTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABSALAS
    cSql = "INSERT INTO TABSALAS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABSALAS"
    mBDUtente.Execute cSql, dbFailOnError
   
   ' apaga os registos da Temp32
    cSql = "DELETE * FROM FICHA_UTENTES_IDOSOS"
    ' apaga o registo em Temp32
    mBDUtenteTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_UTENTES_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES_IDOSOS WHERE NUM_UTENTE=" & lNum_Utente
    ' insere o registo em Temp32
    mBDUtente.Execute cSql
                        
    With rptFichaIndividual
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
            GoTo SairDoProcedimento
        End If
        .WindowParentHandle = fFrmMDIPrincipal.hwnd
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
        .DataFiles(1) = cBDComNomeUtilizador
        .DataFiles(2) = cBDComNomeUtilizador
        .PrintFileLinesPerPage = 60
        .CopiesToPrinter = 1
        'Passa para o Mapa os dados da Empresa
        .Formulas(0) = "Titulo_1='" & Mapa.Titulo_1 & "'"
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_4 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_5 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                      
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Impressão Ficha de Utente", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryUtente As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' Campos Obrigatórios
    If Trim$(txtNome.Text) = vbNullString Then
        MsgBox "Nome é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Utente"
        txtNome.SetFocus
        Exit Sub
    End If
    If Trim$(cboInstituicao.Text) = vbNullString Then
        MsgBox "Instituição é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Utente"
        tabFichaUtente.Tab = 4
        cboInstituicao.SetFocus
        Exit Sub
    End If
    If Trim$(cboSalas.Text) = vbNullString Then
        MsgBox "Sala é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Utente"
        tabFichaUtente.Tab = 4
        cboSalas.SetFocus
        Exit Sub
    End If
On Error GoTo TrataErro
    ' começa a transação
    mWSUtente.BeginTrans
    If cBotao = "Novo" Then
        Set qryUtente = mBDUtente.QueryDefs("CAIF UTENTES Insere")
        ' parametros de input
        qryUtente.Parameters("Num_Utente") = lNovoNumUtenteCAIF
        qryUtente.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryUtente.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)
        qryUtente.Parameters("Nome") = txtNome.Text
        qryUtente.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryUtente.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryUtente.Parameters("Sexo") = "F"
        End If
        qryUtente.Parameters("Morada") = txtMorada.Text
        qryUtente.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryUtente.Parameters("Local") = txtLocal.Text
        qryUtente.Parameters("Telefone1") = txtTelefone.Text
        qryUtente.Parameters("Telefone2") = txtTelemovel.Text
        qryUtente.Parameters("Grau_Dependencia") = txtDependencia.Text
        qryUtente.Parameters("Nome_Contacto1") = txtNome_Contacto1.Text
        qryUtente.Parameters("Tel_Contacto1") = txtTelefone_Contacto1.Text
        qryUtente.Parameters("Tel_Emp_Contacto1") = txtTelefone_Emp_Contacto1.Text
        qryUtente.Parameters("Tlm_Contacto1") = txtTelemovel_Contacto1.Text
        qryUtente.Parameters("Parente_Contacto1") = txtParentesco_Contacto1.Text
        qryUtente.Parameters("Nome_Contacto2") = txtNome_Contacto2.Text
        qryUtente.Parameters("Tel_Contacto2") = txtTelefone_Contacto2.Text
        qryUtente.Parameters("Tel_Emp_Contacto2") = txtTelefone_Emp_Contacto2.Text
        qryUtente.Parameters("Tlm_Contacto2") = txtTelemovel_Contacto2.Text
        qryUtente.Parameters("Parente_Contacto2") = txtParentesco_Contacto2.Text
        qryUtente.Parameters("Nat_Freg") = txtFreguesia.Text
        qryUtente.Parameters("Nat_Conc") = txtConcelho.Text
        qryUtente.Parameters("Num_BI") = txtNum_BI.Text
        qryUtente.Parameters("Data_BI") = dcboData_BI.DateValue
        qryUtente.Parameters("Loc_Emi_BI") = txtLocal_Emi_BI.Text
        qryUtente.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryUtente.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        qryUtente.Parameters("Centro_Saude") = txtCentro_Saude.Text
        qryUtente.Parameters("Medico_Familia") = txtMedico_Familia.Text
        qryUtente.Parameters("Agregado") = txtAgregado.Value
        If optTipo_Mensalidade(0).Value Then
            qryUtente.Parameters("Tipo_Mensalidade") = True
        Else
            qryUtente.Parameters("Tipo_Mensalidade") = False
        End If
        qryUtente.Parameters("Mensalidade") = txtMensalidade.Value
        qryUtente.Parameters("Mensalidade_Base") = txtMensalidade_Base.Value
        qryUtente.Parameters("Comparticipacao") = txtComparticipacao.Value
        qryUtente.Parameters("Prox_Mensalidade") = txtProxMensalidade.Value
        qryUtente.Parameters("Prox_Mensalidade_Base") = txtProxMensalidade_Base.Value
        qryUtente.Parameters("Prox_Comparticipacao") = dProx_Comparticipacao
        qryUtente.Parameters("Observacoes") = txtObservacoes.Text
        qryUtente.Parameters("Num_Cama") = txtNum_Cama.Value
        qryUtente.Parameters("Nome_Abrev") = txtNome_Abrev.Text
        qryUtente.Parameters("Utiliz") = gUtilizador.Nome
        If optCasado(0).Value Then
            qryUtente.Parameters("Casado") = True
        Else
            qryUtente.Parameters("Casado") = False
        End If
        If chkEstaNoCBESQ.Value = vbChecked Then
            qryUtente.Parameters("NoCBESQ") = True
        Else
            qryUtente.Parameters("NoCBESQ") = False
        End If
        qryUtente.Parameters("Num_SNS") = txtSNS.Text
        ' Novos Campos Rendimentos / Despesas
        qryUtente.Parameters("R_TD") = txtR_TD.Value
        qryUtente.Parameters("R_P") = txtR_P.Value
        qryUtente.Parameters("R_PA") = txtR_PA.Value
        qryUtente.Parameters("R_TI") = txtR_TI.Value
        qryUtente.Parameters("R_R") = txtR_R.Value
        qryUtente.Parameters("R_RCI") = txtR_RCI.Value
        qryUtente.Parameters("R_SD") = txtR_SD.Value
        qryUtente.Parameters("R_AF") = txtR_AF.Value
        qryUtente.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryUtente.Parameters("R_O") = txtR_O.Value
        qryUtente.Parameters("D_IRS") = txtD_IRS.Value
        qryUtente.Parameters("D_SS") = txtD_SS.Value
        qryUtente.Parameters("D_IMI") = txtD_IMI.Value
        qryUtente.Parameters("D_JEH") = txtD_JEH.Value
        
      ElseIf cBotao = "Altera" Then
        Set qryUtente = mBDUtente.QueryDefs("CAIF UTENTES Altera")
        ' parametros de input
        qryUtente.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryUtente.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)
        qryUtente.Parameters("Nome") = txtNome.Text
        qryUtente.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryUtente.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryUtente.Parameters("Sexo") = "F"
        End If
        qryUtente.Parameters("Morada") = txtMorada.Text
        qryUtente.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryUtente.Parameters("Local") = txtLocal.Text
        qryUtente.Parameters("Telefone1") = txtTelefone.Text
        qryUtente.Parameters("Telefone2") = txtTelemovel.Text
        qryUtente.Parameters("Grau_Dependencia") = txtDependencia.Text
        qryUtente.Parameters("Nome_Contacto1") = txtNome_Contacto1.Text
        qryUtente.Parameters("Tel_Contacto1") = txtTelefone_Contacto1.Text
        qryUtente.Parameters("Tel_Emp_Contacto1") = txtTelefone_Emp_Contacto1.Text
        qryUtente.Parameters("Tlm_Contacto1") = txtTelemovel_Contacto1.Text
        qryUtente.Parameters("Parente_Contacto1") = txtParentesco_Contacto1.Text
        qryUtente.Parameters("Nome_Contacto2") = txtNome_Contacto2.Text
        qryUtente.Parameters("Tel_Contacto2") = txtTelefone_Contacto2.Text
        qryUtente.Parameters("Tel_Emp_Contacto2") = txtTelefone_Emp_Contacto2.Text
        qryUtente.Parameters("Tlm_Contacto2") = txtTelemovel_Contacto2.Text
        qryUtente.Parameters("Parente_Contacto2") = txtParentesco_Contacto2.Text
        qryUtente.Parameters("Nat_Freg") = txtFreguesia.Text
        qryUtente.Parameters("Nat_Conc") = txtConcelho.Text
        qryUtente.Parameters("Num_BI") = txtNum_BI.Text
        qryUtente.Parameters("Data_BI") = dcboData_BI.DateValue
        qryUtente.Parameters("Loc_Emi_BI") = txtLocal_Emi_BI.Text
        qryUtente.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryUtente.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        qryUtente.Parameters("Centro_Saude") = txtCentro_Saude.Text
        qryUtente.Parameters("Medico_Familia") = txtMedico_Familia.Text
        qryUtente.Parameters("Agregado") = txtAgregado.Value
        If optTipo_Mensalidade(0).Value Then
            qryUtente.Parameters("Tipo_Mensalidade") = True
        Else
            qryUtente.Parameters("Tipo_Mensalidade") = False
        End If
        
        qryUtente.Parameters("Mensalidade") = txtMensalidade.Value
        qryUtente.Parameters("Mensalidade_Base") = txtMensalidade_Base.Value
        qryUtente.Parameters("Comparticipacao") = txtComparticipacao.Value
        qryUtente.Parameters("Prox_Mensalidade") = txtProxMensalidade.Value
        qryUtente.Parameters("Prox_Mensalidade_Base") = txtProxMensalidade_Base.Value
        qryUtente.Parameters("Prox_Comparticipacao") = dProx_Comparticipacao
        qryUtente.Parameters("Observacoes") = txtObservacoes.Text
        qryUtente.Parameters("Utiliz") = gUtilizador.Nome
        qryUtente.Parameters("Num_Utente") = lNum_Utente
        qryUtente.Parameters("Num_Cama") = txtNum_Cama.Value
        qryUtente.Parameters("Nome_Abrev") = txtNome_Abrev.Text
        
        If optCasado(0).Value Then
            qryUtente.Parameters("Casado") = True
        Else
            qryUtente.Parameters("Casado") = False
        End If
        If chkEstaNoCBESQ.Value = vbChecked Then
            qryUtente.Parameters("NoCBESQ") = True
        Else
            qryUtente.Parameters("NoCBESQ") = False
        End If
        qryUtente.Parameters("Num_SNS") = txtSNS.Text
        ' Novos Campos Rendimentos / Despesas
        qryUtente.Parameters("R_TD") = txtR_TD.Value
        qryUtente.Parameters("R_P") = txtR_P.Value
        qryUtente.Parameters("R_PA") = txtR_PA.Value
        qryUtente.Parameters("R_TI") = txtR_TI.Value
        qryUtente.Parameters("R_R") = txtR_R.Value
        qryUtente.Parameters("R_RCI") = txtR_RCI.Value
        qryUtente.Parameters("R_SD") = txtR_SD.Value
        qryUtente.Parameters("R_AF") = txtR_AF.Value
        qryUtente.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryUtente.Parameters("R_O") = txtR_O.Value
        qryUtente.Parameters("D_IRS") = txtD_IRS.Value
        qryUtente.Parameters("D_SS") = txtD_SS.Value
        qryUtente.Parameters("D_IMI") = txtD_IMI.Value
        qryUtente.Parameters("D_JEH") = txtD_JEH.Value

    End If
    ' executa a query
    qryUtente.Execute dbFailOnError
    
    mWSUtente.CommitTrans
    ' faz o refresh da frmGestaoUtentes
    frmCAIFGestaoUtentes.datUtentes.Refresh
    frmCAIFGestaoUtentes.sgrdGestaoUtentes.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSUtente.Rollback
    Call ErrosGerais(cBotao & " Utente", Err.Number, Err.Description)
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



Private Sub Form_Load()
    cBotao = cBotaoOrigem
    If cBotao <> "Novo" Then
        lNum_Utente = CLng(frmCAIFGestaoUtentes.sgrdGestaoUtentes.Columns(0).Text)
    End If

    If cBotao = "Ficha" Then
        Me.Caption = Me.Caption & " Nº " & lNum_Utente
        cmdOK.Visible = False
        cmdImprimir.Visible = True
    ElseIf cBotao = "Novo" Then
        Me.Caption = "Nova " & Me.Caption
        cmdOK.Visible = True
        cmdImprimir.Visible = False
    ElseIf cBotao = "Altera" Then
        Me.Caption = "Alteração na " & Me.Caption & " Nº " & lNum_Utente
        cmdOK.Visible = True
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
    For Each mBD In mWSUtente.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSUtente = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub


Private Sub optCasado_Click(Index As Integer)
    Select Case Index
        Case 0
            chkEstaNoCBESQ.Enabled = True
        Case 1
            chkEstaNoCBESQ.Enabled = False
            chkEstaNoCBESQ.Value = vbUnchecked
    End Select

End Sub

Private Sub optTipo_Mensalidade_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Calcula a mensalidade com base na tabela de mensalidades
            Call CalculaMensalidadeAutomaticaCAIF
        Case 1
            ' O Calculo da mensalidade é manual
            txtProxMensalidade_Base.Text = 0
            txtProxMensalidade.Text = 0
    End Select
End Sub

Public Sub CalculaMensalidadeAutomaticaCAIF()
    Dim recTABMENSALIDADE As Recordset
    Dim cSql
    Dim lValorPerCapita
    Dim lValorPerCapitaCasal
    Dim lValorPerCapita_40
    Dim lValorPerCapita_80
    Dim lValorPerCapitaCasal_80
'    Dim lValorRendMensConjuge_10
    
    
' Calcula Valor Per-Capita segundo formula dada pelo CBESQ
    lValorPerCapita = ((txtR_TD.Value + txtR_P.Value + txtR_PA.Value + txtR_TI.Value + txtR_R.Value + txtR_RCI.Value + txtR_SD.Value + txtR_AF.Value + txtR_O.Value) - (txtD_IRS.Value + txtD_SS.Value + txtD_IMI.Value + txtD_JEH.Value)) _
                        / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
    lValorPerCapitaCasal = ((txtR_TD.Value + txtR_P.Value + txtR_PA.Value + txtR_TI.Value + txtR_R.Value + txtR_RCI.Value + txtR_SD.Value + txtR_AF.Value + txtR_O.Value) - (txtD_IRS.Value + txtD_SS.Value + txtD_IMI.Value + txtD_JEH.Value)) _
                        / IIf(txtAgregado.Text = 0, 2 * 12, txtAgregado.Text * 12)
        
    lValorPerCapita_40 = lValorPerCapita * 0.4
    lValorPerCapita_80 = lValorPerCapita * 0.8
    lValorPerCapitaCasal_80 = lValorPerCapitaCasal * 0.8
'    lValorRendMensConjuge_10 = txtRendMensConj.Value * 0.1
    
    cSql = "SELECT * FROM TABMENSALIDADE_IDOSOS WHERE COD_MENSALIDADE='" & cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) & "'"
    
    Set recTABMENSALIDADE = mBDUtente.OpenRecordset(cSql, dbOpenSnapshot)
   
    If (cCodificaInstituicao(cboInstituicao.Text) = "501" And cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) = "001") Then
        ' Calculo para o Lar
        ' Com Conjuge
        If optCasado(0).Value = True Then
            'conjuge esta no Lar
            If chkEstaNoCBESQ.Value = vbChecked Then
                If (lValorPerCapitaCasal_80) > recTABMENSALIDADE!VALOR_MAX Then
                    txtProxMensalidade_Base.Text = recTABMENSALIDADE!MENSALIDADE
                    txtProxMensalidade.Text = recTABMENSALIDADE!MENSALIDADE
                    dProx_Comparticipacao = 0
                Else
                    txtProxMensalidade_Base.Text = lValorPerCapita_80
                    txtProxMensalidade.Text = lValorPerCapita_80
                    dProx_Comparticipacao = 0
                End If
            'conjuge não esta no Lar
            Else
                If (lValorPerCapita_80) > recTABMENSALIDADE!VALOR_MAX Then
                    txtProxMensalidade_Base.Text = recTABMENSALIDADE!MENSALIDADE
                    txtProxMensalidade.Text = recTABMENSALIDADE!MENSALIDADE
                    dProx_Comparticipacao = 0
                Else
                    txtProxMensalidade_Base.Text = lValorPerCapita_80
                    txtProxMensalidade.Text = lValorPerCapita_80
                    dProx_Comparticipacao = 0
                End If
            End If
        ' Sem Conjuge
        ElseIf optCasado(1).Value = True Then
            If (lValorPerCapita_80) > recTABMENSALIDADE!VALOR_MAX Then
                txtProxMensalidade_Base.Text = recTABMENSALIDADE!MENSALIDADE
                txtProxMensalidade.Text = recTABMENSALIDADE!MENSALIDADE
                dProx_Comparticipacao = 0
            Else
                txtProxMensalidade_Base.Text = lValorPerCapita_80
                txtProxMensalidade.Text = lValorPerCapita_80
                dProx_Comparticipacao = 0
            End If
        End If
    ElseIf (cCodificaInstituicao(cboInstituicao.Text) = "501" And cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) = "002") Then
        ' Calculo para o Centro de Dia
        If lValorPerCapita_40 > recTABMENSALIDADE!VALOR_MAX Then
            txtProxMensalidade_Base.Text = recTABMENSALIDADE!MENSALIDADE
            txtProxMensalidade.Text = recTABMENSALIDADE!MENSALIDADE
            dProx_Comparticipacao = 0
        Else
            txtProxMensalidade_Base.Text = lValorPerCapita_40
            txtProxMensalidade.Text = lValorPerCapita_40
            dProx_Comparticipacao = 0
        End If
    End If
    
    recTABMENSALIDADE.Close
    Set recTABMENSALIDADE = Nothing
End Sub



Private Sub txtAgregado_LostFocus()
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomaticaCAIF
    End If
End Sub
















