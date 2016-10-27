VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCAIFFichaInscricao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAIF - Ficha de Inscrição"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
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
   ScaleHeight     =   7290
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   4200
      Top             =   6420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "2010"
      ToolTipText     =   "Imprimir a Ficha de Inscrição"
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "2003"
      ToolTipText     =   "Sair da Visualização da Ficha de Inscrição"
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   54
      Tag             =   "2002"
      ToolTipText     =   "Inserir / Alterar a Ficha de Inscrição"
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabFichaInscricao 
      Height          =   5940
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tabs            =   9
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
      TabCaption(0)   =   "&1 - Inscrição"
      TabPicture(0)   =   "CBESQ037.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Inscrição"
      TabPicture(1)   =   "CBESQ037.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraNaturalidade"
      Tab(1).Control(1)=   "fraBI"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3 - Inscrição"
      TabPicture(2)   =   "CBESQ037.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraEstado"
      Tab(2).Control(1)=   "fraFiliacao"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&4 - Agregado"
      TabPicture(3)   =   "CBESQ037.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOutrosDados"
      Tab(3).Control(1)=   "fraAgragado"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&5 - Responsável"
      TabPicture(4)   =   "CBESQ037.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraResponsavel"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Observações"
      TabPicture(5)   =   "CBESQ037.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraObservacoes"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&7 - Valência"
      TabPicture(6)   =   "CBESQ037.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraInstituicao"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&8 - Rendimentos"
      TabPicture(7)   =   "CBESQ037.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraRendimentos"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "&9 - Despesas"
      TabPicture(8)   =   "CBESQ037.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame1"
      Tab(8).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2280
         Left            =   -74760
         TabIndex        =   105
         Top             =   1080
         Width           =   7170
         Begin GTMaskNum.GTMaskNum txtD_IRS 
            Height          =   360
            Left            =   5430
            TabIndex        =   106
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
            TabIndex        =   107
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
            TabIndex        =   108
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
            TabIndex        =   109
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
            Index           =   48
            Left            =   150
            TabIndex        =   113
            Top             =   300
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Segurança Social (Anual) .............................................................."
            Height          =   240
            Index           =   49
            Left            =   150
            TabIndex        =   112
            Top             =   750
            Width           =   5115
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I M I (Anual) ........................................................................................."
            Height          =   240
            Index           =   50
            Left            =   150
            TabIndex        =   111
            Top             =   1230
            Width           =   5070
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Juros / Amortização Empréstimo Habitação (Anual) ........."
            Height          =   240
            Index           =   51
            Left            =   150
            TabIndex        =   110
            Top             =   1710
            Width           =   5010
         End
      End
      Begin VB.Frame fraRendimentos 
         Height          =   4560
         Left            =   -74760
         TabIndex        =   86
         Top             =   1080
         Width           =   7170
         Begin VB.TextBox txtR_O_DESC 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   87
            Top             =   4080
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtR_TD 
            Height          =   360
            Left            =   5430
            TabIndex        =   88
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
            TabIndex        =   89
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
            TabIndex        =   90
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
            TabIndex        =   91
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
            TabIndex        =   92
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
            TabIndex        =   93
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
            TabIndex        =   94
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
            TabIndex        =   95
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
            TabIndex        =   96
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
            Index           =   40
            Left            =   150
            TabIndex        =   104
            Top             =   3630
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Subsidio Desemprego (Anual) ....................................................."
            Height          =   240
            Index           =   41
            Left            =   150
            TabIndex        =   103
            Top             =   3150
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "R C I (Anual) ........................................................................................"
            Height          =   240
            Index           =   42
            Left            =   150
            TabIndex        =   102
            Top             =   2670
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Rendas (Anual) .................................................................................."
            Height          =   240
            Index           =   43
            Left            =   150
            TabIndex        =   101
            Top             =   2190
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Independente (Anual) ...................................................."
            Height          =   240
            Index           =   44
            Left            =   150
            TabIndex        =   100
            Top             =   1710
            Width           =   5175
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões Alimentos (Anual) ............................................................"
            Height          =   240
            Index           =   45
            Left            =   150
            TabIndex        =   99
            Top             =   1230
            Width           =   5160
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões (Anual) ................................................................................"
            Height          =   240
            Index           =   46
            Left            =   150
            TabIndex        =   98
            Top             =   750
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Dependente (Anual) ......................................................"
            Height          =   240
            Index           =   47
            Left            =   150
            TabIndex        =   97
            Top             =   300
            Width           =   5145
         End
      End
      Begin VB.Frame fraInstituicao 
         Caption         =   " Instituição e Valência "
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
         TabIndex        =   79
         Top             =   1200
         Width           =   7065
         Begin VB.TextBox txtSalas 
            Height          =   360
            Left            =   150
            TabIndex        =   81
            Top             =   1185
            Width           =   5250
         End
         Begin VB.TextBox txtInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   80
            Top             =   525
            Width           =   5250
         End
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   330
            Left            =   150
            TabIndex        =   82
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
            TabIndex        =   83
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
            Index           =   20
            Left            =   150
            TabIndex        =   85
            Top             =   900
            Width           =   795
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Instituição"
            Height          =   240
            Index           =   39
            Left            =   150
            TabIndex        =   84
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame fraObservacoes 
         Height          =   3720
         Left            =   -74760
         TabIndex        =   74
         Top             =   1080
         Width           =   7170
         Begin VB.TextBox txtObservacoes 
            Height          =   900
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   540
            Width           =   6000
         End
         Begin VB.TextBox txtReservado 
            Height          =   900
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   1860
            Width           =   6000
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   240
            Index           =   14
            Left            =   150
            TabIndex        =   78
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Reservado à Institução"
            Height          =   240
            Index           =   15
            Left            =   150
            TabIndex        =   77
            Top             =   1620
            Width           =   2070
         End
      End
      Begin VB.Frame fraResponsavel 
         Caption         =   " Filiação"
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
         Height          =   3615
         Left            =   -74760
         TabIndex        =   61
         Top             =   1080
         Width           =   7170
         Begin VB.TextBox txtNomeResp 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   67
            Top             =   540
            Width           =   6900
         End
         Begin VB.TextBox txtParentesco 
            Height          =   360
            Left            =   165
            MaxLength       =   50
            TabIndex        =   66
            Top             =   1185
            Width           =   6900
         End
         Begin VB.TextBox txtLocalTrabResp 
            Height          =   360
            Left            =   165
            MaxLength       =   50
            TabIndex        =   65
            Top             =   1875
            Width           =   6900
         End
         Begin VB.TextBox txtResidenciaResp 
            Height          =   360
            Left            =   165
            MaxLength       =   50
            TabIndex        =   64
            Top             =   2565
            Width           =   6900
         End
         Begin VB.TextBox txtTelResidResp 
            Height          =   360
            Left            =   1815
            MaxLength       =   12
            TabIndex        =   63
            Top             =   3180
            Width           =   1500
         End
         Begin VB.TextBox txtTelTrabResp 
            Height          =   360
            Left            =   165
            MaxLength       =   12
            TabIndex        =   62
            Top             =   3180
            Width           =   1500
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   33
            Left            =   150
            TabIndex        =   73
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   34
            Left            =   165
            TabIndex        =   72
            Top             =   945
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Local de Trabalho"
            Height          =   240
            Index           =   35
            Left            =   165
            TabIndex        =   71
            Top             =   1635
            Width           =   1650
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Residência"
            Height          =   240
            Index           =   36
            Left            =   165
            TabIndex        =   70
            Top             =   2325
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Tel. Residência"
            Height          =   240
            Index           =   37
            Left            =   1815
            TabIndex        =   69
            Top             =   2940
            Width           =   1425
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Tel. Trabalho"
            Height          =   240
            Index           =   38
            Left            =   165
            TabIndex        =   68
            Top             =   2940
            Width           =   1215
         End
      End
      Begin VB.Frame fraOutrosDados 
         Height          =   960
         Left            =   -74760
         TabIndex        =   58
         Top             =   3240
         Width           =   7170
         Begin GTMaskNum.GTMaskNum txtAgregado 
            Height          =   360
            Left            =   5775
            TabIndex        =   59
            Top             =   360
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
            Index           =   13
            Left            =   150
            TabIndex        =   60
            Top             =   405
            Width           =   915
         End
      End
      Begin VB.Frame fraAgragado 
         Height          =   2205
         Left            =   -74760
         TabIndex        =   48
         Top             =   1050
         Width           =   7170
         Begin VB.CheckBox chkViveSo 
            Caption         =   "Vive Só "
            Height          =   240
            Left            =   150
            TabIndex        =   49
            Top             =   285
            Width           =   2400
         End
         Begin VB.TextBox txtPorque 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   51
            Top             =   870
            Width           =   6900
         End
         Begin VB.TextBox txtNomeAcompanhante 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   53
            Top             =   1605
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Porquê ?"
            Height          =   240
            Index           =   31
            Left            =   150
            TabIndex        =   50
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Acompanhante"
            Height          =   240
            Index           =   32
            Left            =   150
            TabIndex        =   52
            Top             =   1335
            Width           =   2250
         End
      End
      Begin VB.Frame fraEstado 
         Caption         =   " Estado Civil "
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
         TabIndex        =   37
         Top             =   1185
         Width           =   7170
         Begin VB.CheckBox chkConjInscr 
            Caption         =   "Conjugue Inscreveu-se"
            Height          =   240
            Left            =   4530
            TabIndex        =   40
            Top             =   570
            Width           =   2400
         End
         Begin VB.TextBox txtEstado_Civil 
            Height          =   360
            Left            =   135
            TabIndex        =   39
            Top             =   510
            Width           =   4200
         End
         Begin VB.TextBox txtConjugue 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   42
            Top             =   1305
            Width           =   6900
         End
         Begin SSDataWidgets_B.SSDBCombo cboEstado_Civil 
            Height          =   360
            Left            =   135
            TabIndex        =   57
            Top             =   510
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil"
            Height          =   240
            Index           =   11
            Left            =   135
            TabIndex        =   38
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Conjugue"
            Height          =   240
            Index           =   26
            Left            =   150
            TabIndex        =   41
            Top             =   1065
            Width           =   855
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
         Left            =   -74850
         TabIndex        =   28
         Top             =   2970
         Width           =   7170
         Begin VB.TextBox txtDistrito 
            Height          =   360
            Left            =   3660
            MaxLength       =   30
            TabIndex        =   36
            Top             =   1230
            Width           =   3390
         End
         Begin VB.TextBox txtLugar 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   30
            Top             =   570
            Width           =   3390
         End
         Begin VB.TextBox txtConcelho 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   34
            Top             =   1230
            Width           =   3390
         End
         Begin VB.TextBox txtFreguesia 
            Height          =   360
            Left            =   3660
            MaxLength       =   30
            TabIndex        =   32
            Top             =   570
            Width           =   3390
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   240
            Index           =   25
            Left            =   3660
            TabIndex        =   35
            Top             =   990
            Width           =   615
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Lugar"
            Height          =   240
            Index           =   24
            Left            =   150
            TabIndex        =   29
            Top             =   330
            Width           =   510
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Concellho"
            Height          =   240
            Index           =   23
            Left            =   150
            TabIndex        =   33
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Freguesia"
            Height          =   240
            Index           =   22
            Left            =   3660
            TabIndex        =   31
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Frame fraFiliacao 
         Caption         =   " Filiação"
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
         Top             =   3000
         Width           =   7170
         Begin VB.TextBox txtNome_Pai 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   45
            Top             =   540
            Width           =   6900
         End
         Begin VB.TextBox txtNome_Mae 
            Height          =   360
            Left            =   165
            MaxLength       =   50
            TabIndex        =   47
            Top             =   1185
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pai"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   44
            Top             =   300
            Width           =   300
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mãe"
            Height          =   240
            Index           =   9
            Left            =   165
            TabIndex        =   46
            Top             =   945
            Width           =   405
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
         TabIndex        =   21
         Top             =   1050
         Width           =   7170
         Begin VB.TextBox txtLocal_Emi_BI 
            Height          =   360
            Left            =   120
            MaxLength       =   40
            TabIndex        =   27
            Top             =   1200
            Width           =   6300
         End
         Begin VB.TextBox txtNum_BI 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   23
            Top             =   540
            Width           =   1500
         End
         Begin GTMaskDate.GTMaskDate dcboData_BI 
            Height          =   375
            Left            =   2100
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   22
            Top             =   300
            Width           =   225
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão"
            Height          =   240
            Index           =   10
            Left            =   2100
            TabIndex        =   24
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Local de Emissão"
            Height          =   240
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1620
         End
      End
      Begin VB.Frame fraInscricao 
         Caption         =   " Inscrição "
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
         Height          =   3720
         Left            =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   7170
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   20
            Top             =   3210
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   18
            Top             =   3210
            Width           =   1500
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   14
            Top             =   2550
            Width           =   1110
         End
         Begin VB.TextBox txtMorada 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   12
            Top             =   1890
            Width           =   6900
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   16
            Top             =   2550
            Width           =   5655
         End
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   3
            Top             =   570
            Width           =   6900
         End
         Begin VB.OptionButton optSexo 
            Height          =   300
            Index           =   0
            Left            =   2280
            TabIndex        =   7
            Top             =   1230
            Value           =   -1  'True
            Width           =   210
         End
         Begin VB.OptionButton optSexo 
            Height          =   300
            Index           =   1
            Left            =   3720
            TabIndex        =   9
            Top             =   1230
            Width           =   210
         End
         Begin GTMaskDate.GTMaskDate dcboData_Nasc 
            Height          =   360
            Left            =   150
            TabIndex        =   5
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
            Index           =   7
            Left            =   1800
            TabIndex        =   19
            Top             =   2970
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   17
            Top             =   2970
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   11
            Top             =   1650
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   13
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   4
            Left            =   1395
            TabIndex        =   15
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nascimento"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   990
            Width           =   1845
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            Height          =   240
            Index           =   17
            Left            =   2280
            TabIndex        =   6
            Top             =   990
            Width           =   465
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Masculino"
            Height          =   240
            Index           =   18
            Left            =   2550
            TabIndex        =   8
            Top             =   1230
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Feminino"
            Height          =   240
            Index           =   19
            Left            =   3990
            TabIndex        =   10
            Top             =   1230
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmCAIFFichaInscricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSInscricoes_Idosos As Workspace
Dim mBDInscricoes_Idosos As Database
Dim mBDInscricoes_IdososTemp As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lNum_Inscricao
Dim cNomeMapa

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSInscricoes_Idosos = DBEngine.CreateWorkspace(cBotao & "Inscricoes_Idosos", gUtilizador.Nome, gUtilizador.Password)
    Set mBDInscricoes_Idosos = mWSInscricoes_Idosos.OpenDatabase(cBD_Path & cNomeBD)
    If cBotao = "Ficha" Then
        Set mBDInscricoes_IdososTemp = mWSInscricoes_Idosos.OpenDatabase(cBDComNomeUtilizador)
    End If
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("CAIF - Ficha Inscrição-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function
Private Function lNovoNumInscricao() As Long
    ' define as variacveis
    Dim recNumInscricao As Recordset
    
    ' atribui o SQL
    cSql = "SELECT NUM_INSCRICAO FROM INSCRICOES_IDOSOS Order by NUM_INSCRICAO"
    ' abre a tabela
    Set recNumInscricao = mBDInscricoes_Idosos.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
    ' define qual o proximo número de inscrição
    With recNumInscricao
        If .EOF And .BOF Then
            lNovoNumInscricao = CLng(Year(Date) & "0001")
        Else
            .MoveLast
            If Mid$(recNumInscricao!NUM_INSCRICAO, 1, 4) = Trim$(Year(Date)) Then
                lNovoNumInscricao = recNumInscricao!NUM_INSCRICAO + 1
            Else
                lNovoNumInscricao = CLng(Year(Date) & "0001")
            End If
        End If
    End With
    ' fecha a tabela
    recNumInscricao.Close
    Set recNumInscricao = Nothing
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
        
        txtNum_BI.Locked = True
        txtNum_BI.TabStop = False
        txtNum_BI.BackColor = &H8000000F
        
        dcboData_BI.Locked = True
        dcboData_BI.CalDropDown = False
        dcboData_BI.TabStop = False
        
        txtLocal_Emi_BI.Locked = True
        txtLocal_Emi_BI.TabStop = False
        txtLocal_Emi_BI.BackColor = &H8000000F
        
        txtLugar.Locked = True
        txtLugar.TabStop = False
        txtLugar.BackColor = &H8000000F
        
        txtFreguesia.Locked = True
        txtFreguesia.TabStop = False
        txtFreguesia.BackColor = &H8000000F
        
        txtConcelho.Locked = True
        txtConcelho.TabStop = False
        txtConcelho.BackColor = &H8000000F
        
        txtDistrito.Locked = True
        txtDistrito.TabStop = False
        txtDistrito.BackColor = &H8000000F
        
        cboEstado_Civil.Visible = False
        txtEstado_Civil.Visible = True
        txtEstado_Civil.Locked = True
        txtEstado_Civil.TabStop = False
        txtEstado_Civil.BackColor = &H8000000F
        
        chkConjInscr.Enabled = False
        chkConjInscr.TabStop = False
        chkConjInscr.BackColor = &H8000000F
        
        txtConjugue.Locked = True
        txtConjugue.TabStop = False
        txtConjugue.BackColor = &H8000000F
        
        txtNome_Pai.Locked = True
        txtNome_Pai.TabStop = False
        txtNome_Pai.BackColor = &H8000000F
        
        txtNome_Mae.Locked = True
        txtNome_Mae.TabStop = False
        txtNome_Mae.BackColor = &H8000000F
        
        chkViveSo.Enabled = False
        chkViveSo.TabStop = False
        chkViveSo.BackColor = &H8000000F
       
        txtPorque.Locked = True
        txtPorque.TabStop = False
        txtPorque.BackColor = &H8000000F
        
        txtNomeAcompanhante.Locked = True
        txtNomeAcompanhante.TabStop = False
        txtNomeAcompanhante.BackColor = &H8000000F
        
        txtAgregado.Locked = True
        txtAgregado.TabStop = False
        
        txtNomeResp.Locked = True
        txtNomeResp.TabStop = False
        txtNomeResp.BackColor = &H8000000F
       
        txtParentesco.Locked = True
        txtParentesco.TabStop = False
        txtParentesco.BackColor = &H8000000F
       
        txtLocalTrabResp.Locked = True
        txtLocalTrabResp.TabStop = False
        txtLocalTrabResp.BackColor = &H8000000F
       
        txtResidenciaResp.Locked = True
        txtResidenciaResp.TabStop = False
        txtResidenciaResp.BackColor = &H8000000F
        
        txtTelTrabResp.Locked = True
        txtTelTrabResp.TabStop = False
        txtTelTrabResp.BackColor = &H8000000F
        
        txtTelResidResp.Locked = True
        txtTelResidResp.TabStop = False
        txtTelResidResp.BackColor = &H8000000F
        
        txtObservacoes.Locked = True
        txtObservacoes.TabStop = False
        txtObservacoes.BackColor = &H8000000F
        
        txtReservado.Locked = True
        txtReservado.TabStop = False
        txtReservado.BackColor = &H8000000F
        
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
        cboEstado_Civil.Visible = True
        txtEstado_Civil.Visible = False
        
        cboInstituicao.Visible = True
        txtInstituicao.Visible = False
        
        cboSalas.Visible = True
        txtSalas.Visible = False
    End If
End Sub

Public Sub CamposLimpaCarrega()
    Dim recInscricao As Recordset
    ' Ficha
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM INSCRICOES_IDOSOS WHERE NUM_INSCRICAO=" & lNum_Inscricao
        Set recInscricao = mBDInscricoes_Idosos.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
        'Poe os campos com os dados da Inscricao
        txtNome.Text = vFiltraCamposNulos(recInscricao!Nome)
        dcboData_Nasc.Text = vFiltraCamposNulos(recInscricao!DATA_NASC)
        If vFiltraCamposNulos(recInscricao!SEXO) = "M" Then
            optSexo(0).Value = True
        ElseIf vFiltraCamposNulos(recInscricao!SEXO) = "F" Then
            optSexo(1).Value = True
        Else
            optSexo(0).Value = False
            optSexo(1).Value = False
        End If
        txtMorada.Text = vFiltraCamposNulos(recInscricao!MORADA)
        txtCodigoPostal.Text = vFiltraCamposNulos(recInscricao!COD_POSTAL)
        txtLocal.Text = vFiltraCamposNulos(recInscricao!LOCAL)
        txtTelefone.Text = vFiltraCamposNulos(recInscricao!TELEFONE1)
        txtTelemovel.Text = vFiltraCamposNulos(recInscricao!TELEFONE2)
        txtNum_BI.Text = vFiltraCamposNulos(recInscricao!NUM_BI)
        dcboData_BI.Text = vFiltraCamposNulos(recInscricao!DATA_BI)
        txtLocal_Emi_BI.Text = vFiltraCamposNulos(recInscricao!LOCAL_EMI_BI)
        txtLugar.Text = vFiltraCamposNulos(recInscricao!NAT_LUGAR)
        txtFreguesia.Text = vFiltraCamposNulos(recInscricao!NAT_FREG)
        txtConcelho.Text = vFiltraCamposNulos(recInscricao!NAT_CONC)
        txtDistrito.Text = vFiltraCamposNulos(recInscricao!NAT_DISTR)
        If cBotao = "Ficha" Then
            txtEstado_Civil.Text = cDescodificaEstadoCivil(recInscricao!COD_ESTADO_CIVIL)
        ElseIf cBotao = "Altera" Then
            cboEstado_Civil.Text = cDescodificaEstadoCivil(recInscricao!COD_ESTADO_CIVIL)
        End If
        chkConjInscr.Value = IIf(vFiltraCamposNulos(recInscricao!CONJ_INSCR), vbChecked, vbUnchecked)
        txtConjugue.Text = vFiltraCamposNulos(recInscricao!NOME_CONJUGE)
        txtNome_Pai.Text = vFiltraCamposNulos(recInscricao!NOME_PAI)
        txtNome_Mae.Text = vFiltraCamposNulos(recInscricao!NOME_MAE)
        chkViveSo.Value = IIf(vFiltraCamposNulos(recInscricao!VIVE_SO), vbChecked, vbUnchecked)
        txtPorque.Text = vFiltraCamposNulos(recInscricao!VIVE_SO_PORQUE)
        txtNomeAcompanhante.Text = vFiltraCamposNulos(recInscricao!NOME_ACOMP)
        txtAgregado.Text = vFiltraCamposNulos(recInscricao!AGREGADO)
        txtNomeResp.Text = vFiltraCamposNulos(recInscricao!NOME_RESP)
        txtParentesco.Text = vFiltraCamposNulos(recInscricao!PARENTE_RESP)
        txtLocalTrabResp.Text = vFiltraCamposNulos(recInscricao!MORADA_TRAB_RESP)
        txtResidenciaResp.Text = vFiltraCamposNulos(recInscricao!MORADA_RESID_RESP)
        txtTelTrabResp.Text = vFiltraCamposNulos(recInscricao!TELEFONE_TRAB_RESP)
        txtTelResidResp.Text = vFiltraCamposNulos(recInscricao!TELEFONE_RESID_RESP)
        txtObservacoes.Text = vFiltraCamposNulos(recInscricao!OBSERVACOES)
        txtReservado.Text = vFiltraCamposNulos(recInscricao!RESERVADO)
        If cBotao = "Ficha" Then
            txtInstituicao.Text = cDescodificaInstituicao(recInscricao!COD_INST)
            txtSalas.Text = cDescodificaSala(recInscricao!COD_INST, recInscricao!COD_SALA)
        ElseIf cBotao = "Altera" Then
            cboInstituicao.Text = cDescodificaInstituicao(recInscricao!COD_INST)
            cboSalas.Text = cDescodificaSala(recInscricao!COD_INST, recInscricao!COD_SALA)
        End If
      
        ' Novos Campos Rendimentos / Despesas
        txtR_TD.Text = vFiltraCamposNulos(recInscricao!R_TD)
        txtR_P.Text = vFiltraCamposNulos(recInscricao!R_P)
        txtR_PA.Text = vFiltraCamposNulos(recInscricao!R_PA)
        txtR_TI.Text = vFiltraCamposNulos(recInscricao!R_TI)
        txtR_R.Text = vFiltraCamposNulos(recInscricao!R_R)
        txtR_RCI.Text = vFiltraCamposNulos(recInscricao!R_RCI)
        txtR_SD.Text = vFiltraCamposNulos(recInscricao!R_SD)
        txtR_AF.Text = vFiltraCamposNulos(recInscricao!R_AF)
        txtR_O_DESC.Text = vFiltraCamposNulos(recInscricao!R_O_DESC)
        txtR_O.Text = vFiltraCamposNulos(recInscricao!R_O)
        txtD_IRS.Text = vFiltraCamposNulos(recInscricao!D_IRS)
        txtD_SS.Text = vFiltraCamposNulos(recInscricao!D_SS)
        txtD_IMI.Text = vFiltraCamposNulos(recInscricao!D_IMI)
        txtD_JEH.Text = vFiltraCamposNulos(recInscricao!D_JEH)
        
        ' fecha o recordset
        recInscricao.Close
        Set recInscricao = Nothing
    ElseIf cBotao = "Novo" Then
        'Poe os campos preparados para nova ficha
        txtNome.Text = vbNullString
        dcboData_Nasc.Text = vbNullString
        txtMorada.Text = vbNullString
        txtCodigoPostal.Text = vbNullString
        txtLocal.Text = vbNullString
        txtTelefone.Text = vbNullString
        txtTelemovel.Text = vbNullString
        txtNum_BI.Text = vbNullString
        dcboData_BI.Text = vbNullString
        txtLocal_Emi_BI.Text = vbNullString
        txtLugar.Text = vbNullString
        txtFreguesia.Text = vbNullString
        txtConcelho.Text = vbNullString
        txtDistrito.Text = vbNullString
        cboEstado_Civil.Text = vbNullString
        chkConjInscr.Value = vbUnchecked
        txtConjugue.Text = vbNullString
        txtNome_Pai.Text = vbNullString
        txtNome_Mae.Text = vbNullString
        chkViveSo.Value = vbUnchecked
        txtPorque.Text = vbNullString
        txtNomeAcompanhante.Text = vbNullString
        txtAgregado.Text = 0
        txtNomeResp.Text = vbNullString
        txtParentesco.Text = vbNullString
        txtLocalTrabResp.Text = vbNullString
        txtResidenciaResp.Text = vbNullString
        txtTelTrabResp.Text = vbNullString
        txtTelResidResp.Text = vbNullString
        txtObservacoes.Text = vbNullString
        txtReservado.Text = vbNullString
        cboInstituicao.Text = vbNullString
        cboSalas.Text = vbNullString
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

    cNomeMapa = "CBESQ028.RPT"
On Error GoTo TrataErro
    ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDInscricoes_IdososTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDInscricoes_Idosos.Execute cSql, dbFailOnError
    
    ' apaga os registos da Temp32
    cSql = "DELETE * FROM FICHA_INSCRICOES_IDOSOS"
    ' apaga o registo em Temp32
    mBDInscricoes_IdososTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_INSCRICOES_IDOSOS IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM INSCRICOES_IDOSOS WHERE NUM_INSCRICAO=" & lNum_Inscricao
    ' insere o registo em Temp32
    mBDInscricoes_Idosos.Execute cSql
                        
    With rptFichaIndividual
        'Carrega o Nome do Report se ele existir
        If tFicheiroExiste(cApl_Path & "\MAPAS\" & cNomeMapa) Then
            .ReportFileName = cApl_Path & "\MAPAS\" & cNomeMapa
        Else
            MsgBox "Não foi encontrado o Mapa!", vbInformation, "Impressão da Ficha de Inscrição"
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
        .DataFiles(1) = cBDComNomeUtilizador
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
    Call ErrosGerais("CAIF - Impressão Ficha de Inscrição", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
End Sub

Private Sub cmdOK_Click()
    Dim mProcessamento As Processamento
    Dim qryInscricao As QueryDef
    
    Set mProcessamento = New Processamento
    
    ' Campos Obrigatórios
    If Trim$(txtNome.Text) = vbNullString Then
        MsgBox "Nome é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Inscrição"
        txtNome.SetFocus
        Exit Sub
    End If
    If Trim$(cboInstituicao.Text) = vbNullString Then
        MsgBox "Instituição é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Inscrição"
        tabFichaInscricao.Tab = 6
        cboInstituicao.SetFocus
        Exit Sub
    End If
    If Trim$(cboSalas.Text) = vbNullString Then
        MsgBox "Valência é um campo Obrigatório !!!", vbInformation + vbOKOnly, "Ficha de Inscrição"
        tabFichaInscricao.Tab = 6
        cboSalas.SetFocus
        Exit Sub
    End If
    
On Error GoTo TrataErro
    ' começa a transação
    mWSInscricoes_Idosos.BeginTrans
    If cBotao = "Novo" Then
        Set qryInscricao = mBDInscricoes_Idosos.QueryDefs("CAIF INSCRICOES Insere")
        ' parametros de input
        qryInscricao.Parameters("Num_Inscricao") = lNovoNumInscricao
        qryInscricao.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryInscricao.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)
        qryInscricao.Parameters("Nome") = txtNome.Text
        qryInscricao.Parameters("Morada") = txtMorada.Text
        qryInscricao.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryInscricao.Parameters("Local") = txtLocal.Text
        qryInscricao.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryInscricao.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryInscricao.Parameters("Sexo") = "F"
        End If
        qryInscricao.Parameters("Cod_Estado_Civil") = cCodificaEstadoCivil(cboEstado_Civil.Text)
        qryInscricao.Parameters("Num_BI") = txtNum_BI.Text
        qryInscricao.Parameters("Data_BI") = dcboData_BI.DateValue
        qryInscricao.Parameters("Local_Emi_BI") = txtLocal_Emi_BI.Text
        qryInscricao.Parameters("Nat_Lugar") = txtLugar.Text
        qryInscricao.Parameters("Nat_Freg") = txtFreguesia.Text
        qryInscricao.Parameters("Nat_Conc") = txtConcelho.Text
        qryInscricao.Parameters("Nat_Distr") = txtDistrito.Text
        qryInscricao.Parameters("Nome_Conjuge") = txtConjugue.Text
        qryInscricao.Parameters("Conj_Inscr") = chkConjInscr.Value
        qryInscricao.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryInscricao.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryInscricao.Parameters("Telefone1") = txtTelefone.Text
        qryInscricao.Parameters("Telefone2") = txtTelemovel.Text
        qryInscricao.Parameters("Vive_So") = chkViveSo.Value
        qryInscricao.Parameters("Vive_So_Porque") = txtPorque.Text
        qryInscricao.Parameters("Nome_Acomp") = txtNomeAcompanhante.Text
        qryInscricao.Parameters("Agregado") = txtAgregado.Value
        qryInscricao.Parameters("Nome_Resp") = txtNomeResp.Text
        qryInscricao.Parameters("Parente_Resp") = txtParentesco.Text
        qryInscricao.Parameters("Morada_Trab_Resp") = txtLocalTrabResp.Text
        qryInscricao.Parameters("Telefone_Trab_Resp") = txtTelTrabResp.Text
        qryInscricao.Parameters("Morada_Resid_Resp") = txtResidenciaResp.Text
        qryInscricao.Parameters("Telefone_Resid_Resp") = txtTelResidResp.Text
        qryInscricao.Parameters("Observacoes") = txtObservacoes.Text
        qryInscricao.Parameters("Reservado") = txtReservado.Text
        qryInscricao.Parameters("Utilizador") = gUtilizador.Nome
        ' Novos Campos Rendimentos / Despesas
        qryInscricao.Parameters("R_TD") = txtR_TD.Value
        qryInscricao.Parameters("R_P") = txtR_P.Value
        qryInscricao.Parameters("R_PA") = txtR_PA.Value
        qryInscricao.Parameters("R_TI") = txtR_TI.Value
        qryInscricao.Parameters("R_R") = txtR_R.Value
        qryInscricao.Parameters("R_RCI") = txtR_RCI.Value
        qryInscricao.Parameters("R_SD") = txtR_SD.Value
        qryInscricao.Parameters("R_AF") = txtR_AF.Value
        qryInscricao.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryInscricao.Parameters("R_O") = txtR_O.Value
        qryInscricao.Parameters("D_IRS") = txtD_IRS.Value
        qryInscricao.Parameters("D_SS") = txtD_SS.Value
        qryInscricao.Parameters("D_IMI") = txtD_IMI.Value
        qryInscricao.Parameters("D_JEH") = txtD_JEH.Value
       
    ElseIf cBotao = "Altera" Then
        Set qryInscricao = mBDInscricoes_Idosos.QueryDefs("CAIF INSCRICOES Altera")
        ' parametros de input
        qryInscricao.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryInscricao.Parameters("Cod_Sala") = cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text)
        qryInscricao.Parameters("Nome") = txtNome.Text
        qryInscricao.Parameters("Morada") = txtMorada.Text
        qryInscricao.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryInscricao.Parameters("Local") = txtLocal.Text
        qryInscricao.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryInscricao.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryInscricao.Parameters("Sexo") = "F"
        End If
        qryInscricao.Parameters("Cod_Estado_Civil") = cCodificaEstadoCivil(cboEstado_Civil.Text)
        qryInscricao.Parameters("Num_BI") = txtNum_BI.Text
        qryInscricao.Parameters("Data_BI") = dcboData_BI.DateValue
        qryInscricao.Parameters("Local_Emi_BI") = txtLocal_Emi_BI.Text
        qryInscricao.Parameters("Nat_Lugar") = txtLugar.Text
        qryInscricao.Parameters("Nat_Freg") = txtFreguesia.Text
        qryInscricao.Parameters("Nat_Conc") = txtConcelho.Text
        qryInscricao.Parameters("Nat_Distr") = txtDistrito.Text
        qryInscricao.Parameters("Nome_Conjuge") = txtConjugue.Text
        qryInscricao.Parameters("Conj_Inscr") = chkConjInscr.Value
        qryInscricao.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryInscricao.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryInscricao.Parameters("Telefone1") = txtTelefone.Text
        qryInscricao.Parameters("Telefone2") = txtTelemovel.Text
        qryInscricao.Parameters("Vive_So") = chkViveSo.Value
        qryInscricao.Parameters("Vive_So_Porque") = txtPorque.Text
        qryInscricao.Parameters("Nome_Acomp") = txtNomeAcompanhante.Text
        qryInscricao.Parameters("Agregado") = txtAgregado.Value
        qryInscricao.Parameters("Nome_Resp") = txtNomeResp.Text
        qryInscricao.Parameters("Parente_Resp") = txtParentesco.Text
        qryInscricao.Parameters("Morada_Trab_Resp") = txtLocalTrabResp.Text
        qryInscricao.Parameters("Telefone_Trab_Resp") = txtTelTrabResp.Text
        qryInscricao.Parameters("Morada_Resid_Resp") = txtResidenciaResp.Text
        qryInscricao.Parameters("Telefone_Resid_Resp") = txtTelResidResp.Text
        qryInscricao.Parameters("Observacoes") = txtObservacoes.Text
        qryInscricao.Parameters("Reservado") = txtReservado.Text
        qryInscricao.Parameters("Utilizador") = gUtilizador.Nome
        qryInscricao.Parameters("Num_Inscricao") = lNum_Inscricao
        ' Novos Campos Rendimentos / Despesas
        qryInscricao.Parameters("R_TD") = txtR_TD.Value
        qryInscricao.Parameters("R_P") = txtR_P.Value
        qryInscricao.Parameters("R_PA") = txtR_PA.Value
        qryInscricao.Parameters("R_TI") = txtR_TI.Value
        qryInscricao.Parameters("R_R") = txtR_R.Value
        qryInscricao.Parameters("R_RCI") = txtR_RCI.Value
        qryInscricao.Parameters("R_SD") = txtR_SD.Value
        qryInscricao.Parameters("R_AF") = txtR_AF.Value
        qryInscricao.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryInscricao.Parameters("R_O") = txtR_O.Value
        qryInscricao.Parameters("D_IRS") = txtD_IRS.Value
        qryInscricao.Parameters("D_SS") = txtD_SS.Value
        qryInscricao.Parameters("D_IMI") = txtD_IMI.Value
        qryInscricao.Parameters("D_JEH") = txtD_JEH.Value
        
    End If
    
    ' executa a query
    qryInscricao.Execute dbFailOnError
    
    mWSInscricoes_Idosos.CommitTrans
    ' faz o refresh da frmGestaoInscricoes
    frmCAIFGestaoInscricoes.datInscricoes.Refresh
    frmCAIFGestaoInscricoes.sgrdGestaoInscricoes.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSInscricoes_Idosos.Rollback
    Call ErrosGerais(" CAIF - " & cBotao & " Inscrição", Err.Number, Err.Description)
    Resume SairDoProcedimento
    
SairDoProcedimento:
    Unload Me
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
        lNum_Inscricao = frmCAIFGestaoInscricoes.sgrdGestaoInscricoes.Columns(0).Value
    End If

    If cBotao = "Ficha" Then
        Me.Caption = Me.Caption & " Nº " & lNum_Inscricao
        cmdOK.Visible = False
        cmdImprimir.Visible = True
    ElseIf cBotao = "Novo" Then
        Me.Caption = "Nova " & Me.Caption
        cmdOK.Visible = True
        cmdImprimir.Visible = False
    ElseIf cBotao = "Altera" Then
        Me.Caption = "Alteração na " & Me.Caption & " Nº " & lNum_Inscricao
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
    For Each mBD In mWSInscricoes_Idosos.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSInscricoes_Idosos = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub




