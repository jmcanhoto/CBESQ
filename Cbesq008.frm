VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFichaInscricao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Inscrição"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
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
   ScaleHeight     =   6795
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   5820
      Top             =   5955
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   6585
      Style           =   1  'Graphical
      TabIndex        =   83
      Tag             =   "2010"
      ToolTipText     =   "Imprimir a Ficha de Inscrição"
      Top             =   5775
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   85
      Tag             =   "2003"
      ToolTipText     =   "Sair da Visualização da Ficha de Inscrição"
      Top             =   5775
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6585
      Style           =   1  'Graphical
      TabIndex        =   84
      Tag             =   "2002"
      ToolTipText     =   "Inserir / Alterar a Ficha de Inscrição"
      Top             =   5775
      Width           =   1200
   End
   Begin TabDlg.SSTab tabFichaInscricao 
      Height          =   5580
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9843
      _Version        =   393216
      Style           =   1
      Tabs            =   6
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
      TabPicture(0)   =   "Cbesq008.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Pais"
      TabPicture(1)   =   "Cbesq008.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPai"
      Tab(1).Control(1)=   "fraMae"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3 - Outros Dados"
      TabPicture(2)   =   "Cbesq008.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOutrosDados"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Observações"
      TabPicture(3)   =   "Cbesq008.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraObservacoes"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5 - Rendimentos"
      TabPicture(4)   =   "Cbesq008.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraRendimentos"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Despesas"
      TabPicture(5)   =   "Cbesq008.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3930
         Left            =   -74880
         TabIndex        =   70
         Top             =   840
         Width           =   8790
         Begin VB.TextBox txtD_O_DESC 
            Height          =   360
            Left            =   180
            MaxLength       =   50
            TabIndex        =   90
            Top             =   3105
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtD_IRS 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   72
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
            Index           =   0
            Left            =   5430
            TabIndex        =   75
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
            Index           =   0
            Left            =   5430
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
         Begin GTMaskNum.GTMaskNum txtD_JEH 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   81
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
         Begin GTMaskNum.GTMaskNum txtD_IRS 
            Height          =   360
            Index           =   1
            Left            =   7050
            TabIndex        =   73
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
            Index           =   1
            Left            =   7050
            TabIndex        =   76
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
            Index           =   1
            Left            =   7050
            TabIndex        =   79
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
            Index           =   1
            Left            =   7050
            TabIndex        =   82
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
         Begin GTMaskNum.GTMaskNum txtD_SC 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   91
            Top             =   2115
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
         Begin GTMaskNum.GTMaskNum txtD_SC 
            Height          =   360
            Index           =   1
            Left            =   7035
            TabIndex        =   92
            Top             =   2115
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
         Begin GTMaskNum.GTMaskNum txtD_T 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   93
            Top             =   2595
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
         Begin GTMaskNum.GTMaskNum txtD_T 
            Height          =   360
            Index           =   1
            Left            =   7035
            TabIndex        =   94
            Top             =   2595
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
         Begin GTMaskNum.GTMaskNum txtD_O 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   95
            Top             =   3090
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
         Begin GTMaskNum.GTMaskNum txtD_O 
            Height          =   360
            Index           =   1
            Left            =   7035
            TabIndex        =   96
            Top             =   3075
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
            Caption         =   "Saúde / Crónica (Anual)................................................................."
            Height          =   240
            Index           =   11
            Left            =   180
            TabIndex        =   98
            Top             =   2145
            Width           =   5040
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Transportes (Anual) ........................................................................"
            Height          =   240
            Index           =   31
            Left            =   180
            TabIndex        =   97
            Top             =   2625
            Width           =   5055
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Mãe"
            Height          =   240
            Index           =   3
            Left            =   7065
            TabIndex        =   89
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Pai"
            Height          =   240
            Index           =   2
            Left            =   5490
            TabIndex        =   88
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I R S (Anual) ........................................................................................."
            Height          =   240
            Index           =   36
            Left            =   150
            TabIndex        =   71
            Top             =   300
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Segurança Social (Anual) .............................................................."
            Height          =   240
            Index           =   35
            Left            =   150
            TabIndex        =   74
            Top             =   750
            Width           =   5115
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I M I (Anual) ........................................................................................."
            Height          =   240
            Index           =   34
            Left            =   150
            TabIndex        =   77
            Top             =   1230
            Width           =   5070
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Juros / Amortização Empréstimo Habitação (Anual) ........."
            Height          =   240
            Index           =   33
            Left            =   150
            TabIndex        =   80
            Top             =   1710
            Width           =   5010
         End
      End
      Begin VB.Frame fraObservacoes 
         Height          =   3720
         Left            =   -74850
         TabIndex        =   34
         Top             =   900
         Width           =   8805
         Begin VB.TextBox txtInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   41
            Top             =   3180
            Width           =   5250
         End
         Begin VB.TextBox txtObservacoes 
            Height          =   900
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   540
            Width           =   8400
         End
         Begin VB.TextBox txtReservado 
            Height          =   900
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1860
            Width           =   8400
         End
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   40
            Top             =   3180
            Width           =   5265
            _Version        =   196617
            DataMode        =   2
            Columns(0).Width=   3200
            _ExtentX        =   9287
            _ExtentY        =   635
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Instituição para onde se destina a Inscrição"
            Height          =   240
            Index           =   20
            Left            =   150
            TabIndex        =   39
            Top             =   2940
            Width           =   3855
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   240
            Index           =   14
            Left            =   150
            TabIndex        =   35
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Reservado à Institução"
            Height          =   240
            Index           =   15
            Left            =   150
            TabIndex        =   37
            Top             =   1620
            Width           =   2070
         End
      End
      Begin VB.Frame fraOutrosDados 
         Height          =   3720
         Left            =   -74850
         TabIndex        =   29
         Top             =   900
         Width           =   8805
         Begin VB.CommandButton cmdCalcularRP 
            Caption         =   "Calcular"
            Height          =   300
            Left            =   1905
            TabIndex        =   101
            Top             =   2460
            Width           =   1320
         End
         Begin VB.TextBox txtOndeFica 
            Height          =   900
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   555
            Width           =   8400
         End
         Begin GTMaskNum.GTMaskNum txtAgregado 
            Height          =   360
            Left            =   150
            TabIndex        =   33
            Top             =   1740
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
         Begin GTMaskNum.GTMaskNum txtRP 
            Height          =   360
            Left            =   150
            TabIndex        =   99
            Top             =   2445
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
            Locked          =   -1  'True
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
            Caption         =   "Rend. PerCapita"
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   100
            Top             =   2205
            Width           =   1485
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Agregado"
            Height          =   240
            Index           =   13
            Left            =   150
            TabIndex        =   32
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Onde fica presentemente"
            Height          =   240
            Index           =   16
            Left            =   150
            TabIndex        =   30
            Top             =   315
            Width           =   2250
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
         Top             =   900
         Width           =   8805
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
            Width           =   8400
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   16
            Top             =   2550
            Width           =   7200
         End
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   3
            Top             =   570
            Width           =   8400
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
            Top             =   1200
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
      Begin VB.Frame fraMae 
         Caption         =   " Mãe "
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
         TabIndex        =   24
         Top             =   2820
         Width           =   8805
         Begin VB.TextBox txtTelefone_Emp_Mae 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   28
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Mae 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   26
            Top             =   540
            Width           =   8400
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   25
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   240
            Index           =   10
            Left            =   150
            TabIndex        =   27
            Top             =   960
            Width           =   810
         End
      End
      Begin VB.Frame fraPai 
         Caption         =   " Pai "
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
         TabIndex        =   19
         Top             =   900
         Width           =   8805
         Begin VB.TextBox txtTelefone_Emp_Pai 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   23
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Pai 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   21
            Top             =   540
            Width           =   8400
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   20
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   7
            Left            =   150
            TabIndex        =   22
            Top             =   960
            Width           =   1320
         End
      End
      Begin VB.Frame fraRendimentos 
         Height          =   4560
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   8820
         Begin VB.TextBox txtR_O_DESC 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   67
            Top             =   4080
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtR_TD 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   44
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
            Index           =   0
            Left            =   5430
            TabIndex        =   47
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
            Index           =   0
            Left            =   5400
            TabIndex        =   50
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
            Index           =   0
            Left            =   5400
            TabIndex        =   53
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
            Index           =   0
            Left            =   5400
            TabIndex        =   56
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
         Begin GTMaskNum.GTMaskNum txtR_RSI 
            Height          =   360
            Index           =   0
            Left            =   5400
            TabIndex        =   59
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
            Index           =   0
            Left            =   5400
            TabIndex        =   62
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
            Index           =   0
            Left            =   5400
            TabIndex        =   65
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
            Index           =   0
            Left            =   5400
            TabIndex        =   68
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
         Begin GTMaskNum.GTMaskNum txtR_TD 
            Height          =   360
            Index           =   1
            Left            =   7050
            TabIndex        =   45
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
            Index           =   1
            Left            =   7050
            TabIndex        =   48
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
            Index           =   1
            Left            =   7020
            TabIndex        =   51
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
            Index           =   1
            Left            =   7020
            TabIndex        =   54
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
            Index           =   1
            Left            =   7020
            TabIndex        =   57
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
         Begin GTMaskNum.GTMaskNum txtR_RSI 
            Height          =   360
            Index           =   1
            Left            =   7020
            TabIndex        =   60
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
            Index           =   1
            Left            =   7020
            TabIndex        =   63
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
            Index           =   1
            Left            =   7020
            TabIndex        =   66
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
            Index           =   1
            Left            =   7020
            TabIndex        =   69
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
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Mãe"
            Height          =   240
            Index           =   1
            Left            =   7050
            TabIndex        =   87
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Pai"
            Height          =   240
            Index           =   0
            Left            =   5460
            TabIndex        =   86
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Abono (Anual) ...................................................................................."
            Height          =   240
            Index           =   28
            Left            =   150
            TabIndex        =   64
            Top             =   3630
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Subsidio Desemprego (Anual) ....................................................."
            Height          =   240
            Index           =   27
            Left            =   150
            TabIndex        =   61
            Top             =   3150
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "R S I (Anual) ........................................................................................"
            Height          =   240
            Index           =   26
            Left            =   150
            TabIndex        =   58
            Top             =   2670
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Rendas (Anual) .................................................................................."
            Height          =   240
            Index           =   25
            Left            =   150
            TabIndex        =   55
            Top             =   2190
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Independente (Anual) ...................................................."
            Height          =   240
            Index           =   24
            Left            =   150
            TabIndex        =   52
            Top             =   1710
            Width           =   5175
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões Alimentos (Anual) ............................................................"
            Height          =   240
            Index           =   23
            Left            =   150
            TabIndex        =   49
            Top             =   1230
            Width           =   5160
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões (Anual) ................................................................................"
            Height          =   240
            Index           =   21
            Left            =   150
            TabIndex        =   46
            Top             =   750
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Dependente (Anual) ......................................................"
            Height          =   240
            Index           =   22
            Left            =   150
            TabIndex        =   43
            Top             =   300
            Width           =   5145
         End
      End
   End
End
Attribute VB_Name = "frmFichaInscricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mWSInscricao As Workspace
Dim mBDInscricao As Database
Dim mBDInscricaoTemp As Database
    
Dim tBDAberta

Dim cBotao
Dim cSql
Dim lNum_Inscricao
Dim cNomeMapa

Public Sub CalculaRP()
    Dim lValorPerCapita
    Dim lValorDespesasLimitadas
    
    ' Calcula Despesas Limitadas
' 20150515 introduzido o coampo IMI representa o ERPI
'    lValorDespesasLimitadas = (txtD_JEH(0).Value + txtD_SC(0).Value + txtD_T(0).Value + txtD_JEH(1).Value + txtD_SC(1).Value + txtD_T(1).Value)
    lValorDespesasLimitadas = (txtD_JEH(0).Value + txtD_SC(0).Value + txtD_T(0).Value + txtD_IMI(0).Value + _
                                txtD_JEH(1).Value + txtD_SC(1).Value + txtD_T(1).Value + txtD_IMI(1).Value)
    If lValorDespesasLimitadas > 6060 Then
        lValorDespesasLimitadas = 6060
    End If
    
    ' Calcula Valor Per-Capita segundo formula dada pelo CBESQ
'    lValorPerCapita = (((txtR_TD(0).Value + txtR_P(0).Value + txtR_PA(0).Value + txtR_TI(0).Value + txtR_R(0).Value + txtR_RSI(0).Value + txtR_SD(0).Value + txtR_AF(0).Value + txtR_O(0).Value) + _
'                        (txtR_TD(1).Value + txtR_P(1).Value + txtR_PA(1).Value + txtR_TI(1).Value + txtR_R(1).Value + txtR_RSI(1).Value + txtR_SD(1).Value + txtR_AF(1).Value + txtR_O(1).Value)) - _
'                        ((txtD_IRS(0).Value + txtD_SS(0).Value + txtD_IMI(0).Value + txtD_O(0).Value) + _
'                        (txtD_IRS(1).Value + txtD_SS(1).Value + txtD_IMI(1).Value + txtD_O(1).Value) + _
'                        lValorDespesasLimitadas)) / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
' 20150515 retirado o campo IMI representa ERPI
    lValorPerCapita = (((txtR_TD(0).Value + txtR_P(0).Value + txtR_PA(0).Value + txtR_TI(0).Value + txtR_R(0).Value + txtR_RSI(0).Value + txtR_SD(0).Value + txtR_AF(0).Value + txtR_O(0).Value) + _
                        (txtR_TD(1).Value + txtR_P(1).Value + txtR_PA(1).Value + txtR_TI(1).Value + txtR_R(1).Value + txtR_RSI(1).Value + txtR_SD(1).Value + txtR_AF(1).Value + txtR_O(1).Value)) - _
                        ((txtD_IRS(0).Value + txtD_SS(0).Value + txtD_O(0).Value) + _
                        (txtD_IRS(1).Value + txtD_SS(1).Value + txtD_O(1).Value) + _
                        lValorDespesasLimitadas)) / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
    
    txtRP.Value = Round(lValorPerCapita, 2)
End Sub

'Esta função vai abrir a Base de Dados
Private Function tAbreBD()
    Dim tRetorno
    
    tRetorno = True
    
    On Error GoTo TrataErro
    Set mWSInscricao = DBEngine.CreateWorkspace(cBotao & "Inscricao", gUtilizador.Nome, gUtilizador.Password)
    Set mBDInscricao = mWSInscricao.OpenDatabase(cBD_Path & cNomeBD)
    If cBotao = "Ficha" Then
        Set mBDInscricaoTemp = mWSInscricao.OpenDatabase(cBDComNomeUtilizador)
    End If
    
    GoTo SaiDaFuncao

TrataErro:
    tRetorno = False
    Call ErrosGerais("Ficha Inscrição-Abrir BD", Err.Number, Err.Description)
    Resume SaiDaFuncao

SaiDaFuncao:
    tAbreBD = tRetorno
End Function
Private Function lNovoNumInscricao() As Long
    ' define as variacveis
    Dim recNumInscricao As Recordset
    
    ' atribui o SQL
    cSql = "SELECT NUM_INSCRICAO FROM INSCRICOES Order by NUM_INSCRICAO"
    ' abre a tabela
    Set recNumInscricao = mBDInscricao.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
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
        
        txtNome_Pai.Locked = True
        txtNome_Pai.TabStop = False
        txtNome_Pai.BackColor = &H8000000F
        
        txtTelefone_Emp_Pai.Locked = True
        txtTelefone_Emp_Pai.TabStop = False
        txtTelefone_Emp_Pai.BackColor = &H8000000F
        
        txtNome_Mae.Locked = True
        txtNome_Mae.TabStop = False
        txtNome_Mae.BackColor = &H8000000F
        
        txtTelefone_Emp_Mae.Locked = True
        txtTelefone_Emp_Mae.TabStop = False
        txtTelefone_Emp_Mae.BackColor = &H8000000F
        
        txtAgregado.Locked = True
        txtAgregado.TabStop = False
        
        txtOndeFica.Locked = True
        txtOndeFica.TabStop = False
        txtOndeFica.BackColor = &H8000000F
        
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
        
        ' Novos Campos Rendimentos / Despesas
        txtR_TD(0).Locked = True
        txtR_TD(0).TabStop = False

        txtR_P(0).Locked = True
        txtR_P(0).TabStop = False

        txtR_PA(0).Locked = True
        txtR_PA(0).TabStop = False

        txtR_TI(0).Locked = True
        txtR_TI(0).TabStop = False

        txtR_R(0).Locked = True
        txtR_R(0).TabStop = False

        txtR_RSI(0).Locked = True
        txtR_RSI(0).TabStop = False

        txtR_SD(0).Locked = True
        txtR_SD(0).TabStop = False

        txtR_AF(0).Locked = True
        txtR_AF(0).TabStop = False

        txtR_O_DESC.Locked = True
        txtR_O_DESC.TabStop = False
        txtR_O_DESC.BackColor = &H8000000F
        
        txtR_O(0).Locked = True
        txtR_O(0).TabStop = False
        
        txtD_IRS(0).Locked = True
        txtD_IRS(0).TabStop = False
        
        txtD_SS(0).Locked = True
        txtD_SS(0).TabStop = False
        
        txtD_IMI(0).Locked = True
        txtD_IMI(0).TabStop = False
        
        txtD_JEH(0).Locked = True
        txtD_JEH(0).TabStop = False
        
        txtR_TD(1).Locked = True
        txtR_TD(1).TabStop = False

        txtR_P(1).Locked = True
        txtR_P(1).TabStop = False

        txtR_PA(1).Locked = True
        txtR_PA(1).TabStop = False

        txtR_TI(1).Locked = True
        txtR_TI(1).TabStop = False

        txtR_R(1).Locked = True
        txtR_R(1).TabStop = False

        txtR_RSI(1).Locked = True
        txtR_RSI(1).TabStop = False

        txtR_SD(1).Locked = True
        txtR_SD(1).TabStop = False

        txtR_AF(1).Locked = True
        txtR_AF(1).TabStop = False
        
        txtR_O(1).Locked = True
        txtR_O(1).TabStop = False
        
        txtD_IRS(1).Locked = True
        txtD_IRS(1).TabStop = False
        
        txtD_SS(1).Locked = True
        txtD_SS(1).TabStop = False
        
        txtD_IMI(1).Locked = True
        txtD_IMI(1).TabStop = False
        
        txtD_JEH(1).Locked = True
        txtD_JEH(1).TabStop = False
        
        txtD_SC(0).Locked = True
        txtD_SC(0).TabStop = False
    
        txtD_T(0).Locked = True
        txtD_T(0).TabStop = False
    
        txtD_O_DESC.Locked = True
        txtD_O_DESC.TabStop = False
        txtD_O_DESC.BackColor = &H8000000F
        
        txtD_O(0).Locked = True
        txtD_O(0).TabStop = False
    
        txtD_SC(1).Locked = True
        txtD_SC(1).TabStop = False
    
        txtD_T(1).Locked = True
        txtD_T(1).TabStop = False
    
        txtD_O(1).Locked = True
        txtD_O(1).TabStop = False
        
        txtRP.Locked = True
        txtRP.TabStop = False
        
    Else
        cboInstituicao.Visible = True
        txtInstituicao.Visible = False
    End If
End Sub

Public Sub CamposLimpaCarrega()
    Dim recInscricao As Recordset
    ' Ficha
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM INSCRICOES WHERE NUM_INSCRICAO=" & lNum_Inscricao
        Set recInscricao = mBDInscricao.OpenRecordset(cSql, dbOpenSnapshot, dbReadOnly)
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
        txtTelefone.Text = vFiltraCamposNulos(recInscricao!TELEFONE)
        txtNome_Pai.Text = vFiltraCamposNulos(recInscricao!NOME_PAI)
        txtTelefone_Emp_Pai.Text = vFiltraCamposNulos(recInscricao!TEL_EMP_PAI)
        txtNome_Mae.Text = vFiltraCamposNulos(recInscricao!NOME_MAE)
        txtTelefone_Emp_Mae.Text = vFiltraCamposNulos(recInscricao!TEL_EMP_MAE)
        txtAgregado.Text = vFiltraCamposNulos(recInscricao!AGREGADO)
        txtOndeFica.Text = vFiltraCamposNulos(recInscricao!ONDE_FICA)
        txtObservacoes.Text = vFiltraCamposNulos(recInscricao!OBSERVACOES)
        txtReservado.Text = vFiltraCamposNulos(recInscricao!RESERVADO)
        If cBotao = "Ficha" Then
            txtInstituicao.Text = cDescodificaInstituicao(recInscricao!COD_INST)
        ElseIf cBotao = "Altera" Then
            cboInstituicao.Text = cDescodificaInstituicao(recInscricao!COD_INST)
        End If
        ' Novos Campos Rendimentos / Despesas
        txtR_TD(0).Text = vFiltraCamposNulos(recInscricao!R_TD_1)
        txtR_P(0).Text = vFiltraCamposNulos(recInscricao!R_P_1)
        txtR_PA(0).Text = vFiltraCamposNulos(recInscricao!R_PA_1)
        txtR_TI(0).Text = vFiltraCamposNulos(recInscricao!R_TI_1)
        txtR_R(0).Text = vFiltraCamposNulos(recInscricao!R_R_1)
        txtR_RSI(0).Text = vFiltraCamposNulos(recInscricao!R_RSI_1)
        txtR_SD(0).Text = vFiltraCamposNulos(recInscricao!R_SD_1)
        txtR_AF(0).Text = vFiltraCamposNulos(recInscricao!R_AF_1)
        txtR_O_DESC.Text = vFiltraCamposNulos(recInscricao!R_O_DESC)
        txtR_O(0).Text = vFiltraCamposNulos(recInscricao!R_O_1)
        txtD_IRS(0).Text = vFiltraCamposNulos(recInscricao!D_IRS_1)
        txtD_SS(0).Text = vFiltraCamposNulos(recInscricao!D_SS_1)
        txtD_IMI(0).Text = vFiltraCamposNulos(recInscricao!D_IMI_1)
        txtD_JEH(0).Text = vFiltraCamposNulos(recInscricao!D_JEH_1)
        txtR_TD(1).Text = vFiltraCamposNulos(recInscricao!R_TD_2)
        txtR_P(1).Text = vFiltraCamposNulos(recInscricao!R_P_2)
        txtR_PA(1).Text = vFiltraCamposNulos(recInscricao!R_PA_2)
        txtR_TI(1).Text = vFiltraCamposNulos(recInscricao!R_TI_2)
        txtR_R(1).Text = vFiltraCamposNulos(recInscricao!R_R_2)
        txtR_RSI(1).Text = vFiltraCamposNulos(recInscricao!R_RSI_2)
        txtR_SD(1).Text = vFiltraCamposNulos(recInscricao!R_SD_2)
        txtR_AF(1).Text = vFiltraCamposNulos(recInscricao!R_AF_2)
        txtR_O(1).Text = vFiltraCamposNulos(recInscricao!R_O_2)
        txtD_IRS(1).Text = vFiltraCamposNulos(recInscricao!D_IRS_2)
        txtD_SS(1).Text = vFiltraCamposNulos(recInscricao!D_SS_2)
        txtD_IMI(1).Text = vFiltraCamposNulos(recInscricao!D_IMI_2)
        txtD_JEH(1).Text = vFiltraCamposNulos(recInscricao!D_JEH_2)
        
        txtD_SC(0).Text = vFiltraCamposNulos(recInscricao!D_SC_1)
        txtD_T(0).Text = vFiltraCamposNulos(recInscricao!D_T_1)
        txtD_O_DESC.Text = vFiltraCamposNulos(recInscricao!D_O_DESC)
        txtD_O(0).Text = vFiltraCamposNulos(recInscricao!D_O_1)
        txtD_SC(1).Text = vFiltraCamposNulos(recInscricao!D_SC_2)
        txtD_T(1).Text = vFiltraCamposNulos(recInscricao!D_T_2)
        txtD_O(1).Text = vFiltraCamposNulos(recInscricao!D_O_2)
        txtRP.Value = vFiltraCamposNulos(recInscricao!RP)
       
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
        txtNome_Pai.Text = vbNullString
        txtTelefone_Emp_Pai.Text = vbNullString
        txtNome_Mae.Text = vbNullString
        txtTelefone_Emp_Mae.Text = vbNullString
        txtAgregado.Text = 0
        txtOndeFica.Text = vbNullString
        txtObservacoes.Text = vbNullString
        txtReservado.Text = vbNullString
        ' Novos Campos Rendimentos / Despesas
        txtR_TD(0).Text = 0
        txtR_P(0).Text = 0
        txtR_PA(0).Text = 0
        txtR_TI(0).Text = 0
        txtR_R(0).Text = 0
        txtR_RSI(0).Text = 0
        txtR_SD(0).Text = 0
        txtR_AF(0).Text = 0
        txtR_O_DESC.Text = vbNullString
        txtR_O(0).Text = 0
        txtD_IRS(0).Text = 0
        txtD_SS(0).Text = 0
        txtD_IMI(0).Text = 0
        txtD_JEH(0).Text = 0
        txtR_TD(1).Text = 0
        txtR_P(1).Text = 0
        txtR_PA(1).Text = 0
        txtR_TI(1).Text = 0
        txtR_R(1).Text = 0
        txtR_RSI(1).Text = 0
        txtR_SD(1).Text = 0
        txtR_AF(1).Text = 0
        txtR_O(1).Text = 0
        txtD_IRS(1).Text = 0
        txtD_SS(1).Text = 0
        txtD_IMI(1).Text = 0
        txtD_JEH(1).Text = 0
        
        txtD_SC(0).Text = 0
        txtD_T(0).Text = 0
        txtD_O_DESC.Text = vbNullString
        txtD_O(0).Text = 0
        txtD_SC(1).Text = 0
        txtD_T(1).Text = 0
        txtD_O(1).Text = 0
        txtRP.Text = 0
        
    End If
End Sub
'Este procedimento activa ou desactiva o default do botao OK
Private Sub BotaoOKDefault(ByVal Propriedade As Boolean)
  cmdOK.Default = Propriedade
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
        .HeadStyleSet = "Cabecalho"
        .Font.Name = "MS Sans Serif"
        .Font.Size = 10
        .DataFieldList = "Column 0"
                            
        ' coluna 0
        .Columns.Add 0
        .Columns(0).Caption = "Nome da Institução"
        .Columns(0).CaptionAlignment = ssColCapAlignCenter
        .Columns(0).Width = .Width
    End With
End Sub

Private Sub cmdCalcularRP_Click()
    Call CalculaRP
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim mProcessamento As Processamento
    
    Set mProcessamento = New Processamento

    cNomeMapa = "CBESQ001.RPT"
On Error GoTo TrataErro
    ' apaga registos da TABINSTITUICAO
    cSql = "DELETE * FROM TABINSTITUICAO;"
    mBDInscricaoTemp.Execute cSql, dbFailOnError
    ' insere os registos em TABINSTITUICAO
    cSql = "INSERT INTO TABINSTITUICAO IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM TABINSTITUICAO"
    mBDInscricao.Execute cSql, dbFailOnError
    
    ' apaga os registos da Temp32
    cSql = "DELETE * FROM FICHA_INSCRICOES"
    ' apaga o registo em Temp32
    mBDInscricaoTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_INSCRICOES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM INSCRICOES WHERE NUM_INSCRICAO=" & lNum_Inscricao
    ' insere o registo em Temp32
    mBDInscricao.Execute cSql
                        
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
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
                      
        'executa o Report
        .Action = 1
    End With
    
    GoTo SairDoProcedimento
    
TrataErro:
    Call ErrosGerais("Impressão Ficha de Inscrição", Err.Number, Err.Description)
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
        tabFichaInscricao.Tab = 3
        cboInstituicao.SetFocus
        Exit Sub
    End If
On Error GoTo TrataErro
    ' começa a transação
    mWSInscricao.BeginTrans
    If cBotao = "Novo" Then
        Set qryInscricao = mBDInscricao.QueryDefs("INSCRICOES Insere")
        ' parametros de input
        qryInscricao.Parameters("Num_Inscricao") = lNovoNumInscricao
        qryInscricao.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryInscricao.Parameters("Nome") = txtNome.Text
        qryInscricao.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryInscricao.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryInscricao.Parameters("Sexo") = "F"
        End If
        qryInscricao.Parameters("Morada") = txtMorada.Text
        qryInscricao.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryInscricao.Parameters("Local") = txtLocal.Text
        qryInscricao.Parameters("Telefone") = txtTelefone.Text
        qryInscricao.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryInscricao.Parameters("Tel_Emp_Pai") = txtTelefone_Emp_Pai.Text
        qryInscricao.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryInscricao.Parameters("Tel_Emp_Mae") = txtTelefone_Emp_Mae.Text
        qryInscricao.Parameters("Agregado") = txtAgregado.Value
        qryInscricao.Parameters("Onde_Fica") = txtOndeFica.Text
        qryInscricao.Parameters("Observacoes") = txtObservacoes.Text
        qryInscricao.Parameters("Reservado") = txtReservado.Text
        qryInscricao.Parameters("Utiliz") = gUtilizador.Nome
        ' Novos Campos Rendimentos / Despesas
        qryInscricao.Parameters("R_TD_1") = txtR_TD(0).Value
        qryInscricao.Parameters("R_P_1") = txtR_P(0).Value
        qryInscricao.Parameters("R_PA_1") = txtR_PA(0).Value
        qryInscricao.Parameters("R_TI_1") = txtR_TI(0).Value
        qryInscricao.Parameters("R_R_1") = txtR_R(0).Value
        qryInscricao.Parameters("R_RSI_1") = txtR_RSI(0).Value
        qryInscricao.Parameters("R_SD_1") = txtR_SD(0).Value
        qryInscricao.Parameters("R_AF_1") = txtR_AF(0).Value
        qryInscricao.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryInscricao.Parameters("R_O_1") = txtR_O(0).Value
        qryInscricao.Parameters("D_IRS_1") = txtD_IRS(0).Value
        qryInscricao.Parameters("D_SS_1") = txtD_SS(0).Value
        qryInscricao.Parameters("D_IMI_1") = txtD_IMI(0).Value
        qryInscricao.Parameters("D_JEH_1") = txtD_JEH(0).Value
        qryInscricao.Parameters("R_TD_2") = txtR_TD(1).Value
        qryInscricao.Parameters("R_P_2") = txtR_P(1).Value
        qryInscricao.Parameters("R_PA_2") = txtR_PA(1).Value
        qryInscricao.Parameters("R_TI_2") = txtR_TI(1).Value
        qryInscricao.Parameters("R_R_2") = txtR_R(1).Value
        qryInscricao.Parameters("R_RSI_2") = txtR_RSI(1).Value
        qryInscricao.Parameters("R_SD_2") = txtR_SD(1).Value
        qryInscricao.Parameters("R_AF_2") = txtR_AF(1).Value
        qryInscricao.Parameters("R_O_2") = txtR_O(1).Value
        qryInscricao.Parameters("D_IRS_2") = txtD_IRS(1).Value
        qryInscricao.Parameters("D_SS_2") = txtD_SS(1).Value
        qryInscricao.Parameters("D_IMI_2") = txtD_IMI(1).Value
        qryInscricao.Parameters("D_JEH_2") = txtD_JEH(1).Value
        
        qryInscricao.Parameters("D_SC_1") = txtD_SC(0).Value
        qryInscricao.Parameters("D_T_1") = txtD_T(0).Value
        qryInscricao.Parameters("D_O_DESC") = txtD_O_DESC.Text
        qryInscricao.Parameters("D_O_1") = txtD_O(0).Value
        qryInscricao.Parameters("D_SC_2") = txtD_SC(1).Value
        qryInscricao.Parameters("D_T_2") = txtD_T(1).Value
        qryInscricao.Parameters("D_O_2") = txtD_O(1).Value
        qryInscricao.Parameters("RP") = txtRP.Value
        
    ElseIf cBotao = "Altera" Then
        Set qryInscricao = mBDInscricao.QueryDefs("INSCRICOES Altera")
        ' parametros de input
        qryInscricao.Parameters("Cod_Inst") = cCodificaInstituicao(cboInstituicao.Text)
        qryInscricao.Parameters("Nome") = txtNome.Text
        qryInscricao.Parameters("Data_Nasc") = dcboData_Nasc.DateValue
        If optSexo(0).Value Then
            qryInscricao.Parameters("Sexo") = "M"
        ElseIf optSexo(1).Value Then
            qryInscricao.Parameters("Sexo") = "F"
        End If
        qryInscricao.Parameters("Morada") = txtMorada.Text
        qryInscricao.Parameters("Cod_Postal") = txtCodigoPostal.Text
        qryInscricao.Parameters("Local") = txtLocal.Text
        qryInscricao.Parameters("Telefone") = txtTelefone.Text
        qryInscricao.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryInscricao.Parameters("Tel_Emp_Pai") = txtTelefone_Emp_Pai.Text
        qryInscricao.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryInscricao.Parameters("Tel_Emp_Mae") = txtTelefone_Emp_Mae.Text
        qryInscricao.Parameters("Agregado") = txtAgregado.Value
        qryInscricao.Parameters("Onde_Fica") = txtOndeFica.Text
        qryInscricao.Parameters("Observacoes") = txtObservacoes.Text
        qryInscricao.Parameters("Reservado") = txtReservado.Text
        qryInscricao.Parameters("Num_Inscricao") = lNum_Inscricao
        qryInscricao.Parameters("Utiliz") = gUtilizador.Nome
        ' Novos Campos Rendimentos / Despesas
        qryInscricao.Parameters("R_TD_1") = txtR_TD(0).Value
        qryInscricao.Parameters("R_P_1") = txtR_P(0).Value
        qryInscricao.Parameters("R_PA_1") = txtR_PA(0).Value
        qryInscricao.Parameters("R_TI_1") = txtR_TI(0).Value
        qryInscricao.Parameters("R_R_1") = txtR_R(0).Value
        qryInscricao.Parameters("R_RSI_1") = txtR_RSI(0).Value
        qryInscricao.Parameters("R_SD_1") = txtR_SD(0).Value
        qryInscricao.Parameters("R_AF_1") = txtR_AF(0).Value
        qryInscricao.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryInscricao.Parameters("R_O_1") = txtR_O(0).Value
        qryInscricao.Parameters("D_IRS_1") = txtD_IRS(0).Value
        qryInscricao.Parameters("D_SS_1") = txtD_SS(0).Value
        qryInscricao.Parameters("D_IMI_1") = txtD_IMI(0).Value
        qryInscricao.Parameters("D_JEH_1") = txtD_JEH(0).Value
        qryInscricao.Parameters("R_TD_2") = txtR_TD(1).Value
        qryInscricao.Parameters("R_P_2") = txtR_P(1).Value
        qryInscricao.Parameters("R_PA_2") = txtR_PA(1).Value
        qryInscricao.Parameters("R_TI_2") = txtR_TI(1).Value
        qryInscricao.Parameters("R_R_2") = txtR_R(1).Value
        qryInscricao.Parameters("R_RSI_2") = txtR_RSI(1).Value
        qryInscricao.Parameters("R_SD_2") = txtR_SD(1).Value
        qryInscricao.Parameters("R_AF_2") = txtR_AF(1).Value
        qryInscricao.Parameters("R_O_2") = txtR_O(1).Value
        qryInscricao.Parameters("D_IRS_2") = txtD_IRS(1).Value
        qryInscricao.Parameters("D_SS_2") = txtD_SS(1).Value
        qryInscricao.Parameters("D_IMI_2") = txtD_IMI(1).Value
        qryInscricao.Parameters("D_JEH_2") = txtD_JEH(1).Value
        
        qryInscricao.Parameters("D_SC_1") = txtD_SC(0).Value
        qryInscricao.Parameters("D_T_1") = txtD_T(0).Value
        qryInscricao.Parameters("D_O_DESC") = txtD_O_DESC.Text
        qryInscricao.Parameters("D_O_1") = txtD_O(0).Value
        qryInscricao.Parameters("D_SC_2") = txtD_SC(1).Value
        qryInscricao.Parameters("D_T_2") = txtD_T(1).Value
        qryInscricao.Parameters("D_O_2") = txtD_O(1).Value
        qryInscricao.Parameters("RP") = txtRP.Value
        
    End If
    
    ' executa a query
    qryInscricao.Execute dbFailOnError
    
    mWSInscricao.CommitTrans
    ' faz o refresh da frmGestaoInscricoes
    frmGestaoInscricoes.datInscricoes.Refresh
    frmGestaoInscricoes.sgrdGestaoInscricoes.Refresh
    
    GoTo SairDoProcedimento
    
TrataErro:
    mWSInscricao.Rollback
    Call ErrosGerais(cBotao & " Inscrição", Err.Number, Err.Description)
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
        lNum_Inscricao = frmGestaoInscricoes.sgrdGestaoInscricoes.Columns(0).Value
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
    
    Call CarregacboInstituicao(cboInstituicao)
    Call CamposEnabledDisabled
    Call CamposLimpaCarrega
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mBD As Database, mTab As Recordset
    For Each mBD In mWSInscricao.Databases
        For Each mTab In mBD.Recordsets
            mTab.Close
        Next
        mBD.Close
    Next
    Set mWSInscricao = Nothing
    fFrmMDIPrincipal.atbOpcoes.Tools("mnuJan").Menu.Tools.Remove Me.Caption
End Sub




Private Sub txtD_IMI_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_IRS_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_JEH_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_O_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_SC_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_SS_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtD_T_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_AF_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_O_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_P_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_PA_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_R_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_RSI_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_SD_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_TD_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


Private Sub txtR_TI_LostFocus(Index As Integer)
    Call CalculaRP
End Sub


