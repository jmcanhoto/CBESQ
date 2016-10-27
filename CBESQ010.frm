VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "gtnum32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmFichaUtente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Utente"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
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
   ScaleHeight     =   6885
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport rptFichaIndividual 
      Left            =   6165
      Top             =   6045
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   900
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   233
      Tag             =   "2010"
      Top             =   5865
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   900
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   235
      Tag             =   "2003"
      Top             =   5865
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   900
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   234
      Tag             =   "2002"
      Top             =   5865
      Width           =   1200
   End
   Begin TabDlg.SSTab tabFichaUtente 
      Height          =   5580
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   9843
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   5
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
      TabPicture(0)   =   "CBESQ010.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picFoto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraUtente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDatas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2 - Pais"
      TabPicture(1)   =   "CBESQ010.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPai"
      Tab(1).Control(1)=   "fraMae"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&3 - Enc. de Educação"
      TabPicture(2)   =   "CBESQ010.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraEnc_Edu"
      Tab(2).Control(1)=   "fraOutrosDados"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&4 - Outros Dados"
      TabPicture(3)   =   "CBESQ010.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraMensalidade"
      Tab(3).Control(1)=   "fraCedula"
      Tab(3).Control(2)=   "fraNaturalidade"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&5 - Observações"
      TabPicture(4)   =   "CBESQ010.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraInstituicao"
      Tab(4).Control(1)=   "fraObservacoes"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "&6 - Rendimentos"
      TabPicture(5)   =   "CBESQ010.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraRendimentos"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&7 - Despesas"
      TabPicture(6)   =   "CBESQ010.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&8 - Saúde"
      TabPicture(7)   =   "CBESQ010.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "SSTab1"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "&9 - Autorizações"
      TabPicture(8)   =   "CBESQ010.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "SSTab2"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "10 - CATL"
      TabPicture(9)   =   "CBESQ010.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame2"
      Tab(9).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   " Dados CATL (só para o CATL) "
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
         Height          =   3810
         Left            =   -74850
         TabIndex        =   247
         Top             =   900
         Width           =   6360
         Begin VB.TextBox txtAlmoco 
            Height          =   360
            Left            =   150
            MaxLength       =   20
            TabIndex        =   256
            Top             =   3210
            Width           =   3255
         End
         Begin VB.TextBox txtTelefoneEscola 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   254
            Top             =   2550
            Width           =   1500
         End
         Begin VB.TextBox txtHorarioEscola 
            Height          =   360
            Left            =   150
            MaxLength       =   20
            TabIndex        =   252
            Top             =   1890
            Width           =   3390
         End
         Begin VB.TextBox txtNomeProf 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   249
            Top             =   1230
            Width           =   6000
         End
         Begin VB.TextBox txtNomeEscola 
            Height          =   360
            Left            =   150
            MaxLength       =   100
            TabIndex        =   248
            Top             =   570
            Width           =   6000
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Almoça Escola ou CBESQ"
            Height          =   240
            Index           =   104
            Left            =   150
            TabIndex        =   257
            Top             =   2970
            Width           =   2370
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Escola"
            Height          =   240
            Index           =   103
            Left            =   150
            TabIndex        =   255
            Top             =   2310
            Width           =   1485
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Horário da Escola"
            Height          =   240
            Index           =   102
            Left            =   150
            TabIndex        =   253
            Top             =   1650
            Width           =   1635
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome Professor(a)"
            Height          =   240
            Index           =   101
            Left            =   150
            TabIndex        =   251
            Top             =   990
            Width           =   1710
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome Escola"
            Height          =   240
            Index           =   100
            Left            =   150
            TabIndex        =   250
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Frame fraOutrosDados 
         Caption         =   " Outros Dados "
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
         Height          =   2550
         Left            =   -74850
         TabIndex        =   52
         Top             =   2760
         Width           =   4755
         Begin VB.TextBox txtSNS 
            Height          =   360
            Left            =   2400
            MaxLength       =   20
            TabIndex        =   60
            Top             =   1890
            Width           =   1950
         End
         Begin VB.TextBox txtNum_Contribuinte 
            Height          =   360
            Left            =   2400
            MaxLength       =   12
            TabIndex        =   56
            Top             =   930
            Width           =   1950
         End
         Begin VB.TextBox txtSeguranca_Social 
            Height          =   360
            Left            =   2400
            MaxLength       =   12
            TabIndex        =   58
            Top             =   1410
            Width           =   1950
         End
         Begin GTMaskNum.GTMaskNum txtAgregado 
            Height          =   360
            Left            =   2400
            TabIndex        =   54
            Top             =   450
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nº  de S.N.S."
            Height          =   240
            Index           =   94
            Left            =   300
            TabIndex        =   59
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "N I F"
            Height          =   240
            Index           =   8
            Left            =   300
            TabIndex        =   55
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "N I S S"
            Height          =   240
            Index           =   12
            Left            =   300
            TabIndex        =   57
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Agregado Familiar "
            Height          =   240
            Index           =   13
            Left            =   300
            TabIndex        =   53
            Top             =   480
            Width           =   1725
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
         Height          =   2520
         Left            =   -74880
         TabIndex        =   73
         Top             =   2790
         Width           =   8820
         Begin VB.TextBox txtProxEscalao 
            BackColor       =   &H8000000F&
            Height          =   360
            Left            =   3900
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   244
            TabStop         =   0   'False
            Top             =   570
            Width           =   900
         End
         Begin VB.TextBox txtEscalao 
            BackColor       =   &H8000000F&
            Height          =   360
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   240
            TabStop         =   0   'False
            Top             =   570
            Width           =   900
         End
         Begin VB.OptionButton optTipo_Mensalidade 
            Caption         =   "Manual"
            Height          =   330
            Index           =   1
            Left            =   300
            TabIndex        =   75
            Top             =   720
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.OptionButton optTipo_Mensalidade 
            Caption         =   "Automático"
            Height          =   330
            Index           =   0
            Left            =   300
            TabIndex        =   74
            Top             =   360
            Width           =   1410
         End
         Begin GTMaskNum.GTMaskNum txtMensalidade 
            Height          =   360
            Left            =   5550
            TabIndex        =   85
            Top             =   1230
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
            Left            =   2100
            TabIndex        =   81
            Top             =   1890
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
            MaskAllowNegative=   0   'False
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
            Left            =   2100
            TabIndex        =   79
            Top             =   1230
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
            Left            =   5550
            TabIndex        =   83
            Top             =   570
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
            Left            =   2100
            TabIndex        =   77
            Top             =   570
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
         Begin GTMaskNum.GTMaskNum txtProxPCT 
            Height          =   360
            Left            =   3900
            TabIndex        =   245
            TabStop         =   0   'False
            Top             =   1230
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   635
            _StockProps     =   77
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
            BackColor       =   -2147483633
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
            MaskAllowNegative=   0   'False
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
         Begin GTMaskNum.GTMaskNum txtPCT 
            Height          =   360
            Left            =   7350
            TabIndex        =   246
            TabStop         =   0   'False
            Top             =   1230
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   635
            _StockProps     =   77
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
            BackColor       =   -2147483633
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
            MaskAllowNegative=   0   'False
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
         Begin GTMaskDate.GTMaskDate dcboData_Calculo 
            Height          =   360
            Left            =   150
            TabIndex        =   259
            Top             =   1905
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
            Caption         =   "Data Cálculo"
            Height          =   240
            Index           =   106
            Left            =   150
            TabIndex        =   260
            Top             =   1665
            Width           =   1155
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   240
            Index           =   99
            Left            =   7350
            TabIndex        =   243
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Esc."
            Height          =   240
            Index           =   98
            Left            =   7350
            TabIndex        =   242
            Top             =   300
            Width           =   390
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. %"
            Height          =   240
            Index           =   97
            Left            =   3900
            TabIndex        =   241
            Top             =   960
            Width           =   675
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. Esc."
            Height          =   240
            Index           =   96
            Left            =   3900
            TabIndex        =   239
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. Mens. Base"
            Height          =   240
            Index           =   34
            Left            =   2100
            TabIndex        =   76
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mensal. Base"
            Height          =   240
            Index           =   33
            Left            =   5550
            TabIndex        =   82
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Próx. Mensal."
            Height          =   240
            Index           =   32
            Left            =   2100
            TabIndex        =   78
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Desconto (%)"
            Height          =   240
            Index           =   26
            Left            =   2100
            TabIndex        =   80
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Mensalidade"
            Height          =   240
            Index           =   25
            Left            =   5550
            TabIndex        =   84
            Top             =   960
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3660
         Left            =   -74760
         TabIndex        =   126
         Top             =   735
         Width           =   8850
         Begin VB.TextBox txtD_O_DESC 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   147
            Top             =   3150
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtD_IRS 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   130
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
            TabIndex        =   133
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
            Left            =   5400
            TabIndex        =   136
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
            Left            =   5400
            TabIndex        =   139
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
            Left            =   7035
            TabIndex        =   131
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
            Left            =   7035
            TabIndex        =   134
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
            Left            =   7005
            TabIndex        =   137
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
            Left            =   7005
            TabIndex        =   140
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
            Left            =   5400
            TabIndex        =   142
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
         Begin GTMaskNum.GTMaskNum txtD_SC 
            Height          =   360
            Index           =   1
            Left            =   7005
            TabIndex        =   143
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
         Begin GTMaskNum.GTMaskNum txtD_T 
            Height          =   360
            Index           =   0
            Left            =   5400
            TabIndex        =   145
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
         Begin GTMaskNum.GTMaskNum txtD_T 
            Height          =   360
            Index           =   1
            Left            =   7005
            TabIndex        =   146
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
         Begin GTMaskNum.GTMaskNum txtD_O 
            Height          =   360
            Index           =   0
            Left            =   5400
            TabIndex        =   148
            Top             =   3135
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
            Left            =   7005
            TabIndex        =   149
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
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Transportes (Anual) ........................................................................"
            Height          =   240
            Index           =   31
            Left            =   150
            TabIndex        =   144
            Top             =   2670
            Width           =   5055
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Saúde / Crónica (Anual)................................................................."
            Height          =   240
            Index           =   11
            Left            =   150
            TabIndex        =   141
            Top             =   2190
            Width           =   5040
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Mãe"
            Height          =   240
            Index           =   3
            Left            =   7095
            TabIndex        =   128
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Pai"
            Height          =   240
            Index           =   2
            Left            =   5490
            TabIndex        =   127
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Juros / Amortização Empréstimo Habitação (Anual) ........."
            Height          =   240
            Index           =   46
            Left            =   150
            TabIndex        =   138
            Top             =   1710
            Width           =   5010
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "E R P I (Anual) ...................................................................................."
            Height          =   240
            Index           =   45
            Left            =   150
            TabIndex        =   135
            Top             =   1230
            Width           =   5100
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Segurança Social (Anual) .............................................................."
            Height          =   240
            Index           =   44
            Left            =   150
            TabIndex        =   132
            Top             =   750
            Width           =   5115
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "I R S (Anual) ........................................................................................."
            Height          =   240
            Index           =   43
            Left            =   150
            TabIndex        =   129
            Top             =   300
            Width           =   5145
         End
      End
      Begin VB.Frame fraRendimentos 
         Height          =   4560
         Left            =   -74760
         TabIndex        =   96
         Top             =   720
         Width           =   8850
         Begin VB.TextBox txtR_O_DESC 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   123
            Top             =   4080
            Width           =   5100
         End
         Begin GTMaskNum.GTMaskNum txtR_TD 
            Height          =   360
            Index           =   0
            Left            =   5430
            TabIndex        =   100
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
            TabIndex        =   103
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
            TabIndex        =   106
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
         Begin GTMaskNum.GTMaskNum txtR_R 
            Height          =   360
            Index           =   0
            Left            =   5400
            TabIndex        =   112
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
            TabIndex        =   115
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
            TabIndex        =   118
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
            TabIndex        =   121
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
            TabIndex        =   124
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
            Left            =   7065
            TabIndex        =   101
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
            Left            =   7065
            TabIndex        =   104
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
            Left            =   7035
            TabIndex        =   107
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
            Left            =   7035
            TabIndex        =   110
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
            Left            =   7035
            TabIndex        =   113
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
            Left            =   7035
            TabIndex        =   116
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
            Left            =   7035
            TabIndex        =   119
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
            Left            =   7035
            TabIndex        =   122
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
            Left            =   7035
            TabIndex        =   125
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
            Left            =   7110
            TabIndex        =   98
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Caption         =   "Pai"
            Height          =   240
            Index           =   0
            Left            =   5490
            TabIndex        =   97
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Dependente (Anual) ......................................................"
            Height          =   240
            Index           =   42
            Left            =   150
            TabIndex        =   99
            Top             =   300
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões (Anual) ................................................................................"
            Height          =   240
            Index           =   41
            Left            =   150
            TabIndex        =   102
            Top             =   750
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Pensões Alimentos (Anual) ............................................................"
            Height          =   240
            Index           =   40
            Left            =   150
            TabIndex        =   105
            Top             =   1230
            Width           =   5160
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Trabalho Independente (Anual) ...................................................."
            Height          =   240
            Index           =   39
            Left            =   150
            TabIndex        =   108
            Top             =   1710
            Width           =   5175
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Rendas (Anual) .................................................................................."
            Height          =   240
            Index           =   38
            Left            =   150
            TabIndex        =   111
            Top             =   2190
            Width           =   5130
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "R S I (Anual) ......................................................................................."
            Height          =   240
            Index           =   37
            Left            =   150
            TabIndex        =   114
            Top             =   2670
            Width           =   5055
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Subsidio Desemprego (Anual) ....................................................."
            Height          =   240
            Index           =   36
            Left            =   150
            TabIndex        =   117
            Top             =   3150
            Width           =   5145
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Abono (Anual) ...................................................................................."
            Height          =   240
            Index           =   35
            Left            =   150
            TabIndex        =   120
            Top             =   3630
            Width           =   5100
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
         TabIndex        =   20
         Top             =   2880
         Width           =   2235
         Begin GTMaskDate.GTMaskDate dcboData_Admissao 
            Height          =   375
            Left            =   150
            TabIndex        =   22
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
            TabIndex        =   24
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
            TabIndex        =   21
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Demissão"
            Height          =   240
            Index           =   29
            Left            =   135
            TabIndex        =   23
            Top             =   960
            Width           =   930
         End
      End
      Begin VB.Frame fraInstituicao 
         Caption         =   " Equipamento e Sala "
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
         Height          =   3075
         Left            =   -74850
         TabIndex        =   89
         Top             =   2370
         Width           =   8790
         Begin VB.TextBox txtNomeEducadora 
            BackColor       =   &H8000000F&
            Height          =   360
            Left            =   150
            TabIndex        =   261
            Top             =   2580
            Width           =   5250
         End
         Begin VB.TextBox txtValencia 
            BackColor       =   &H8000000F&
            Height          =   360
            Left            =   150
            TabIndex        =   238
            Top             =   1890
            Width           =   5250
         End
         Begin VB.TextBox txtSalas 
            Height          =   360
            Left            =   150
            TabIndex        =   95
            Top             =   1230
            Width           =   5250
         End
         Begin VB.TextBox txtInstituicao 
            Height          =   360
            Left            =   150
            TabIndex        =   92
            Top             =   570
            Width           =   5250
         End
         Begin SSDataWidgets_B.SSDBCombo cboInstituicao 
            Height          =   330
            Left            =   150
            TabIndex        =   91
            Top             =   570
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
            TabIndex        =   94
            Top             =   1230
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
            Caption         =   "Nome Educadora"
            Height          =   240
            Index           =   105
            Left            =   165
            TabIndex        =   258
            Top             =   2310
            Width           =   1605
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Resposta Social"
            Height          =   240
            Index           =   95
            Left            =   135
            TabIndex        =   237
            Top             =   1650
            Width           =   1500
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Sala"
            Height          =   240
            Index           =   28
            Left            =   150
            TabIndex        =   93
            Top             =   990
            Width           =   420
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Equipamento"
            Height          =   240
            Index           =   27
            Left            =   150
            TabIndex        =   90
            Top             =   330
            Width           =   1200
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
         Height          =   1440
         Left            =   -74850
         TabIndex        =   86
         Top             =   900
         Width           =   8790
         Begin VB.TextBox txtObservacoes 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   88
            Top             =   540
            Width           =   8460
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   240
            Index           =   14
            Left            =   150
            TabIndex        =   87
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Frame fraCedula 
         Caption         =   " Nº de Identificação "
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
         Left            =   -70950
         TabIndex        =   66
         Top             =   900
         Width           =   4890
         Begin VB.TextBox txtLocal_Cedula 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   72
            Top             =   1230
            Width           =   4500
         End
         Begin GTMaskNum.GTMaskNum txtNum_Cedula 
            Height          =   360
            Left            =   150
            TabIndex        =   68
            Top             =   570
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
         Begin GTMaskDate.GTMaskDate dcboData_Cedula 
            Height          =   360
            Left            =   2730
            TabIndex        =   70
            Top             =   570
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
            Caption         =   "Nº"
            Height          =   240
            Index           =   22
            Left            =   150
            TabIndex        =   67
            Top             =   330
            Width           =   225
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   240
            Index           =   23
            Left            =   2730
            TabIndex        =   69
            Top             =   330
            Width           =   435
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   240
            Index           =   24
            Left            =   150
            TabIndex        =   71
            Top             =   990
            Width           =   1035
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
         TabIndex        =   61
         Top             =   900
         Width           =   3750
         Begin VB.TextBox txtFreguesia 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   63
            Top             =   570
            Width           =   3390
         End
         Begin VB.TextBox txtConcelho 
            Height          =   360
            Left            =   150
            MaxLength       =   30
            TabIndex        =   65
            Top             =   1230
            Width           =   3390
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Freguesia"
            Height          =   240
            Index           =   20
            Left            =   150
            TabIndex        =   62
            Top             =   330
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Concellho"
            Height          =   240
            Index           =   21
            Left            =   150
            TabIndex        =   64
            Top             =   990
            Width           =   900
         End
      End
      Begin VB.Frame fraEnc_Edu 
         Caption         =   " Encarregado de Educação "
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
         TabIndex        =   43
         Top             =   900
         Width           =   8790
         Begin VB.TextBox txtTelemovel_Enc_Edu 
            Height          =   360
            Left            =   3720
            MaxLength       =   12
            TabIndex        =   51
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa_Enc_Edu 
            Height          =   360
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   49
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Enc_Edu 
            Height          =   360
            Left            =   120
            MaxLength       =   12
            TabIndex        =   47
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Enc_Edu 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   45
            Top             =   540
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   88
            Left            =   3720
            TabIndex        =   50
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   87
            Left            =   1920
            TabIndex        =   48
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   15
            Left            =   150
            TabIndex        =   46
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   16
            Left            =   150
            TabIndex        =   44
            Top             =   300
            Width           =   555
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
         TabIndex        =   34
         Top             =   2820
         Width           =   8880
         Begin VB.TextBox txtTelemovel_Mae 
            Height          =   360
            Left            =   3720
            MaxLength       =   12
            TabIndex        =   42
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa_Mae 
            Height          =   360
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   40
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Emp_Mae 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   38
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Mae 
            Height          =   360
            Left            =   135
            MaxLength       =   50
            TabIndex        =   36
            Top             =   540
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   86
            Left            =   3720
            TabIndex        =   41
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   85
            Left            =   1920
            TabIndex        =   39
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   35
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   10
            Left            =   135
            TabIndex        =   37
            Top             =   960
            Width           =   1320
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
         TabIndex        =   25
         Top             =   900
         Width           =   8880
         Begin VB.TextBox txtTelemovel_Pai 
            Height          =   360
            Left            =   3720
            MaxLength       =   12
            TabIndex        =   33
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa_Pai 
            Height          =   360
            Left            =   1920
            MaxLength       =   12
            TabIndex        =   31
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Emp_Pai 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   29
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox txtNome_Pai 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   27
            Top             =   540
            Width           =   6900
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   84
            Left            =   3720
            TabIndex        =   32
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   83
            Left            =   1920
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   26
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   7
            Left            =   150
            TabIndex        =   28
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
         Height          =   3720
         Left            =   150
         TabIndex        =   2
         Top             =   900
         Width           =   6600
         Begin VB.TextBox txtTelefone 
            Height          =   360
            Left            =   150
            MaxLength       =   12
            TabIndex        =   19
            Top             =   3210
            Width           =   1500
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   360
            Left            =   150
            MaxLength       =   8
            TabIndex        =   15
            Top             =   2550
            Width           =   1110
         End
         Begin VB.TextBox txtMorada 
            Height          =   360
            Left            =   150
            MaxLength       =   60
            TabIndex        =   13
            Top             =   1890
            Width           =   6300
         End
         Begin VB.TextBox txtLocal 
            Height          =   360
            Left            =   1305
            MaxLength       =   50
            TabIndex        =   17
            Top             =   2550
            Width           =   5160
         End
         Begin VB.TextBox txtNome 
            Height          =   360
            Left            =   150
            MaxLength       =   50
            TabIndex        =   4
            Top             =   570
            Width           =   6300
         End
         Begin VB.OptionButton optSexo 
            Height          =   300
            Index           =   0
            Left            =   2280
            TabIndex        =   8
            Top             =   1230
            Value           =   -1  'True
            Width           =   210
         End
         Begin VB.OptionButton optSexo 
            Height          =   300
            Index           =   1
            Left            =   3720
            TabIndex        =   10
            Top             =   1230
            Width           =   210
         End
         Begin GTMaskDate.GTMaskDate dcboData_Nasc 
            Height          =   360
            Left            =   150
            TabIndex        =   6
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
            Caption         =   "Telefone"
            Height          =   240
            Index           =   5
            Left            =   150
            TabIndex        =   18
            Top             =   2970
            Width           =   810
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Morada"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   12
            Top             =   1650
            Width           =   705
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Postal"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   14
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Localidade"
            Height          =   240
            Index           =   4
            Left            =   1305
            TabIndex        =   16
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Data de Nascimento"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   990
            Width           =   1845
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Sexo"
            Height          =   240
            Index           =   17
            Left            =   2280
            TabIndex        =   7
            Top             =   990
            Width           =   465
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Masculino"
            Height          =   240
            Index           =   18
            Left            =   2550
            TabIndex        =   9
            Top             =   1230
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Feminino"
            Height          =   240
            Index           =   19
            Left            =   3990
            TabIndex        =   11
            Top             =   1230
            Width           =   825
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
         TabIndex        =   1
         Top             =   1020
         Width           =   1620
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4200
         Left            =   -74685
         TabIndex        =   150
         Top             =   1050
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7408
         _Version        =   393216
         TabOrientation  =   3
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   " "
         TabPicture(0)   =   "CBESQ010.frx":0118
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblTexto(47)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblTexto(48)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTexto(49)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblTexto(50)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtDoencaCronica"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtDoencasFrequentes"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtAlergias"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtCuidadosEspeciais"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   " "
         TabPicture(1)   =   "CBESQ010.frx":0134
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblTexto(51)"
         Tab(1).Control(1)=   "lblTexto(52)"
         Tab(1).Control(2)=   "txtDoencasGraves"
         Tab(1).Control(3)=   "txtMedicamento"
         Tab(1).ControlCount=   4
         Begin VB.TextBox txtMedicamento 
            Height          =   660
            Left            =   -74850
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   161
            Top             =   1350
            Width           =   8040
         End
         Begin VB.TextBox txtDoencasGraves 
            Height          =   660
            Left            =   -74850
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   159
            Top             =   450
            Width           =   8040
         End
         Begin VB.TextBox txtCuidadosEspeciais 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   236
            Top             =   3150
            Width           =   8040
         End
         Begin VB.TextBox txtAlergias 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   156
            Top             =   2250
            Width           =   8040
         End
         Begin VB.TextBox txtDoencasFrequentes 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   154
            Top             =   1350
            Width           =   8040
         End
         Begin VB.TextBox txtDoencaCronica 
            Height          =   660
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   152
            Top             =   450
            Width           =   8040
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Toma algum medicamento permanente"
            Height          =   240
            Index           =   52
            Left            =   -74850
            TabIndex        =   160
            Top             =   1125
            Width           =   3540
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Doenças graves na família"
            Height          =   240
            Index           =   51
            Left            =   -74850
            TabIndex        =   158
            Top             =   210
            Width           =   2400
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Cuidados especiais de saúde"
            Height          =   240
            Index           =   50
            Left            =   150
            TabIndex        =   157
            Top             =   2910
            Width           =   2700
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Alergias a"
            Height          =   240
            Index           =   49
            Left            =   150
            TabIndex        =   155
            Top             =   2010
            Width           =   915
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Doenças Frequentes"
            Height          =   240
            Index           =   48
            Left            =   150
            TabIndex        =   153
            Top             =   1110
            Width           =   1890
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Doenças Crónicas"
            Height          =   240
            Index           =   47
            Left            =   150
            TabIndex        =   151
            Top             =   210
            Width           =   1665
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3075
         Left            =   -74700
         TabIndex        =   162
         Top             =   1050
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5424
         _Version        =   393216
         TabOrientation  =   3
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   " "
         TabPicture(0)   =   "CBESQ010.frx":0150
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblTexto(58)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblTexto(57)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTexto(56)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblTexto(55)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblTexto(54)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblTexto(53)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblTexto(89)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtParentescoAut(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtIdadeAut(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtOutro(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtTelemovel(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtNomeAut(0)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtTelefoneAut(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtTelefone_Casa(0)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   " "
         TabPicture(1)   =   "CBESQ010.frx":016C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblTexto(64)"
         Tab(1).Control(1)=   "lblTexto(63)"
         Tab(1).Control(2)=   "lblTexto(59)"
         Tab(1).Control(3)=   "lblTexto(60)"
         Tab(1).Control(4)=   "lblTexto(61)"
         Tab(1).Control(5)=   "lblTexto(62)"
         Tab(1).Control(6)=   "lblTexto(66)"
         Tab(1).Control(7)=   "txtParentescoAut(1)"
         Tab(1).Control(8)=   "txtIdadeAut(1)"
         Tab(1).Control(9)=   "txtNomeAut(1)"
         Tab(1).Control(10)=   "txtTelefoneAut(1)"
         Tab(1).Control(11)=   "txtTelemovel(1)"
         Tab(1).Control(12)=   "txtOutro(1)"
         Tab(1).Control(13)=   "txtTelefone_Casa(1)"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   " "
         TabPicture(2)   =   "CBESQ010.frx":0188
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblTexto(70)"
         Tab(2).Control(1)=   "lblTexto(69)"
         Tab(2).Control(2)=   "lblTexto(65)"
         Tab(2).Control(3)=   "lblTexto(67)"
         Tab(2).Control(4)=   "lblTexto(68)"
         Tab(2).Control(5)=   "lblTexto(72)"
         Tab(2).Control(6)=   "lblTexto(73)"
         Tab(2).Control(7)=   "txtParentescoAut(2)"
         Tab(2).Control(8)=   "txtIdadeAut(2)"
         Tab(2).Control(9)=   "txtNomeAut(2)"
         Tab(2).Control(10)=   "txtTelefoneAut(2)"
         Tab(2).Control(11)=   "txtTelemovel(2)"
         Tab(2).Control(12)=   "txtOutro(2)"
         Tab(2).Control(13)=   "txtTelefone_Casa(2)"
         Tab(2).ControlCount=   14
         TabCaption(3)   =   " "
         TabPicture(3)   =   "CBESQ010.frx":01A4
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lblTexto(71)"
         Tab(3).Control(1)=   "lblTexto(75)"
         Tab(3).Control(2)=   "lblTexto(76)"
         Tab(3).Control(3)=   "lblTexto(74)"
         Tab(3).Control(4)=   "lblTexto(78)"
         Tab(3).Control(5)=   "lblTexto(79)"
         Tab(3).Control(6)=   "lblTexto(80)"
         Tab(3).Control(7)=   "txtNomeAut(3)"
         Tab(3).Control(8)=   "txtIdadeAut(3)"
         Tab(3).Control(9)=   "txtParentescoAut(3)"
         Tab(3).Control(10)=   "txtTelefoneAut(3)"
         Tab(3).Control(11)=   "txtTelemovel(3)"
         Tab(3).Control(12)=   "txtOutro(3)"
         Tab(3).Control(13)=   "txtTelefone_Casa(3)"
         Tab(3).ControlCount=   14
         TabCaption(4)   =   " "
         TabPicture(4)   =   "CBESQ010.frx":01C0
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblTexto(77)"
         Tab(4).Control(1)=   "lblTexto(81)"
         Tab(4).Control(2)=   "lblTexto(82)"
         Tab(4).Control(3)=   "lblTexto(90)"
         Tab(4).Control(4)=   "lblTexto(91)"
         Tab(4).Control(5)=   "lblTexto(92)"
         Tab(4).Control(6)=   "lblTexto(93)"
         Tab(4).Control(7)=   "txtNomeAut(4)"
         Tab(4).Control(8)=   "txtIdadeAut(4)"
         Tab(4).Control(9)=   "txtParentescoAut(4)"
         Tab(4).Control(10)=   "txtTelefoneAut(4)"
         Tab(4).Control(11)=   "txtTelemovel(4)"
         Tab(4).Control(12)=   "txtOutro(4)"
         Tab(4).Control(13)=   "txtTelefone_Casa(4)"
         Tab(4).ControlCount=   14
         Begin VB.TextBox txtTelefone_Casa 
            Height          =   360
            Index           =   4
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   228
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtOutro 
            Height          =   360
            Index           =   4
            Left            =   -69930
            MaxLength       =   12
            TabIndex        =   232
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Index           =   4
            Left            =   -71550
            MaxLength       =   12
            TabIndex        =   230
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefoneAut 
            Height          =   360
            Index           =   4
            Left            =   -74775
            MaxLength       =   12
            TabIndex        =   226
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa 
            Height          =   360
            Index           =   3
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   214
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtOutro 
            Height          =   360
            Index           =   3
            Left            =   -69930
            MaxLength       =   12
            TabIndex        =   218
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Index           =   3
            Left            =   -71550
            MaxLength       =   12
            TabIndex        =   216
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefoneAut 
            Height          =   360
            Index           =   3
            Left            =   -74775
            MaxLength       =   12
            TabIndex        =   212
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa 
            Height          =   360
            Index           =   2
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   200
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtOutro 
            Height          =   360
            Index           =   2
            Left            =   -69930
            MaxLength       =   12
            TabIndex        =   204
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Index           =   2
            Left            =   -71550
            MaxLength       =   12
            TabIndex        =   202
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefoneAut 
            Height          =   360
            Index           =   2
            Left            =   -74775
            MaxLength       =   12
            TabIndex        =   198
            Top             =   1830
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa 
            Height          =   360
            Index           =   1
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   186
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtOutro 
            Height          =   360
            Index           =   1
            Left            =   -69930
            MaxLength       =   12
            TabIndex        =   190
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Index           =   1
            Left            =   -71550
            MaxLength       =   12
            TabIndex        =   188
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtTelefoneAut 
            Height          =   360
            Index           =   1
            Left            =   -74775
            MaxLength       =   12
            TabIndex        =   184
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtTelefone_Casa 
            Height          =   360
            Index           =   0
            Left            =   1830
            MaxLength       =   12
            TabIndex        =   172
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtParentescoAut 
            Height          =   360
            Index           =   4
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   224
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtIdadeAut 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   4
            Left            =   -74775
            MaxLength       =   2
            TabIndex        =   222
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtNomeAut 
            Height          =   360
            Index           =   4
            Left            =   -74775
            MaxLength       =   50
            TabIndex        =   220
            Top             =   465
            Width           =   6900
         End
         Begin VB.TextBox txtParentescoAut 
            Height          =   360
            Index           =   3
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   210
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtIdadeAut 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   3
            Left            =   -74775
            MaxLength       =   2
            TabIndex        =   208
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtNomeAut 
            Height          =   360
            Index           =   3
            Left            =   -74775
            MaxLength       =   50
            TabIndex        =   206
            Top             =   480
            Width           =   6900
         End
         Begin VB.TextBox txtTelefoneAut 
            Height          =   360
            Index           =   0
            Left            =   225
            MaxLength       =   12
            TabIndex        =   170
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtNomeAut 
            Height          =   360
            Index           =   0
            Left            =   225
            MaxLength       =   50
            TabIndex        =   164
            Top             =   465
            Width           =   6900
         End
         Begin VB.TextBox txtTelemovel 
            Height          =   360
            Index           =   0
            Left            =   3450
            MaxLength       =   12
            TabIndex        =   174
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtOutro 
            Height          =   360
            Index           =   0
            Left            =   5070
            MaxLength       =   12
            TabIndex        =   176
            Top             =   1815
            Width           =   1500
         End
         Begin VB.TextBox txtIdadeAut 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   225
            MaxLength       =   2
            TabIndex        =   166
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtParentescoAut 
            Height          =   360
            Index           =   0
            Left            =   1830
            MaxLength       =   12
            TabIndex        =   168
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtNomeAut 
            Height          =   360
            Index           =   1
            Left            =   -74775
            MaxLength       =   50
            TabIndex        =   178
            Top             =   465
            Width           =   6900
         End
         Begin VB.TextBox txtIdadeAut 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   -74775
            MaxLength       =   2
            TabIndex        =   180
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtParentescoAut 
            Height          =   360
            Index           =   1
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   182
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtNomeAut 
            Height          =   360
            Index           =   2
            Left            =   -74790
            MaxLength       =   50
            TabIndex        =   192
            Top             =   465
            Width           =   6900
         End
         Begin VB.TextBox txtIdadeAut 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   2
            Left            =   -74775
            MaxLength       =   2
            TabIndex        =   194
            Top             =   1140
            Width           =   1500
         End
         Begin VB.TextBox txtParentescoAut 
            Height          =   360
            Index           =   2
            Left            =   -73170
            MaxLength       =   12
            TabIndex        =   196
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   93
            Left            =   -73170
            TabIndex        =   227
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Outro"
            Height          =   240
            Index           =   92
            Left            =   -69930
            TabIndex        =   231
            Top             =   1575
            Width           =   480
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   91
            Left            =   -71550
            TabIndex        =   229
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   90
            Left            =   -74775
            TabIndex        =   225
            Top             =   1575
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   80
            Left            =   -73170
            TabIndex        =   213
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Outro"
            Height          =   240
            Index           =   79
            Left            =   -69930
            TabIndex        =   217
            Top             =   1575
            Width           =   480
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   78
            Left            =   -71550
            TabIndex        =   215
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   74
            Left            =   -74775
            TabIndex        =   211
            Top             =   1575
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   73
            Left            =   -73170
            TabIndex        =   199
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Outro"
            Height          =   240
            Index           =   72
            Left            =   -69930
            TabIndex        =   203
            Top             =   1575
            Width           =   480
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   68
            Left            =   -71550
            TabIndex        =   201
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   67
            Left            =   -74775
            TabIndex        =   197
            Top             =   1575
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   66
            Left            =   -73170
            TabIndex        =   185
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Outro"
            Height          =   240
            Index           =   62
            Left            =   -69930
            TabIndex        =   189
            Top             =   1575
            Width           =   480
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   61
            Left            =   -71550
            TabIndex        =   187
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   60
            Left            =   -74775
            TabIndex        =   183
            Top             =   1575
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Casa"
            Height          =   240
            Index           =   89
            Left            =   1830
            TabIndex        =   171
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   82
            Left            =   -73170
            TabIndex        =   223
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   240
            Index           =   81
            Left            =   -74775
            TabIndex        =   221
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   77
            Left            =   -74775
            TabIndex        =   219
            Top             =   225
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   76
            Left            =   -73170
            TabIndex        =   209
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   240
            Index           =   75
            Left            =   -74775
            TabIndex        =   207
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   71
            Left            =   -74775
            TabIndex        =   205
            Top             =   225
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   53
            Left            =   225
            TabIndex        =   163
            Top             =   225
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telefone Emp."
            Height          =   240
            Index           =   54
            Left            =   225
            TabIndex        =   169
            Top             =   1575
            Width           =   1320
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Telemóvel"
            Height          =   240
            Index           =   55
            Left            =   3450
            TabIndex        =   173
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Outro"
            Height          =   240
            Index           =   56
            Left            =   5070
            TabIndex        =   175
            Top             =   1575
            Width           =   480
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   240
            Index           =   57
            Left            =   225
            TabIndex        =   165
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   58
            Left            =   1830
            TabIndex        =   167
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   59
            Left            =   -74775
            TabIndex        =   177
            Top             =   225
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   240
            Index           =   63
            Left            =   -74775
            TabIndex        =   179
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   64
            Left            =   -73170
            TabIndex        =   181
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   240
            Index           =   65
            Left            =   -74775
            TabIndex        =   191
            Top             =   225
            Width           =   555
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Idade"
            Height          =   240
            Index           =   69
            Left            =   -74775
            TabIndex        =   193
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblTexto 
            AutoSize        =   -1  'True
            Caption         =   "Parentesco"
            Height          =   240
            Index           =   70
            Left            =   -73170
            TabIndex        =   195
            Top             =   900
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frmFichaUtente"
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
    Dim Index As Integer
    
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
        
        dcboData_Admissao.Locked = True
        dcboData_Admissao.CalDropDown = False
        dcboData_Admissao.TabStop = False
        
        dcboData_Demissao.Locked = True
        dcboData_Demissao.CalDropDown = False
        dcboData_Demissao.TabStop = False
        
        txtNome_Pai.Locked = True
        txtNome_Pai.TabStop = False
        txtNome_Pai.BackColor = &H8000000F
        
        txtTelefone_Emp_Pai.Locked = True
        txtTelefone_Emp_Pai.TabStop = False
        txtTelefone_Emp_Pai.BackColor = &H8000000F
        
        txtTelefone_Casa_Pai.Locked = True
        txtTelefone_Casa_Pai.TabStop = False
        txtTelefone_Casa_Pai.BackColor = &H8000000F
        
        txtTelemovel_Pai.Locked = True
        txtTelemovel_Pai.TabStop = False
        txtTelemovel_Pai.BackColor = &H8000000F
        
        txtNome_Mae.Locked = True
        txtNome_Mae.TabStop = False
        txtNome_Mae.BackColor = &H8000000F
        
        txtTelefone_Emp_Mae.Locked = True
        txtTelefone_Emp_Mae.TabStop = False
        txtTelefone_Emp_Mae.BackColor = &H8000000F
        
        txtTelefone_Casa_Mae.Locked = True
        txtTelefone_Casa_Mae.TabStop = False
        txtTelefone_Casa_Mae.BackColor = &H8000000F
        
        txtTelemovel_Mae.Locked = True
        txtTelemovel_Mae.TabStop = False
        txtTelemovel_Mae.BackColor = &H8000000F
        
        txtNome_Enc_Edu.Locked = True
        txtNome_Enc_Edu.TabStop = False
        txtNome_Enc_Edu.BackColor = &H8000000F
        
        txtTelefone_Enc_Edu.Locked = True
        txtTelefone_Enc_Edu.TabStop = False
        txtTelefone_Enc_Edu.BackColor = &H8000000F
        
        txtTelefone_Casa_Enc_Edu.Locked = True
        txtTelefone_Casa_Enc_Edu.TabStop = False
        txtTelefone_Casa_Enc_Edu.BackColor = &H8000000F
        
        txtTelemovel_Enc_Edu.Locked = True
        txtTelemovel_Enc_Edu.TabStop = False
        txtTelemovel_Enc_Edu.BackColor = &H8000000F
        
        txtFreguesia.Locked = True
        txtFreguesia.TabStop = False
        txtFreguesia.BackColor = &H8000000F
        
        txtConcelho.Locked = True
        txtConcelho.TabStop = False
        txtConcelho.BackColor = &H8000000F
        
        txtNum_Cedula.Locked = True
        txtNum_Cedula.TabStop = False
        
        dcboData_Cedula.Locked = True
        dcboData_Cedula.CalDropDown = False
        dcboData_Cedula.TabStop = False
        
        txtLocal_Cedula.Locked = True
        txtLocal_Cedula.TabStop = False
        txtLocal_Cedula.BackColor = &H8000000F
        
        txtAgregado.Locked = True
        txtAgregado.TabStop = False
        
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
    
        txtValencia.Visible = True
        txtValencia.Locked = True
        txtValencia.TabStop = False
        txtValencia.BackColor = &H8000000F
    
        txtNomeEducadora.Visible = True
        txtNomeEducadora.Locked = True
        txtNomeEducadora.TabStop = False
        txtNomeEducadora.BackColor = &H8000000F
    
        txtNum_Contribuinte.Locked = True
        txtNum_Contribuinte.TabStop = False
        txtNum_Contribuinte.BackColor = &H8000000F
        
        txtSeguranca_Social.Locked = True
        txtSeguranca_Social.TabStop = False
        txtSeguranca_Social.BackColor = &H8000000F
        
        txtSNS.Locked = True
        txtSNS.TabStop = False
        txtSNS.BackColor = &H8000000F
        
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
    
        txtDoencaCronica.Locked = True
        txtDoencaCronica.TabStop = False
        txtDoencaCronica.BackColor = &H8000000F
        
        txtDoencasFrequentes.Locked = True
        txtDoencasFrequentes.TabStop = False
        txtDoencasFrequentes.BackColor = &H8000000F
    
        txtAlergias.Locked = True
        txtAlergias.TabStop = False
        txtAlergias.BackColor = &H8000000F
    
        txtCuidadosEspeciais.Locked = True
        txtCuidadosEspeciais.TabStop = False
        txtCuidadosEspeciais.BackColor = &H8000000F
    
        txtDoencasGraves.Locked = True
        txtDoencasGraves.TabStop = False
        txtDoencasGraves.BackColor = &H8000000F
    
        txtMedicamento.Locked = True
        txtMedicamento.TabStop = False
        txtMedicamento.BackColor = &H8000000F
    
        For Index = 0 To 4
            txtNomeAut(Index).Locked = True
            txtNomeAut(Index).TabStop = False
            txtNomeAut(Index).BackColor = &H8000000F
            
            txtIdadeAut(Index).Locked = True
            txtIdadeAut(Index).TabStop = False
            txtIdadeAut(Index).BackColor = &H8000000F
            
            txtParentescoAut(Index).Locked = True
            txtParentescoAut(Index).TabStop = False
            txtParentescoAut(Index).BackColor = &H8000000F
            
            txtTelefoneAut(Index).Locked = True
            txtTelefoneAut(Index).TabStop = False
            txtTelefoneAut(Index).BackColor = &H8000000F
            
            txtTelefone_Casa(Index).Locked = True
            txtTelefone_Casa(Index).TabStop = False
            txtTelefone_Casa(Index).BackColor = &H8000000F
            
            txtTelemovel(Index).Locked = True
            txtTelemovel(Index).TabStop = False
            txtTelemovel(Index).BackColor = &H8000000F
            
            txtOutro(Index).Locked = True
            txtOutro(Index).TabStop = False
            txtOutro(Index).BackColor = &H8000000F
        Next Index
        
        txtNomeEscola.Locked = True
        txtNomeEscola.TabStop = False
        txtNomeEscola.BackColor = &H8000000F
        
        txtNomeProf.Locked = True
        txtNomeProf.TabStop = False
        txtNomeProf.BackColor = &H8000000F
        
        txtHorarioEscola.Locked = True
        txtHorarioEscola.TabStop = False
        txtHorarioEscola.BackColor = &H8000000F
        
        txtTelefoneEscola.Locked = True
        txtTelefoneEscola.TabStop = False
        txtTelefoneEscola.BackColor = &H8000000F
        
        txtAlmoco.Locked = True
        txtAlmoco.TabStop = False
        txtAlmoco.BackColor = &H8000000F
        
        dcboData_Calculo.Locked = True
        dcboData_Calculo.CalDropDown = False
        dcboData_Calculo.TabStop = False
        
    Else
        
        lblTexto(29).Visible = False
        dcboData_Demissao.Visible = False
        
        cboInstituicao.Visible = True
        txtInstituicao.Visible = False
        
        cboSalas.Visible = True
        txtSalas.Visible = False
        
        txtValencia.Visible = True
        txtValencia.Locked = True
        txtValencia.TabStop = False
        txtValencia.BackColor = &H8000000F
        
        txtNomeEducadora.Visible = True
        txtNomeEducadora.Locked = True
        txtNomeEducadora.TabStop = False
        txtNomeEducadora.BackColor = &H8000000F
        
    End If
End Sub


Public Sub CamposLimpaCarrega()
    Dim recUtente As Recordset
    Dim Index As Integer
    
    ' Ficha
    If cBotao = "Ficha" Or cBotao = "Altera" Then
        ' vai procurar o registo
        ' abre o recordset
        cSql = "SELECT * FROM UTENTES WHERE NUM_UTENTE=" & lNum_Utente
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
        txtTelefone.Text = vFiltraCamposNulos(recUtente!TELEFONE)
        dcboData_Admissao.Text = vFiltraCamposNulos(recUtente!DATA_ENTRADA)
        dcboData_Demissao.Text = vFiltraCamposNulos(recUtente!DATA_SAIDA)
        txtNome_Pai.Text = vFiltraCamposNulos(recUtente!NOME_PAI)
        txtTelefone_Emp_Pai.Text = vFiltraCamposNulos(recUtente!TEL_EMP_PAI)
        txtNome_Mae.Text = vFiltraCamposNulos(recUtente!NOME_MAE)
        txtTelefone_Emp_Mae.Text = vFiltraCamposNulos(recUtente!TEL_EMP_MAE)
        txtNome_Enc_Edu.Text = vFiltraCamposNulos(recUtente!NOME_ENC_EDU)
        txtTelefone_Enc_Edu.Text = vFiltraCamposNulos(recUtente!TEL_EMP_ENC_EDU)
        txtFreguesia.Text = vFiltraCamposNulos(recUtente!NAT_FREG)
        txtConcelho.Text = vFiltraCamposNulos(recUtente!NAT_CONC)
        txtNum_Cedula.Text = vFiltraCamposNulos(recUtente!NUM_CEDULA)
        dcboData_Cedula.Text = vFiltraCamposNulos(recUtente!DATA_CEDULA)
        txtLocal_Cedula.Text = vFiltraCamposNulos(recUtente!LOC_CEDULA)
        txtAgregado.Text = vFiltraCamposNulos(recUtente!AGREGADO)
        txtMensalidade_Base.Value = vFiltraCamposNulos(recUtente!MENSALIDADE_BASE)
        txtMensalidade.Value = vFiltraCamposNulos(recUtente!MENSALIDADE)
        txtComparticipacao.Text = vFiltraCamposNulos(recUtente!COMPARTICIPACAO)
        txtEscalao.Text = vFiltraCamposNulos(recUtente!ESCALAO)
        txtPCT.Text = vFiltraCamposNulos(recUtente!PCT)
        txtObservacoes.Text = vFiltraCamposNulos(recUtente!OBSERVACOES)
'        If cBotao = "Ficha" Then
            txtInstituicao.Text = cDescodificaInstituicao(recUtente!COD_INST)
            txtSalas.Text = cDescodificaSala(recUtente!COD_INST, recUtente!COD_SALA)
'        ElseIf cBotao = "Altera" Then
            cboInstituicao.Text = cDescodificaInstituicao(recUtente!COD_INST)
            cboSalas.Text = cDescodificaSala(recUtente!COD_INST, recUtente!COD_SALA)
'        End If
        txtValencia.Text = cDescodificaValencia(recUtente!COD_INST, recUtente!COD_SALA)
        txtNomeEducadora.Text = cDescodificaNomeEducadora(recUtente!COD_INST, recUtente!COD_SALA)
        
        If vFiltraCamposNulos(recUtente!TIPO_MENSALIDADE) Then
            optTipo_Mensalidade(0).Value = True
            ' Calcula a mensalidade com base na tabela de mensalidades
'            Call CalculaMensalidadeAutomatica
        Else
            optTipo_Mensalidade(1).Value = True
        End If
        txtProxMensalidade.Value = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE)
        txtProxMensalidade_Base.Value = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE_BASE)
        dProx_Comparticipacao = vFiltraCamposNulos(recUtente!PROX_COMPARTICIPACAO)
        dProx_Mensalidade_Base = vFiltraCamposNulos(recUtente!PROX_MENSALIDADE_BASE)
        txtProxEscalao.Text = vFiltraCamposNulos(recUtente!PROX_ESCALAO)
        txtProxPCT.Text = vFiltraCamposNulos(recUtente!PROX_PCT)
        
        txtNum_Contribuinte.Text = vFiltraCamposNulos(recUtente!NUM_CONTRIBUINTE)
        txtSeguranca_Social.Text = vFiltraCamposNulos(recUtente!NUM_SEG_SOCIAL)
        
        ' Novos Campos Rendimentos / Despesas
        txtR_TD(0).Value = vFiltraCamposNulos(recUtente!R_TD_1)
        txtR_P(0).Value = vFiltraCamposNulos(recUtente!R_P_1)
        txtR_PA(0).Value = vFiltraCamposNulos(recUtente!R_PA_1)
        txtR_TI(0).Value = vFiltraCamposNulos(recUtente!R_TI_1)
        txtR_R(0).Value = vFiltraCamposNulos(recUtente!R_R_1)
        txtR_RSI(0).Value = vFiltraCamposNulos(recUtente!R_RSI_1)
        txtR_SD(0).Value = vFiltraCamposNulos(recUtente!R_SD_1)
        txtR_AF(0).Value = vFiltraCamposNulos(recUtente!R_AF_1)
        txtR_O_DESC.Text = vFiltraCamposNulos(recUtente!R_O_DESC)
        txtR_O(0).Value = vFiltraCamposNulos(recUtente!R_O_1)
        txtD_IRS(0).Value = vFiltraCamposNulos(recUtente!D_IRS_1)
        txtD_SS(0).Value = vFiltraCamposNulos(recUtente!D_SS_1)
        txtD_IMI(0).Value = vFiltraCamposNulos(recUtente!D_IMI_1)
        txtD_JEH(0).Value = vFiltraCamposNulos(recUtente!D_JEH_1)
        txtR_TD(1).Value = vFiltraCamposNulos(recUtente!R_TD_2)
        txtR_P(1).Value = vFiltraCamposNulos(recUtente!R_P_2)
        txtR_PA(1).Value = vFiltraCamposNulos(recUtente!R_PA_2)
        txtR_TI(1).Value = vFiltraCamposNulos(recUtente!R_TI_2)
        txtR_R(1).Value = vFiltraCamposNulos(recUtente!R_R_2)
        txtR_RSI(1).Value = vFiltraCamposNulos(recUtente!R_RSI_2)
        txtR_SD(1).Value = vFiltraCamposNulos(recUtente!R_SD_2)
        txtR_AF(1).Value = vFiltraCamposNulos(recUtente!R_AF_2)
        txtR_O(1).Value = vFiltraCamposNulos(recUtente!R_O_2)
        txtD_IRS(1).Value = vFiltraCamposNulos(recUtente!D_IRS_2)
        txtD_SS(1).Value = vFiltraCamposNulos(recUtente!D_SS_2)
        txtD_IMI(1).Value = vFiltraCamposNulos(recUtente!D_IMI_2)
        txtD_JEH(1).Value = vFiltraCamposNulos(recUtente!D_JEH_2)
        
        txtD_SC(0).Value = vFiltraCamposNulos(recUtente!D_SC_1)
        txtD_T(0).Value = vFiltraCamposNulos(recUtente!D_T_1)
        txtD_O_DESC.Text = vFiltraCamposNulos(recUtente!D_O_DESC)
        txtD_O(0).Value = vFiltraCamposNulos(recUtente!D_O_1)
        txtD_SC(1).Value = vFiltraCamposNulos(recUtente!D_SC_2)
        txtD_T(1).Value = vFiltraCamposNulos(recUtente!D_T_2)
        txtD_O(1).Value = vFiltraCamposNulos(recUtente!D_O_2)
        
' **********************************************************************
        
        txtDoencaCronica.Text = vFiltraCamposNulos(recUtente!DC)
        txtDoencasFrequentes.Text = vFiltraCamposNulos(recUtente!DF)
        txtAlergias.Text = vFiltraCamposNulos(recUtente!AL)
        txtCuidadosEspeciais.Text = vFiltraCamposNulos(recUtente!CE)
        txtDoencasGraves.Text = vFiltraCamposNulos(recUtente!DG)
        txtMedicamento.Text = vFiltraCamposNulos(recUtente!Me)
        
        txtNomeAut(0).Text = vFiltraCamposNulos(recUtente!NA_1)
        txtIdadeAut(0).Text = vFiltraCamposNulos(recUtente!IA_1)
        txtParentescoAut(0).Text = vFiltraCamposNulos(recUtente!PA_1)
        txtTelefoneAut(0).Text = vFiltraCamposNulos(recUtente!TL_1)
        txtTelemovel(0).Text = vFiltraCamposNulos(recUtente!TM_1)
        txtOutro(0).Text = vFiltraCamposNulos(recUtente!TO_1)
        
        txtNomeAut(1).Text = vFiltraCamposNulos(recUtente!NA_2)
        txtIdadeAut(1).Text = vFiltraCamposNulos(recUtente!IA_2)
        txtParentescoAut(1).Text = vFiltraCamposNulos(recUtente!PA_2)
        txtTelefoneAut(1).Text = vFiltraCamposNulos(recUtente!TL_2)
        txtTelemovel(1).Text = vFiltraCamposNulos(recUtente!TM_2)
        txtOutro(1).Text = vFiltraCamposNulos(recUtente!TO_2)
        
        txtNomeAut(2).Text = vFiltraCamposNulos(recUtente!NA_3)
        txtIdadeAut(2).Text = vFiltraCamposNulos(recUtente!IA_3)
        txtParentescoAut(2).Text = vFiltraCamposNulos(recUtente!PA_3)
        txtTelefoneAut(2).Text = vFiltraCamposNulos(recUtente!TL_3)
        txtTelemovel(2).Text = vFiltraCamposNulos(recUtente!TM_3)
        txtOutro(2).Text = vFiltraCamposNulos(recUtente!TO_3)
        
        txtNomeAut(3).Text = vFiltraCamposNulos(recUtente!NA_4)
        txtIdadeAut(3).Text = vFiltraCamposNulos(recUtente!IA_4)
        txtParentescoAut(3).Text = vFiltraCamposNulos(recUtente!PA_4)
        txtTelefoneAut(3).Text = vFiltraCamposNulos(recUtente!TL_4)
        txtTelemovel(3).Text = vFiltraCamposNulos(recUtente!TM_4)
        txtOutro(3).Text = vFiltraCamposNulos(recUtente!TO_4)
        
        txtNomeAut(4).Text = vFiltraCamposNulos(recUtente!NA_5)
        txtIdadeAut(4).Text = vFiltraCamposNulos(recUtente!IA_5)
        txtParentescoAut(4).Text = vFiltraCamposNulos(recUtente!PA_5)
        txtTelefoneAut(4).Text = vFiltraCamposNulos(recUtente!TL_5)
        txtTelemovel(4).Text = vFiltraCamposNulos(recUtente!TM_5)
        txtOutro(4).Text = vFiltraCamposNulos(recUtente!TO_5)
        
        txtSNS.Text = vFiltraCamposNulos(recUtente!NUM_SNS)
        txtTelefone_Casa_Pai.Text = vFiltraCamposNulos(recUtente!TEL_CASA_PAI)
        txtTelemovel_Pai.Text = vFiltraCamposNulos(recUtente!TELEMOVEL_PAI)
        txtTelefone_Casa_Mae.Text = vFiltraCamposNulos(recUtente!TEL_CASA_MAE)
        txtTelemovel_Mae.Text = vFiltraCamposNulos(recUtente!TELEMOVEL_MAE)
        txtTelefone_Casa_Enc_Edu.Text = vFiltraCamposNulos(recUtente!TEL_CASA_ENC_EDU)
        txtTelemovel_Enc_Edu.Text = vFiltraCamposNulos(recUtente!TELEMOVEL_ENC_EDU)
        txtTelefone_Casa(0).Text = vFiltraCamposNulos(recUtente!TLC_1)
        txtTelefone_Casa(1).Text = vFiltraCamposNulos(recUtente!TLC_2)
        txtTelefone_Casa(2).Text = vFiltraCamposNulos(recUtente!TLC_3)
        txtTelefone_Casa(3).Text = vFiltraCamposNulos(recUtente!TLC_4)
        txtTelefone_Casa(4).Text = vFiltraCamposNulos(recUtente!TLC_5)
        
        txtNomeEscola.Text = vFiltraCamposNulos(recUtente!NOME_ESCOLA)
        txtNomeProf.Text = vFiltraCamposNulos(recUtente!NOME_PROF_ESCOLA)
        txtHorarioEscola.Text = vFiltraCamposNulos(recUtente!HORARIO_ESCOLA)
        txtTelefoneEscola.Text = vFiltraCamposNulos(recUtente!TEL_ESCOLA)
        txtAlmoco.Text = vFiltraCamposNulos(recUtente!ALMOCO_ESCOLA)
        
        dcboData_Calculo.Text = vFiltraCamposNulos(recUtente!DATA_CALCULO)

' **********************************************************************
        
        ' muda label Percentegem
        If (recUtente!COD_INST = "001" And recUtente!COD_SALA = "008") Or _
            (recUtente!COD_INST = "002" And recUtente!COD_SALA = "006") Then
            lblTexto(25).Caption = "1/11 Agosto"
            lblTexto(32).Caption = "1/11 Agosto"
        Else
            lblTexto(25).Caption = "1/10 Julho"
            lblTexto(32).Caption = "1/10 Julho"
        End If
        
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
        dcboData_Admissao.Text = Date
        dcboData_Demissao.Text = vbNullString
        txtNome_Pai.Text = vbNullString
        txtTelefone_Emp_Pai.Text = vbNullString
        txtNome_Mae.Text = vbNullString
        txtTelefone_Emp_Mae.Text = vbNullString
        txtNome_Enc_Edu.Text = vbNullString
        txtTelefone_Enc_Edu.Text = vbNullString
        txtFreguesia.Text = vbNullString
        txtConcelho.Text = vbNullString
        txtNum_Cedula.Text = 0
        dcboData_Cedula.Text = vbNullString
        txtAgregado.Text = 0
        optTipo_Mensalidade(1).Value = True
        txtMensalidade_Base.Text = 0
        txtMensalidade.Text = 0
        txtProxMensalidade.Text = 0
        txtProxMensalidade_Base.Text = 0
        dProx_Mensalidade_Base = 0
        dProx_Comparticipacao = 0
        txtComparticipacao.Text = 0
        txtProxEscalao.Text = vbNullString
        txtProxPCT.Text = 0
        txtEscalao.Text = vbNullString
        txtPCT.Text = 0
        txtObservacoes.Text = vbNullString
        cboInstituicao.Text = vbNullString
        cboSalas.Text = vbNullString
        txtValencia.Text = vbNullString
        txtNomeEducadora.Text = vbNullString
    
        txtNum_Contribuinte = vbNullString
        txtSeguranca_Social.Text = vbNullString
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
    
        txtDoencaCronica.Text = vbNullString
        txtDoencasFrequentes.Text = vbNullString
        txtAlergias.Text = vbNullString
        txtCuidadosEspeciais.Text = vbNullString
        txtDoencasGraves.Text = vbNullString
        txtMedicamento.Text = vbNullString
        
        txtSNS.Text = vbNullString
        txtTelefone_Casa_Pai.Text = vbNullString
        txtTelemovel_Pai.Text = vbNullString
        txtTelefone_Casa_Mae.Text = vbNullString
        txtTelemovel_Mae.Text = vbNullString
        txtTelefone_Casa_Enc_Edu.Text = vbNullString
        txtTelemovel_Enc_Edu.Text = vbNullString
        
        For Index = 0 To 4
            txtNomeAut(Index).Text = vbNullString
            txtIdadeAut(Index).Text = vbNullString
            txtParentescoAut(Index).Text = vbNullString
            txtTelefoneAut(Index).Text = vbNullString
            txtTelemovel(Index).Text = vbNullString
            txtOutro(Index).Text = vbNullString
            txtTelefone_Casa(Index).Text = vbNullString
        Next Index
        
        txtNomeEscola.Text = vbNullString
        txtNomeProf.Text = vbNullString
        txtHorarioEscola.Text = vbNullString
        txtTelefoneEscola.Text = vbNullString
        txtAlmoco.Text = vbNullString
        dcboData_Calculo.Text = vbNullString
        
    End If
End Sub
'Este procedimento activa ou desactiva o default do botao OK
Private Sub BotaoOKDefault(ByVal Propriedade As Boolean)
  cmdOK.Default = Propriedade
End Sub



Private Sub cboInstituicao_Click()
    cboSalas.Text = vbNullString
    txtValencia.Text = vbNullString
    txtNomeEducadora.Text = vbNullString
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

Private Sub cboSalas_Click()
    txtValencia.Text = cDescodificaValencia(cCodificaInstituicao(cboInstituicao.Text), cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text))
    txtNomeEducadora.Text = cDescodificaNomeEducadora(cCodificaInstituicao(cboInstituicao.Text), cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text))
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
    Dim cAno
    
    Set mProcessamento = New Processamento

    cNomeMapa = "CBESQ003.RPT"
    cAno = str$(Year(Date))

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
    cSql = "DELETE * FROM FICHA_UTENTES"
    ' apaga o registo em Temp32
    mBDUtenteTemp.Execute cSql
    
    ' carrega a variavela com o Sql
    cSql = "INSERT INTO FICHA_UTENTES IN '" & cBDComNomeUtilizador & _
            "' SELECT * FROM UTENTES WHERE NUM_UTENTE=" & lNum_Utente
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
        .Formulas(1) = "Titulo_2='" & Mapa.Titulo_2 & "'"
        .Formulas(2) = "Titulo_3='" & Mapa.Titulo_3 & "'"
        .Formulas(3) = "NomeEmpresa='JOCA ® Mod. " & Mid$(cNomeMapa, 6, InStr(cNomeMapa, ".") - 6) & "'"
        .Formulas(4) = "AnoLectivo='" & cAno & "/" & (cAno + 1) & "'"
                      
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
        Set qryUtente = mBDUtente.QueryDefs("UTENTES Insere")
        ' parametros de input
        qryUtente.Parameters("Num_Utente") = lNovoNumUtente
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
        qryUtente.Parameters("Telefone") = txtTelefone.Text
        qryUtente.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryUtente.Parameters("Tel_Emp_Pai") = txtTelefone_Emp_Pai.Text
        qryUtente.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryUtente.Parameters("Tel_Emp_Mae") = txtTelefone_Emp_Mae.Text
        qryUtente.Parameters("Nome_Enc_Edu") = txtNome_Enc_Edu.Text
        qryUtente.Parameters("Tel_Emp_Enc_Edu") = txtTelefone_Enc_Edu.Text
        qryUtente.Parameters("Nat_Freg") = txtFreguesia.Text
        qryUtente.Parameters("Nat_Conc") = txtConcelho.Text
        qryUtente.Parameters("Num_Cedula") = txtNum_Cedula.Value
        qryUtente.Parameters("Data_Cedula") = dcboData_Cedula.DateValue
        qryUtente.Parameters("Loc_Cedula") = txtLocal_Cedula.Text
        qryUtente.Parameters("Agregado") = txtAgregado.Value
        If optTipo_Mensalidade(0).Value Then
            qryUtente.Parameters("Tipo_Mensalidade") = True
        Else
            qryUtente.Parameters("Tipo_Mensalidade") = False
        End If
        qryUtente.Parameters("Prox_Mensalidade_Base") = txtProxMensalidade_Base.Value
        qryUtente.Parameters("Prox_Mensalidade") = txtProxMensalidade.Value
        qryUtente.Parameters("Prox_Comparticipacao") = dProx_Comparticipacao
        qryUtente.Parameters("Mensalidade_Base") = txtMensalidade_Base.Value
        qryUtente.Parameters("Mensalidade") = txtMensalidade.Value
        qryUtente.Parameters("Comparticipacao") = txtComparticipacao.Value
        qryUtente.Parameters("Observacoes") = txtObservacoes.Text
        qryUtente.Parameters("Utiliz") = gUtilizador.Nome
        qryUtente.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryUtente.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        ' Novos Campos Rendimentos / Despesas
        qryUtente.Parameters("R_TD_1") = txtR_TD(0).Value
        qryUtente.Parameters("R_P_1") = txtR_P(0).Value
        qryUtente.Parameters("R_PA_1") = txtR_PA(0).Value
        qryUtente.Parameters("R_TI_1") = txtR_TI(0).Value
        qryUtente.Parameters("R_R_1") = txtR_R(0).Value
        qryUtente.Parameters("R_RSI_1") = txtR_RSI(0).Value
        qryUtente.Parameters("R_SD_1") = txtR_SD(0).Value
        qryUtente.Parameters("R_AF_1") = txtR_AF(0).Value
        qryUtente.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryUtente.Parameters("R_O_1") = txtR_O(0).Value
        qryUtente.Parameters("D_IRS_1") = txtD_IRS(0).Value
        qryUtente.Parameters("D_SS_1") = txtD_SS(0).Value
        qryUtente.Parameters("D_IMI_1") = txtD_IMI(0).Value
        qryUtente.Parameters("D_JEH_1") = txtD_JEH(0).Value
        qryUtente.Parameters("R_TD_2") = txtR_TD(1).Value
        qryUtente.Parameters("R_P_2") = txtR_P(1).Value
        qryUtente.Parameters("R_PA_2") = txtR_PA(1).Value
        qryUtente.Parameters("R_TI_2") = txtR_TI(1).Value
        qryUtente.Parameters("R_R_2") = txtR_R(1).Value
        qryUtente.Parameters("R_RSI_2") = txtR_RSI(1).Value
        qryUtente.Parameters("R_SD_2") = txtR_SD(1).Value
        qryUtente.Parameters("R_AF_2") = txtR_AF(1).Value
        qryUtente.Parameters("R_O_2") = txtR_O(1).Value
        qryUtente.Parameters("D_IRS_2") = txtD_IRS(1).Value
        qryUtente.Parameters("D_SS_2") = txtD_SS(1).Value
        qryUtente.Parameters("D_IMI_2") = txtD_IMI(1).Value
        qryUtente.Parameters("D_JEH_2") = txtD_JEH(1).Value
      
        qryUtente.Parameters("D_SC_1") = txtD_SC(0).Value
        qryUtente.Parameters("D_T_1") = txtD_T(0).Value
        qryUtente.Parameters("D_O_DESC") = txtD_O_DESC.Text
        qryUtente.Parameters("D_O_1") = txtD_O(0).Value
        qryUtente.Parameters("D_SC_2") = txtD_SC(1).Value
        qryUtente.Parameters("D_T_2") = txtD_T(1).Value
        qryUtente.Parameters("D_O_2") = txtD_O(1).Value
      
' *******************************************************************
        qryUtente.Parameters("DC") = txtDoencaCronica.Text
        qryUtente.Parameters("DF") = txtDoencasFrequentes.Text
        qryUtente.Parameters("AL") = txtAlergias.Text
        qryUtente.Parameters("CE") = txtCuidadosEspeciais.Text
        qryUtente.Parameters("DG") = txtDoencasGraves.Text
        qryUtente.Parameters("ME") = txtMedicamento.Text
      
        qryUtente.Parameters("NA_1") = txtNomeAut(0).Text
        qryUtente.Parameters("IA_1") = txtIdadeAut(0).Text
        qryUtente.Parameters("PA_1") = txtParentescoAut(0).Text
        qryUtente.Parameters("TL_1") = txtTelefoneAut(0).Text
        qryUtente.Parameters("TM_1") = txtTelemovel(0).Text
        qryUtente.Parameters("TO_1") = txtOutro(0).Text
      
        qryUtente.Parameters("NA_2") = txtNomeAut(1).Text
        qryUtente.Parameters("IA_2") = txtIdadeAut(1).Text
        qryUtente.Parameters("PA_2") = txtParentescoAut(1).Text
        qryUtente.Parameters("TL_2") = txtTelefoneAut(1).Text
        qryUtente.Parameters("TM_2") = txtTelemovel(1).Text
        qryUtente.Parameters("TO_2") = txtOutro(1).Text
      
        qryUtente.Parameters("NA_3") = txtNomeAut(2).Text
        qryUtente.Parameters("IA_3") = txtIdadeAut(2).Text
        qryUtente.Parameters("PA_3") = txtParentescoAut(2).Text
        qryUtente.Parameters("TL_3") = txtTelefoneAut(2).Text
        qryUtente.Parameters("TM_3") = txtTelemovel(2).Text
        qryUtente.Parameters("TO_3") = txtOutro(2).Text
      
        qryUtente.Parameters("NA_4") = txtNomeAut(3).Text
        qryUtente.Parameters("IA_4") = txtIdadeAut(3).Text
        qryUtente.Parameters("PA_4") = txtParentescoAut(3).Text
        qryUtente.Parameters("TL_4") = txtTelefoneAut(3).Text
        qryUtente.Parameters("TM_4") = txtTelemovel(3).Text
        qryUtente.Parameters("TO_4") = txtOutro(3).Text
      
        qryUtente.Parameters("NA_5") = txtNomeAut(4).Text
        qryUtente.Parameters("IA_5") = txtIdadeAut(4).Text
        qryUtente.Parameters("PA_5") = txtParentescoAut(4).Text
        qryUtente.Parameters("TL_5") = txtTelefoneAut(4).Text
        qryUtente.Parameters("TM_5") = txtTelemovel(4).Text
        qryUtente.Parameters("TO_5") = txtOutro(4).Text

        qryUtente.Parameters("NUM_SNS") = txtSNS.Text
        qryUtente.Parameters("TEL_CASA_PAI") = txtTelefone_Casa_Pai.Text
        qryUtente.Parameters("TELEMOVEL_PAI") = txtTelemovel_Pai.Text
        qryUtente.Parameters("TEL_CASA_MAE") = txtTelefone_Casa_Mae.Text
        qryUtente.Parameters("TELEMOVEL_MAE") = txtTelemovel_Mae.Text
        qryUtente.Parameters("TEL_CASA_ENC_EDU") = txtTelefone_Casa_Enc_Edu.Text
        qryUtente.Parameters("TELEMOVEL_ENC_EDU") = txtTelemovel_Enc_Edu.Text
        qryUtente.Parameters("TLC_1") = txtTelefone_Casa(0).Text
        qryUtente.Parameters("TLC_2") = txtTelefone_Casa(1).Text
        qryUtente.Parameters("TLC_3") = txtTelefone_Casa(2).Text
        qryUtente.Parameters("TLC_4") = txtTelefone_Casa(3).Text
        qryUtente.Parameters("TLC_5") = txtTelefone_Casa(4).Text

        qryUtente.Parameters("PROX_ESCALAO") = txtProxEscalao.Text
        qryUtente.Parameters("PROX_PCT") = txtProxPCT.Value
        qryUtente.Parameters("ESCALAO") = txtEscalao.Text
        qryUtente.Parameters("PCT") = txtPCT.Value

        qryUtente.Parameters("NOME_ESCOLA") = txtNomeEscola.Text
        qryUtente.Parameters("NOME_PROF_ESCOLA") = txtNomeProf.Text
        qryUtente.Parameters("HORARIO_ESCOLA") = txtHorarioEscola.Text
        qryUtente.Parameters("TEL_ESCOLA") = txtTelefoneEscola.Text
        qryUtente.Parameters("ALMOCO_ESCOLA") = txtAlmoco.Text

        qryUtente.Parameters("DATA_CALCULO") = dcboData_Calculo.DateValue


' *******************************************************************
      
      ElseIf cBotao = "Altera" Then
        Set qryUtente = mBDUtente.QueryDefs("UTENTES Altera")
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
        qryUtente.Parameters("Telefone") = txtTelefone.Text
        qryUtente.Parameters("Nome_Pai") = txtNome_Pai.Text
        qryUtente.Parameters("Tel_Emp_Pai") = txtTelefone_Emp_Pai.Text
        qryUtente.Parameters("Nome_Mae") = txtNome_Mae.Text
        qryUtente.Parameters("Tel_Emp_Mae") = txtTelefone_Emp_Mae.Text
        qryUtente.Parameters("Nome_Enc_Edu") = txtNome_Enc_Edu.Text
        qryUtente.Parameters("Tel_Emp_Enc_Edu") = txtTelefone_Enc_Edu.Text
        qryUtente.Parameters("Nat_Freg") = txtFreguesia.Text
        qryUtente.Parameters("Nat_Conc") = txtConcelho.Text
        qryUtente.Parameters("Num_Cedula") = txtNum_Cedula.Value
        qryUtente.Parameters("Data_Cedula") = dcboData_Cedula.DateValue
        qryUtente.Parameters("Loc_Cedula") = txtLocal_Cedula.Text
        qryUtente.Parameters("Agregado") = txtAgregado.Value
        If optTipo_Mensalidade(0).Value Then
            qryUtente.Parameters("Tipo_Mensalidade") = True
        Else
            qryUtente.Parameters("Tipo_Mensalidade") = False
        End If
        qryUtente.Parameters("Prox_Mensalidade_Base") = txtProxMensalidade_Base.Value
        qryUtente.Parameters("Prox_Mensalidade") = txtProxMensalidade.Value
        qryUtente.Parameters("Prox_Comparticipacao") = dProx_Comparticipacao
        qryUtente.Parameters("Mensalidade_Base") = txtMensalidade_Base.Value
        qryUtente.Parameters("Mensalidade") = txtMensalidade.Value
        qryUtente.Parameters("Comparticipacao") = txtComparticipacao.Value
        qryUtente.Parameters("Observacoes") = txtObservacoes.Text
        qryUtente.Parameters("Utiliz") = gUtilizador.Nome
        qryUtente.Parameters("Num_Utente") = lNum_Utente
        qryUtente.Parameters("Num_Contribuinte") = txtNum_Contribuinte.Text
        qryUtente.Parameters("Num_Seg_Social") = txtSeguranca_Social.Text
        ' Novos Campos Rendimentos / Despesas
        qryUtente.Parameters("R_TD_1") = txtR_TD(0).Value
        qryUtente.Parameters("R_P_1") = txtR_P(0).Value
        qryUtente.Parameters("R_PA_1") = txtR_PA(0).Value
        qryUtente.Parameters("R_TI_1") = txtR_TI(0).Value
        qryUtente.Parameters("R_R_1") = txtR_R(0).Value
        qryUtente.Parameters("R_RSI_1") = txtR_RSI(0).Value
        qryUtente.Parameters("R_SD_1") = txtR_SD(0).Value
        qryUtente.Parameters("R_AF_1") = txtR_AF(0).Value
        qryUtente.Parameters("R_O_DESC") = txtR_O_DESC.Text
        qryUtente.Parameters("R_O_1") = txtR_O(0).Value
        qryUtente.Parameters("D_IRS_1") = txtD_IRS(0).Value
        qryUtente.Parameters("D_SS_1") = txtD_SS(0).Value
        qryUtente.Parameters("D_IMI_1") = txtD_IMI(0).Value
        qryUtente.Parameters("D_JEH_1") = txtD_JEH(0).Value
        qryUtente.Parameters("R_TD_2") = txtR_TD(1).Value
        qryUtente.Parameters("R_P_2") = txtR_P(1).Value
        qryUtente.Parameters("R_PA_2") = txtR_PA(1).Value
        qryUtente.Parameters("R_TI_2") = txtR_TI(1).Value
        qryUtente.Parameters("R_R_2") = txtR_R(1).Value
        qryUtente.Parameters("R_RSI_2") = txtR_RSI(1).Value
        qryUtente.Parameters("R_SD_2") = txtR_SD(1).Value
        qryUtente.Parameters("R_AF_2") = txtR_AF(1).Value
        qryUtente.Parameters("R_O_2") = txtR_O(1).Value
        qryUtente.Parameters("D_IRS_2") = txtD_IRS(1).Value
        qryUtente.Parameters("D_SS_2") = txtD_SS(1).Value
        qryUtente.Parameters("D_IMI_2") = txtD_IMI(1).Value
        qryUtente.Parameters("D_JEH_2") = txtD_JEH(1).Value
        
        qryUtente.Parameters("D_SC_1") = txtD_SC(0).Value
        qryUtente.Parameters("D_T_1") = txtD_T(0).Value
        qryUtente.Parameters("D_O_DESC") = txtD_O_DESC.Text
        qryUtente.Parameters("D_O_1") = txtD_O(0).Value
        qryUtente.Parameters("D_SC_2") = txtD_SC(1).Value
        qryUtente.Parameters("D_T_2") = txtD_T(1).Value
        qryUtente.Parameters("D_O_2") = txtD_O(1).Value
        
' *******************************************************************
        qryUtente.Parameters("DC") = txtDoencaCronica.Text
        qryUtente.Parameters("DF") = txtDoencasFrequentes.Text
        qryUtente.Parameters("AL") = txtAlergias.Text
        qryUtente.Parameters("CE") = txtCuidadosEspeciais.Text
        qryUtente.Parameters("DG") = txtDoencasGraves.Text
        qryUtente.Parameters("ME") = txtMedicamento.Text
      
        qryUtente.Parameters("NA_1") = txtNomeAut(0).Text
        qryUtente.Parameters("IA_1") = txtIdadeAut(0).Text
        qryUtente.Parameters("PA_1") = txtParentescoAut(0).Text
        qryUtente.Parameters("TL_1") = txtTelefoneAut(0).Text
        qryUtente.Parameters("TM_1") = txtTelemovel(0).Text
        qryUtente.Parameters("TO_1") = txtOutro(0).Text
      
        qryUtente.Parameters("NA_2") = txtNomeAut(1).Text
        qryUtente.Parameters("IA_2") = txtIdadeAut(1).Text
        qryUtente.Parameters("PA_2") = txtParentescoAut(1).Text
        qryUtente.Parameters("TL_2") = txtTelefoneAut(1).Text
        qryUtente.Parameters("TM_2") = txtTelemovel(1).Text
        qryUtente.Parameters("TO_2") = txtOutro(1).Text
      
        qryUtente.Parameters("NA_3") = txtNomeAut(2).Text
        qryUtente.Parameters("IA_3") = txtIdadeAut(2).Text
        qryUtente.Parameters("PA_3") = txtParentescoAut(2).Text
        qryUtente.Parameters("TL_3") = txtTelefoneAut(2).Text
        qryUtente.Parameters("TM_3") = txtTelemovel(2).Text
        qryUtente.Parameters("TO_3") = txtOutro(2).Text
      
        qryUtente.Parameters("NA_4") = txtNomeAut(3).Text
        qryUtente.Parameters("IA_4") = txtIdadeAut(3).Text
        qryUtente.Parameters("PA_4") = txtParentescoAut(3).Text
        qryUtente.Parameters("TL_4") = txtTelefoneAut(3).Text
        qryUtente.Parameters("TM_4") = txtTelemovel(3).Text
        qryUtente.Parameters("TO_4") = txtOutro(3).Text
      
        qryUtente.Parameters("NA_5") = txtNomeAut(4).Text
        qryUtente.Parameters("IA_5") = txtIdadeAut(4).Text
        qryUtente.Parameters("PA_5") = txtParentescoAut(4).Text
        qryUtente.Parameters("TL_5") = txtTelefoneAut(4).Text
        qryUtente.Parameters("TM_5") = txtTelemovel(4).Text
        qryUtente.Parameters("TO_5") = txtOutro(4).Text
        
        qryUtente.Parameters("NUM_SNS") = txtSNS.Text
        qryUtente.Parameters("TEL_CASA_PAI") = txtTelefone_Casa_Pai.Text
        qryUtente.Parameters("TELEMOVEL_PAI") = txtTelemovel_Pai.Text
        qryUtente.Parameters("TEL_CASA_MAE") = txtTelefone_Casa_Mae.Text
        qryUtente.Parameters("TELEMOVEL_MAE") = txtTelemovel_Mae.Text
        qryUtente.Parameters("TEL_CASA_ENC_EDU") = txtTelefone_Casa_Enc_Edu.Text
        qryUtente.Parameters("TELEMOVEL_ENC_EDU") = txtTelemovel_Enc_Edu.Text
        qryUtente.Parameters("TLC_1") = txtTelefone_Casa(0).Text
        qryUtente.Parameters("TLC_2") = txtTelefone_Casa(1).Text
        qryUtente.Parameters("TLC_3") = txtTelefone_Casa(2).Text
        qryUtente.Parameters("TLC_4") = txtTelefone_Casa(3).Text
        qryUtente.Parameters("TLC_5") = txtTelefone_Casa(4).Text
        
        qryUtente.Parameters("PROX_ESCALAO") = txtProxEscalao.Text
        qryUtente.Parameters("PROX_PCT") = txtProxPCT.Value
        qryUtente.Parameters("ESCALAO") = txtEscalao.Text
        qryUtente.Parameters("PCT") = txtPCT.Value
        
        qryUtente.Parameters("NOME_ESCOLA") = txtNomeEscola.Text
        qryUtente.Parameters("NOME_PROF_ESCOLA") = txtNomeProf.Text
        qryUtente.Parameters("HORARIO_ESCOLA") = txtHorarioEscola.Text
        qryUtente.Parameters("TEL_ESCOLA") = txtTelefoneEscola.Text
        qryUtente.Parameters("ALMOCO_ESCOLA") = txtAlmoco.Text
        
        qryUtente.Parameters("DATA_CALCULO") = dcboData_Calculo.DateValue
       
' *******************************************************************
    End If
    ' executa a query
    qryUtente.Execute dbFailOnError
    
    mWSUtente.CommitTrans
    ' faz o refresh da frmGestaoUtentes
    frmGestaoUtentes.datUtentes.Refresh
    frmGestaoUtentes.sgrdGestaoUtentes.Refresh
    
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


Private Sub dcboData_Calculo_DropDown()
    If Not IsDate(dcboData_Calculo.Text) Then
        dcboData_Calculo.DateValue = Date
    End If
End Sub


Private Sub dcboData_Cedula_DropDown()
    If Not IsDate(dcboData_Cedula.Text) Then
        dcboData_Cedula.DateValue = Date
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
        lNum_Utente = CLng(frmGestaoUtentes.sgrdGestaoUtentes.Columns(0).Text)
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


Private Sub optTipo_Mensalidade_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Calcula a mensalidade com base na tabela de mensalidades
            Call CalculaMensalidadeAutomatica
        Case 1
            ' O Calculo da mensalidade é manual
            txtProxMensalidade_Base.Text = 0
            txtProxMensalidade.Text = 0
            txtProxEscalao.Text = ""
            txtProxPCT.Text = 0
    End Select
End Sub

Public Sub CalculaMensalidadeAutomatica()
    Dim recTABMENSALIDADE As Recordset
    Dim recTABMES As Recordset
    Dim cSql
    Dim lValorPerCapita
    Dim lValorRendimentos
    Dim lValorDespesas
    Dim lValorDespesasLimitadas
    Dim lValorProxMensalidadeFinal
    Dim tEscalaoMaximo
    
    ' Calcula Rendimentos
    lValorRendimentos = (txtR_TD(0).Value + txtR_P(0).Value + txtR_PA(0).Value + txtR_TI(0).Value + txtR_R(0).Value + txtR_RSI(0).Value + txtR_SD(0).Value + txtR_AF(0).Value + txtR_O(0).Value) + _
                        (txtR_TD(1).Value + txtR_P(1).Value + txtR_PA(1).Value + txtR_TI(1).Value + txtR_R(1).Value + txtR_RSI(1).Value + txtR_SD(1).Value + txtR_AF(1).Value + txtR_O(1).Value)
    ' Calcula Despesas
    lValorDespesas = (txtD_IRS(0).Value + txtD_SS(0).Value + txtD_O(0).Value) + _
                         (txtD_IRS(1).Value + txtD_SS(1).Value + txtD_O(1).Value)
    ' Calcula Despesas Limitadas
'    lValorDespesasLimitadas = (txtD_JEH(0).Value + txtD_SC(0).Value + txtD_T(0).Value + _
'                               txtD_JEH(1).Value + txtD_SC(1).Value + txtD_T(1).Value)
' 20150515 acrescentado o campo ERPI que está no IMI para as despesas
    lValorDespesasLimitadas = (txtD_JEH(0).Value + txtD_SC(0).Value + txtD_T(0).Value + txtD_IMI(0).Value + _
                               txtD_JEH(1).Value + txtD_SC(1).Value + txtD_T(1).Value + txtD_IMI(1).Value)
    If lValorDespesasLimitadas > 6060 Then
        lValorDespesasLimitadas = 6060
    End If
    ' Calcula Valor Per-Capita segundo formula dada pelo CBESQ
'    lValorPerCapita = (((txtR_TD(0).Value + txtR_P(0).Value + txtR_PA(0).Value + txtR_TI(0).Value + txtR_R(0).Value + txtR_RSI(0).Value + txtR_SD(0).Value + txtR_AF(0).Value + txtR_O(0).Value) + _
'                        (txtR_TD(1).Value + txtR_P(1).Value + txtR_PA(1).Value + txtR_TI(1).Value + txtR_R(1).Value + txtR_RSI(1).Value + txtR_SD(1).Value + txtR_AF(1).Value + txtR_O(1).Value)) - _
'                        ((txtD_IRS(0).Value + txtD_SS(0).Value + txtD_IMI(0).Value + txtD_O(0).Value) + _
'                         (txtD_IRS(1).Value + txtD_SS(1).Value + txtD_IMI(1).Value + txtD_O(1).Value) + _
'                        lValorDespesasLimitadas)) / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
' 20150515 acrescentado o campo ERPI que está no IMI para as despesas passou para as Despesas limitadas
'    lValorPerCapita = (((txtR_TD(0).Value + txtR_P(0).Value + txtR_PA(0).Value + txtR_TI(0).Value + txtR_R(0).Value + txtR_RSI(0).Value + txtR_SD(0).Value + txtR_AF(0).Value + txtR_O(0).Value) + _
'                        (txtR_TD(1).Value + txtR_P(1).Value + txtR_PA(1).Value + txtR_TI(1).Value + txtR_R(1).Value + txtR_RSI(1).Value + txtR_SD(1).Value + txtR_AF(1).Value + txtR_O(1).Value)) - _
'                        ((txtD_IRS(0).Value + txtD_SS(0).Value + txtD_O(0).Value) + _
'                         (txtD_IRS(1).Value + txtD_SS(1).Value + txtD_O(1).Value) + _
'                        lValorDespesasLimitadas)) / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
     lValorPerCapita = (lValorRendimentos - (lValorDespesas + lValorDespesasLimitadas)) / IIf(txtAgregado.Text = 0, 1 * 12, txtAgregado.Text * 12)
                        
    'Verifica se é para por no Escalão Máximo
    If (lValorRendimentos + lValorDespesas + lValorDespesasLimitadas) > 0 Then
        tEscalaoMaximo = False
    Else
        tEscalaoMaximo = True
    End If
                        
    ' Vai a TABMENSALIDADES buscar o valor da Mensalidade Base
    If tEscalaoMaximo Then
        cSql = "SELECT * FROM TABMENSALIDADE WHERE COD_INST = '" & cCodificaInstituicao(cboInstituicao.Text) & "'"
        cSql = cSql & " AND COD_SALA = '" & cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) & "'"
        cSql = cSql & " AND COD_MENSALIDADE = '006'"
    Else
        cSql = "SELECT * FROM TABMENSALIDADE WHERE VALOR_MAX >= " & Int(lValorPerCapita) & " AND " & _
            "VALOR_MIN <= " & Int(lValorPerCapita)
        cSql = cSql & " AND COD_INST = '" & cCodificaInstituicao(cboInstituicao.Text) & "'"
        cSql = cSql & " AND COD_SALA = '" & cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) & "'"
    End If
    Set recTABMENSALIDADE = mBDUtente.OpenRecordset(cSql, dbOpenSnapshot)
    
    ' vai por o Escalão e a Percentagem
    txtProxEscalao.Text = recTABMENSALIDADE!COD_MENSALIDADE
    If tEscalaoMaximo Then
        txtProxPCT.Text = 100
    Else
        txtProxPCT.Text = recTABMENSALIDADE!PCT
    End If
    
    If tEscalaoMaximo Then
        txtProxMensalidade_Base.Text = Round(recTABMENSALIDADE!MENS_MAX, 2)
    Else
        txtProxMensalidade_Base.Text = Round(lValorPerCapita * (recTABMENSALIDADE!PCT / 100), 2)
    End If
    
    ' se o calculo for inferior ao MENS_MIN passa a MENS_MIN
    If txtProxMensalidade_Base.Value < recTABMENSALIDADE!MENS_MIN Then
        txtProxMensalidade_Base.Value = recTABMENSALIDADE!MENS_MIN
    ' se o calculo for superior ao MENS_MAX passa a MENS_MAX
    ElseIf txtProxMensalidade_Base.Value > recTABMENSALIDADE!MENS_MAX Then
        txtProxMensalidade_Base.Value = recTABMENSALIDADE!MENS_MAX
    End If
    
'   Desconto tem de ser aqui
    lValorProxMensalidadeFinal = txtProxMensalidade_Base.Value
    txtProxMensalidade_Base.Value = lValorProxMensalidadeFinal - (lValorProxMensalidadeFinal * (txtComparticipacao.Value / 100))

' Vai a TABMESES buscar o valor da percentagem a adicionar a Mensalidade Base
    cSql = "SELECT PERCENTAGEMMENSAL,PERCENTAGEMMENSALATL FROM TABMESES WHERE COD_MES = '" & Format$(Month(Date), "00") & "'"
    Set recTABMES = mBDUtente.OpenRecordset(cSql, dbOpenSnapshot)

' Se Sala for ATL
' Divide a Mensalidade Base por 11 Mês de Agosto
    If (cCodificaInstituicao(cboInstituicao.Text) = "001" And cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) = "008") Or _
        (cCodificaInstituicao(cboInstituicao.Text) = "002" And cCodificaSala(cCodificaInstituicao(cboInstituicao.Text), cboSalas.Text) = "006") Then
        txtProxMensalidade.Value = Round(txtProxMensalidade_Base.Value / recTABMES!PERCENTAGEMMENSALATL, 2)
' Se Sala for Outras Valências
' Divide a Mensalidade Base por 10 Mês de Julho
    Else
        txtProxMensalidade.Value = Round(txtProxMensalidade_Base.Value / recTABMES!PERCENTAGEMMENSAL, 2)
    End If
   
    dProx_Comparticipacao = 0
    
    recTABMENSALIDADE.Close
    Set recTABMENSALIDADE = Nothing
End Sub



Private Sub txtAgregado_LostFocus()
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub








Private Sub txtD_IMI_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtD_IRS_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtD_JEH_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtD_SS_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_AF_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_O_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_P_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_PA_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_R_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub




Private Sub txtR_SD_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_TD_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


Private Sub txtR_TI_LostFocus(Index As Integer)
    If optTipo_Mensalidade(0).Value Then
        ' Calcula a mensalidade com base na tabela de mensalidades
        Call CalculaMensalidadeAutomatica
    End If
End Sub


