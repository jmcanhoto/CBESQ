VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   2385
   ClientTop       =   1905
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00C0C0C0&
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   8895
      Begin VB.PictureBox picLogo 
         Height          =   2895
         Left            =   285
         Picture         =   "CBESQ001.frx":0000
         ScaleHeight     =   2835
         ScaleWidth      =   3540
         TabIndex        =   1
         Top             =   405
         Width           =   3600
      End
      Begin VB.Label lblNomeProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sistema de Gestão de Sócios, Utentes e Funcionários "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   4005
         TabIndex        =   5
         Tag             =   "8"
         Top             =   1065
         Width           =   4590
      End
      Begin VB.Label lblNomeEmpresa 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Joca Software®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   4005
         TabIndex        =   4
         Tag             =   "7"
         Top             =   390
         Width           =   4590
      End
      Begin VB.Label lblPlataforma 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Windows 95/98"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4005
         TabIndex        =   3
         Tag             =   "9"
         Top             =   2685
         Width           =   1815
      End
      Begin VB.Label lblVersao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versão 1.00.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5925
         TabIndex        =   2
         Top             =   2850
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CenterMe Me
    lblVersao.Caption = "Versão " & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "00") '"Versão 1.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Form As Form
    For Each Form In Forms
        If Form Is Me Then Set Form = Nothing: Exit For
    Next
End Sub

