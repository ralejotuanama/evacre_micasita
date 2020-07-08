VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_EvaEmp_51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   4875
   ClientTop       =   2040
   ClientWidth     =   12990
   Icon            =   "EvaCre_frm_053.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      _Version        =   65536
      _ExtentX        =   22939
      _ExtentY        =   11139
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton Command1 
            Height          =   585
            Left            =   1230
            Picture         =   "EvaCre_frm_053.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_053.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12330
            Picture         =   "EvaCre_frm_053.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_053.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerEva 
            Height          =   585
            Left            =   1830
            Picture         =   "EvaCre_frm_053.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Evaluación de Empresas"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4305
         Left            =   30
         TabIndex        =   3
         Top             =   1950
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
         _ExtentY        =   7594
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   3945
            Left            =   30
            TabIndex        =   4
            Top             =   330
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   6959
            _Version        =   393216
            Rows            =   20
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento de Identidad"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Top             =   60
            Width           =   6525
            _Version        =   65536
            _ExtentX        =   11509
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   8670
            TabIndex        =   7
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ult. Evaluac."
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   10260
            TabIndex        =   8
            Top             =   60
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   10
            Top             =   60
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Empresas"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "EvaCre_frm_053.frx":1636
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   435
         Left            =   30
         TabIndex        =   11
         Top             =   1470
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
         _ExtentY        =   767
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Clasif 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   60
            Width           =   3345
         End
         Begin VB.Label Label1 
            Caption         =   "Clasificación:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frm_EvaEmp_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

