VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_EvaEmp_52 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   4845
   ClientTop       =   1785
   ClientWidth     =   9660
   Icon            =   "EvaCre_frm_054.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      _Version        =   65536
      _ExtentX        =   17066
      _ExtentY        =   11404
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   2295
         Left            =   30
         TabIndex        =   1
         Top             =   1470
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
         _ExtentY        =   4048
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
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
            Left            =   690
            TabIndex        =   3
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
            Picture         =   "EvaCre_frm_054.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2595
         Left            =   30
         TabIndex        =   4
         Top             =   3810
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
         _ExtentY        =   4577
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
            Height          =   2235
            Index           =   1
            Left            =   60
            TabIndex        =   5
            Top             =   330
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   3942
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Evaluación"
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
            Left            =   3540
            TabIndex        =   7
            Top             =   60
            Width           =   5655
            _Version        =   65536
            _ExtentX        =   9975
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   1710
            TabIndex        =   8
            Top             =   60
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Evaluación"
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
      Begin Threed.SSPanel SSPanel14 
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
            Left            =   630
            Picture         =   "EvaCre_frm_054.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Imprimir Ficha de Empresa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "EvaCre_frm_054.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Modificación de Datos Generales de la Empresa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueEva 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_054.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Nueva Evaluación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8970
            Picture         =   "EvaCre_frm_054.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerEva 
            Height          =   585
            Left            =   1230
            Picture         =   "EvaCre_frm_054.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Consultar Evaluación"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaEmp_52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

