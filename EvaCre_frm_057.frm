VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   1305
   ClientTop       =   1740
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   12525
   Begin Threed.SSPanel SSPanel1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12525
      _Version        =   65536
      _ExtentX        =   22093
      _ExtentY        =   12303
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_057.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Criterios de Búsqueda por Documento de Identidad"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_057.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Empresa por Documento de Identidad"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11820
            Picture         =   "EvaCre_frm_057.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3495
         Left            =   30
         TabIndex        =   5
         Top             =   3420
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
         _ExtentY        =   6165
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCli 
            Height          =   3375
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   12345
            _ExtentX        =   21775
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   13
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
         TabIndex        =   7
         Top             =   30
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
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
            TabIndex        =   8
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Empresas Empleadoras"
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
            Picture         =   "EvaCre_frm_057.frx":0A56
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   60
            Width           =   3735
         End
         Begin VB.CommandButton cmd_LimBus 
            Height          =   585
            Left            =   11820
            Picture         =   "EvaCre_frm_057.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar Criterios de Búsqueda Alfabética"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   11220
            Picture         =   "EvaCre_frm_057.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Buscar Empresas Alfabéticamente"
            Top             =   30
            Width           =   585
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   390
            Width           =   3735
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Left            =   30
            TabIndex        =   17
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
         _ExtentY        =   1349
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   60
            Width           =   3735
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   390
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

