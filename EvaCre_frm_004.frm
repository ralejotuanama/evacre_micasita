VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_EvaCre_07 
   Caption         =   "Form2"
   ClientHeight    =   9765
   ClientLeft      =   570
   ClientTop       =   840
   ClientWidth     =   13590
   LinkTopic       =   "Form2"
   ScaleHeight     =   9765
   ScaleWidth      =   13590
   Begin Threed.SSPanel SSPanel1 
      Height          =   9765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _Version        =   65536
      _ExtentX        =   23945
      _ExtentY        =   17224
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
         Height          =   3285
         Left            =   30
         TabIndex        =   23
         Top             =   5610
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
         _ExtentY        =   5794
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
         Begin VB.TextBox Text1 
            Height          =   645
            Left            =   1740
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Text            =   "EvaCre_frm_004.frx":0000
            Top             =   2580
            Width           =   11685
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   45
            Left            =   60
            TabIndex        =   27
            Top             =   2160
            Width           =   13395
            _Version        =   65536
            _ExtentX        =   23627
            _ExtentY        =   79
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
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2250
            Width           =   11685
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2055
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   13395
            _ExtentX        =   23627
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label11 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   2580
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Situación Llamada:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   2250
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2865
         Left            =   30
         TabIndex        =   1
         Top             =   1890
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
         _ExtentY        =   5054
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
         Begin MSFlexGridLib.MSFlexGrid grd_Inm_Listad 
            Height          =   2475
            Left            =   60
            TabIndex        =   2
            Top             =   360
            Width           =   13395
            _ExtentX        =   23627
            _ExtentY        =   4366
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   90
            TabIndex        =   3
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Verif."
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   60
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Actividad Económica"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   4290
            TabIndex        =   5
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Persona o Empresa"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   10410
            TabIndex        =   6
            Top             =   60
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
            ForeColor       =   16777215
            BackColor       =   32768
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
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   4800
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
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
         Begin VB.CommandButton cmd_RegLla 
            Height          =   675
            Left            =   12060
            Picture         =   "EvaCre_frm_004.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12780
            Picture         =   "EvaCre_frm_004.frx":030E
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
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
            TabIndex        =   11
            Top             =   60
            Width           =   8415
            _Version        =   65536
            _ExtentX        =   14843
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Créditos - Verificación Telefónica"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "EvaCre_frm_004.frx":0750
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1740
            TabIndex        =   13
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1740
            TabIndex        =   14
            Top             =   390
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   10050
            TabIndex        =   15
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_IngIns 
            Height          =   315
            Left            =   10050
            TabIndex        =   16
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1740
            TabIndex        =   17
            Top             =   720
            Width           =   11685
            _Version        =   65536
            _ExtentX        =   20611
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label10 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label9 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8400
            TabIndex        =   19
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8400
            TabIndex        =   18
            Top             =   390
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   30
         Top             =   8940
         Width           =   13485
         _Version        =   65536
         _ExtentX        =   23786
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12780
            Picture         =   "EvaCre_frm_004.frx":0A5A
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton Command1 
            Height          =   675
            Left            =   12060
            Picture         =   "EvaCre_frm_004.frx":0D64
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

