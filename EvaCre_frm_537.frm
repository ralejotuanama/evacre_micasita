VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_EvaCre_71 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "EvaCre_frm_537.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9420
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   16616
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
         Height          =   1785
         Left            =   60
         TabIndex        =   31
         Top             =   6195
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   3149
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
         Begin VB.CommandButton cmd_VerArc 
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   10980
            TabIndex        =   11
            ToolTipText     =   "Adjuntar Croquis del Negocio"
            Top             =   1290
            Width           =   465
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   10500
            TabIndex        =   10
            ToolTipText     =   "Seleccionar archivo"
            Top             =   1290
            Width           =   465
         End
         Begin VB.CommandButton cmd_VerArc 
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   10980
            TabIndex        =   8
            ToolTipText     =   "Adjuntar Croquis del Negocio"
            Top             =   930
            Width           =   465
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   10500
            TabIndex        =   7
            ToolTipText     =   "Seleccionar archivo"
            Top             =   930
            Width           =   465
         End
         Begin VB.CommandButton cmd_VerArc 
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   10980
            TabIndex        =   5
            ToolTipText     =   "Adjuntar Croquis del Negocio"
            Top             =   540
            Width           =   465
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   10500
            TabIndex        =   4
            ToolTipText     =   "Seleccionar archivo"
            Top             =   540
            Width           =   465
         End
         Begin VB.CommandButton cmd_VerArc 
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   10980
            TabIndex        =   2
            ToolTipText     =   "Adjuntar Croquis del Negocio"
            Top             =   150
            Width           =   465
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   10500
            TabIndex        =   1
            ToolTipText     =   "Seleccionar archivo"
            Top             =   150
            Width           =   465
         End
         Begin Threed.SSPanel pnl_ArcItem 
            Height          =   315
            Index           =   0
            Left            =   1950
            TabIndex        =   0
            Top             =   150
            Width           =   8445
            _Version        =   65536
            _ExtentX        =   14896
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_ArcItem 
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   3
            Top             =   540
            Width           =   8445
            _Version        =   65536
            _ExtentX        =   14896
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_ArcItem 
            Height          =   315
            Index           =   2
            Left            =   1950
            TabIndex        =   6
            Top             =   930
            Width           =   8445
            _Version        =   65536
            _ExtentX        =   14896
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_ArcItem 
            Height          =   315
            Index           =   3
            Left            =   1950
            TabIndex        =   9
            Top             =   1320
            Width           =   8445
            _Version        =   65536
            _ExtentX        =   14896
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label6 
            Caption         =   "Cargar archivo Item4:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1350
            Width           =   1545
         End
         Begin VB.Label Label5 
            Caption         =   "Cargar archivo Item3:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "Cargar archivo Item2:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   570
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "Cargar archivo Item1:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   180
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   60
         TabIndex        =   16
         Top             =   1470
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   390
            Width           =   10005
            _Version        =   65536
            _ExtentX        =   17648
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Left            =   1440
            TabIndex        =   18
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9090
            TabIndex        =   19
            Top             =   60
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/9999"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Solicitud:"
            Height          =   195
            Left            =   7710
            TabIndex        =   22
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   60
            TabIndex        =   20
            Top             =   450
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   23
         Top             =   780
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   1138
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10860
            Picture         =   "EvaCre_frm_537.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_537.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   660
            TabIndex        =   25
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
            Begin MSComDlg.CommonDialog dlg_Guarda 
               Left            =   10260
               Top             =   90
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   660
            TabIndex        =   26
            Top             =   330
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Carga de Anexos MicroEmpresario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "EvaCre_frm_537.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   2235
         Left            =   60
         TabIndex        =   27
         Top             =   2280
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   3942
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   2145
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   3784
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "EvaCre_frm_537.frx":0B9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_537.frx":0BB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "EvaCre_frm_537.frx":0BD2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos del Crédito"
            TabPicture(3)   =   "EvaCre_frm_537.frx":0BEE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1725
               Index           =   0
               Left            =   60
               TabIndex        =   40
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3043
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1725
               Index           =   1
               Left            =   -74940
               TabIndex        =   41
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3043
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1725
               Index           =   3
               Left            =   -74940
               TabIndex        =   42
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3043
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1725
               Index           =   2
               Left            =   -74940
               TabIndex        =   43
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3043
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label7 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74910
               TabIndex        =   44
               Top             =   1470
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1605
         Left            =   60
         TabIndex        =   28
         Top             =   4560
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   2831
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
         Begin MSFlexGridLib.MSFlexGrid grd_ListDoc 
            Height          =   1185
            Left            =   30
            TabIndex        =   29
            Top             =   330
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   2090
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            Top             =   60
            Width           =   9885
            _Version        =   65536
            _ExtentX        =   17436
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Item"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   1305
         Left            =   60
         TabIndex        =   37
         Top             =   8040
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   2302
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
         Begin VB.TextBox txt_Coment 
            Height          =   1095
            Left            =   1920
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   120
            Width           =   9495
         End
         Begin VB.Label Label8 
            Caption         =   "Comentario:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   540
            Width           =   1545
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_71"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Private Sub cmd_BuscaArc_Click(Index As Integer)
   On Error GoTo cmd_BusArc_Error
   dlg_Guarda.Filter = "Todos los archivos (*.*)|*.*"
   dlg_Guarda.ShowOpen
   
   Select Case Index
      Case Index: pnl_ArcItem(Index).Caption = UCase(dlg_Guarda.FileName)
                  pnl_ArcItem(Index).Tag = 0
      If Trim(pnl_ArcItem(Index).Caption) <> "" Then
         cmd_VerArc(Index).Enabled = True
      End If
      Exit Sub
   End Select
   
cmd_BusArc_Error:
   pnl_ArcItem(Index).Caption = ""
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_Contad     As Integer
Dim r_str_ArcNue     As String

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      g_str_Parame = "USP_MIC_ANXVAR ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   
      For r_int_Contad = 0 To 3
         If Trim(pnl_ArcItem(r_int_Contad).Caption) = "" Then
            g_str_Parame = g_str_Parame & "'', "
         Else
            If InStr(Trim(pnl_ArcItem(r_int_Contad).Caption), "\") > 0 Then
               r_str_ArcNue = fs_NomArc(r_int_Contad)
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_ArcNue) & "', "
            Else
               g_str_Parame = g_str_Parame & "'" & Trim(pnl_ArcItem(r_int_Contad).Caption) & "', "
            End If
         End If
      Next r_int_Contad
      
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Coment.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & moddat_g_int_FlgAct_2 & ") "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
   
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Screen.MousePointer = 0
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   
      Screen.MousePointer = 0
   Loop
   
   'Copia el archivo en la ruta general
   If moddat_g_int_FlgGOK = True Then
      For r_int_Contad = 0 To 3
         If pnl_ArcItem(r_int_Contad).Tag = 0 Then
            r_str_ArcNue = fs_NomArc(r_int_Contad)
            If r_str_ArcNue <> "" Then
                If gf_Existe_Archivo(moddat_g_str_RutAnx & "\", r_str_ArcNue) Then
                    Kill moddat_g_str_RutAnx & "\" & r_str_ArcNue
                End If
                FileCopy Trim(pnl_ArcItem(r_int_Contad).Caption), moddat_g_str_RutAnx & "\" & r_str_ArcNue
            End If
         End If
      Next r_int_Contad
   End If
   
   MsgBox "Se grabó satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Function fs_NomArc(ByVal p_indice As Integer) As String
Dim r_str_NomArc     As String
Dim r_str_ExtArc     As String
Dim r_str_CadAux     As String

   fs_NomArc = ""
   r_str_NomArc = Mid(Trim(Trim(pnl_ArcItem(p_indice).Caption)), InStrRev(Trim(pnl_ArcItem(p_indice).Caption), "\") + 1)
   If r_str_NomArc <> "" Then
      r_str_ExtArc = Mid(r_str_NomArc, InStrRev(r_str_NomArc, ".") + 1)
      r_str_CadAux = Mid(r_str_NomArc, 1, InStr(r_str_NomArc, "." & r_str_ExtArc) - 1)
      fs_NomArc = moddat_g_str_NumSol & "_Anexo" & p_indice + 1 & "_" & r_str_CadAux & "." & r_str_ExtArc
   End If
End Function

Private Sub cmd_VerArc_Click(Index As Integer)
   If pnl_ArcItem(Index).Caption <> "" Then
      If InStr(pnl_ArcItem(Index).Caption, "\") > 0 Then
         ShellExecute Me.hwnd, "Open", pnl_ArcItem(Index).Caption, "", "", 1
      Else
         ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & pnl_ArcItem(Index).Caption, "", "", 1
      End If
   End If
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Limpia
   Call fs_Inicia
   'Call fs_Activa(False)

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Información del Inmueble
   Call modmip_gs_DatCre(grd_Listad(3), r_arr_Mtz)
   
   Call gs_CentraForm(Me)
   Call fs_Carga_Docume
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contad     As Integer
       
    grd_ListDoc.ColWidth(0) = 1155
    grd_ListDoc.ColWidth(1) = 9880
    grd_ListDoc.ColAlignment(0) = flexAlignCenterCenter
    grd_ListDoc.ColAlignment(1) = flexAlignLeftCenter
    
    'Inicializando Grid de Cliente y de Cónyuge
    For r_int_Contad = 0 To 3
       grd_Listad(r_int_Contad).ColWidth(0) = 4200:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter '2900
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       
       Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
'       cmd_VerArc(r_int_Contad).Enabled = False
    Next r_int_Contad
End Sub

Private Sub fs_Limpia()
'   txt_NomArc.Text = ""
   Call gs_LimpiaGrid(grd_Listad(0))
   Call gs_LimpiaGrid(grd_Listad(1))
   Call gs_LimpiaGrid(grd_Listad(2))
   Call gs_LimpiaGrid(grd_Listad(3))
   Call gs_LimpiaGrid(grd_ListDoc)
End Sub

Private Sub fs_Buscar()
   pnl_ArcItem(0).Tag = 0
   pnl_ArcItem(1).Tag = 0
   pnl_ArcItem(2).Tag = 0
   pnl_ArcItem(3).Tag = 0

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT ANXVAR_ANXITE1, ANXVAR_ANXITE2, ANXVAR_ANXITE3, ANXVAR_ANXITE4, ANXVAR_COMANX "
   g_str_Parame = g_str_Parame & "    FROM MIC_ANXVAR WHERE ANXVAR_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       moddat_g_int_FlgAct_2 = 1
       Exit Sub
   End If
   
   moddat_g_int_FlgAct_2 = 2
   g_rst_Princi.MoveFirst
      
   If Not IsNull(g_rst_Princi!ANXVAR_ANXITE1) Then
      pnl_ArcItem(0).Caption = Trim(g_rst_Princi!ANXVAR_ANXITE1) & "  "
      pnl_ArcItem(0).Tag = 1
      cmd_VerArc(0).Enabled = True
   End If
   If Not IsNull(g_rst_Princi!ANXVAR_ANXITE2) Then
      pnl_ArcItem(1).Caption = Trim(g_rst_Princi!ANXVAR_ANXITE2) & "  "
      pnl_ArcItem(1).Tag = 1
      cmd_VerArc(1).Enabled = True
   End If
   If Not IsNull(g_rst_Princi!ANXVAR_ANXITE3) Then
      pnl_ArcItem(2).Caption = Trim(g_rst_Princi!ANXVAR_ANXITE3) & "  "
      pnl_ArcItem(2).Tag = 1
      cmd_VerArc(2).Enabled = True
   End If
   If Not IsNull(g_rst_Princi!ANXVAR_ANXITE4) Then
      pnl_ArcItem(3).Caption = Trim(g_rst_Princi!ANXVAR_ANXITE4) & "  "
      pnl_ArcItem(3).Tag = 1
      cmd_VerArc(3).Enabled = True
   End If
   If Not IsNull(g_rst_Princi!ANXVAR_COMANX) Then
      txt_Coment.Text = Trim(g_rst_Princi!ANXVAR_COMANX)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_Docume()
   Call gs_LimpiaGrid(grd_ListDoc)
   
   'Documentos Crediticios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * "
   g_str_Parame = g_str_Parame & "    FROM MNT_PARDES  "
   g_str_Parame = g_str_Parame & "   WHERE PARDES_CODGRP = '530' "
   g_str_Parame = g_str_Parame & "     AND PARDES_CODITE <> '000000'  "
   g_str_Parame = g_str_Parame & "     AND PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_ListDoc.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_ListDoc.Rows = grd_ListDoc.Rows + 1
         grd_ListDoc.Row = grd_ListDoc.Rows - 1
         
         grd_ListDoc.Col = 0: grd_ListDoc.Text = CInt(g_rst_Genera!PARDES_CODITE)
         grd_ListDoc.Col = 1: grd_ListDoc.Text = Trim(g_rst_Genera!PARDES_DESCRI)
         g_rst_Genera.MoveNext
      Loop
      grd_ListDoc.Redraw = True
   End If
   
'   'Cargando Documentos
'   ReDim modatecli_g_arr_DocCre(0)
'
'   grd_ListDoc.Redraw = False
'   For r_int_Contad = 0 To grd_ListDoc.Rows - 1
'      grd_ListDoc.Row = r_int_Contad
'
'      grd_ListDoc.Col = 1
'      If grd_ListDoc.Text = "X" Then
'         ReDim Preserve modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre) + 1)
'
'         grd_ListDoc.Col = 2
'         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_TipDoc = CInt(grd_ListDoc.Text)
'      End If
'   Next r_int_Contad
'
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub grd_ListDoc_SelChange()
   If grd_ListDoc.Rows > 2 Then
      grd_ListDoc.RowSel = grd_ListDoc.Row
   End If
End Sub
Private Sub txt_Coment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,@#$%&;:()/º")
   End If
End Sub
