VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_EvaCre_68 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "EvaCre_frm_534.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9645
      Left            =   -90
      TabIndex        =   0
      Top             =   -90
      Width           =   11745
      _Version        =   65536
      _ExtentX        =   20717
      _ExtentY        =   17013
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   150
         TabIndex        =   1
         Top             =   150
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   690
            TabIndex        =   2
            Top             =   30
            Width           =   2865
            _Version        =   65536
            _ExtentX        =   5054
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
         End
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   3
            Top             =   330
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - MicroEmpresarios"
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
            Picture         =   "EvaCre_frm_534.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   150
         TabIndex        =   4
         Top             =   870
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_VerNeg 
            Height          =   585
            Left            =   1170
            Picture         =   "EvaCre_frm_534.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Carga de Anexos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerLab 
            Height          =   585
            Left            =   600
            Picture         =   "EvaCre_frm_534.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ratios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_534.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPer 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_534.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Verificaciones del Negocio"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   150
         TabIndex        =   7
         Top             =   1560
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            TabIndex        =   8
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
            TabIndex        =   9
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
            Left            =   9450
            TabIndex        =   10
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
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
            Left            =   8100
            TabIndex        =   13
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   90
            TabIndex        =   11
            Top             =   450
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   7230
         Left            =   150
         TabIndex        =   14
         Top             =   2370
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   12753
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
            Height          =   7095
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   12515
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "EvaCre_frm_534.frx":1076
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_534.frx":1092
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "EvaCre_frm_534.frx":10AE
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos del Crédito"
            TabPicture(3)   =   "EvaCre_frm_534.frx":10CA
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Ratios"
            TabPicture(4)   =   "EvaCre_frm_534.frx":10E6
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Comentarios EEFF"
            TabPicture(5)   =   "EvaCre_frm_534.frx":1102
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "SSPanel16"
            Tab(5).Control(1)=   "SSPanel15"
            Tab(5).Control(2)=   "SSPanel13"
            Tab(5).Control(3)=   "SSPanel10"
            Tab(5).Control(4)=   "SSPanel9"
            Tab(5).ControlCount=   5
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6645
               Index           =   0
               Left            =   60
               TabIndex        =   16
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11721
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
               Height          =   6435
               Index           =   1
               Left            =   -74940
               TabIndex        =   17
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   11351
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
               Height          =   6435
               Index           =   3
               Left            =   -74940
               TabIndex        =   18
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   11351
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
               Height          =   6435
               Index           =   2
               Left            =   -74940
               TabIndex        =   19
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   11351
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
               Height          =   6435
               Index           =   4
               Left            =   -74940
               TabIndex        =   22
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   11351
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   23
               Top             =   360
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   2196
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSPanel pnl_Coment 
                  Height          =   1095
                  Index           =   0
                  Left            =   2700
                  TabIndex        =   33
                  Top             =   90
                  Width           =   8385
                  _Version        =   65536
                  _ExtentX        =   14790
                  _ExtentY        =   1931
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
               Begin VB.Label Label6 
                  Caption         =   "a) Liquidez (Capital de Trabajo)"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   24
                  Top             =   465
                  Width           =   2205
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   1275
               Left            =   -74940
               TabIndex        =   25
               Top             =   1650
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   2249
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSPanel pnl_Coment 
                  Height          =   1095
                  Index           =   1
                  Left            =   2700
                  TabIndex        =   34
                  Top             =   90
                  Width           =   8385
                  _Version        =   65536
                  _ExtentX        =   14790
                  _ExtentY        =   1931
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
               Begin VB.Label Label8 
                  Caption         =   "b) Solvencia (Endeudamiento)"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   26
                  Top             =   465
                  Width           =   2175
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   27
               Top             =   2970
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   2196
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSPanel pnl_Coment 
                  Height          =   1095
                  Index           =   2
                  Left            =   2700
                  TabIndex        =   35
                  Top             =   90
                  Width           =   8385
                  _Version        =   65536
                  _ExtentX        =   14790
                  _ExtentY        =   1931
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
               Begin VB.Label Label9 
                  Caption         =   "c) Rentabilidad"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   28
                  Top             =   465
                  Width           =   2085
               End
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   29
               Top             =   4260
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   2196
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSPanel pnl_Coment 
                  Height          =   1095
                  Index           =   3
                  Left            =   2700
                  TabIndex        =   36
                  Top             =   90
                  Width           =   8385
                  _Version        =   65536
                  _ExtentX        =   14790
                  _ExtentY        =   1931
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
               Begin VB.Label Label10 
                  Caption         =   "d) Ciclo del Negocio"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   30
                  Top             =   465
                  Width           =   1725
               End
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   1245
               Left            =   -74940
               TabIndex        =   31
               Top             =   5550
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   2196
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSPanel pnl_Coment 
                  Height          =   1095
                  Index           =   4
                  Left            =   2700
                  TabIndex        =   37
                  Top             =   90
                  Width           =   8385
                  _Version        =   65536
                  _ExtentX        =   14790
                  _ExtentY        =   1931
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
               Begin VB.Label Label12 
                  Caption         =   "e) Endeudamiento de Apalancamiento Financiero"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   32
                  Top             =   375
                  Width           =   2025
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_68"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerLab_Click()
    Screen.MousePointer = 11
    frm_EvaCre_70.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub cmd_VerNeg_Click()
    Screen.MousePointer = 11
    frm_EvaCre_71.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub cmd_VerPer_Click()
    Screen.MousePointer = 11
    frm_EvaCre_69.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Me.Caption = modgen_g_str_NomPlt
   Screen.MousePointer = 11
   moddat_g_int_CodIns = 21
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call fs_Limpia
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Información del Inmueble
   Call modmip_gs_DatCre(grd_Listad(3), r_arr_Mtz)                                        'Buscar Información del Crédito
   
   Call fs_DatRat(moddat_g_str_NumSol, grd_Listad(4))
   Call fs_DatEEFF(moddat_g_str_NumSol)
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de Cónyuge
    For r_int_Contad = 0 To 4
       grd_Listad(r_int_Contad).ColWidth(0) = 4200:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter '2900
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       
       Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
    Next r_int_Contad
   
End Sub
Private Sub fs_Limpia()
Dim r_int_Contad  As Integer

   For r_int_Contad = 0 To 4
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
      pnl_Coment.Item(r_int_Contad).Caption = ""
   Next r_int_Contad
End Sub
Public Sub fs_DatRat(ByVal p_NumSol As String, p_Grid As MSFlexGrid)
Dim r_str_Cadena As String
Dim r_str_CadAux As String

    g_str_Parame = "SELECT * FROM MIC_RATFIN WHERE "
    g_str_Parame = g_str_Parame & "RATFIN_NUMSOL = '" & p_NumSol & "' "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        Exit Sub
    End If
   
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
        p_Grid.Redraw = False
        
        g_rst_Princi.MoveFirst
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ratio de Liquidez":                          p_Grid.CellFontBold = True
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Razón Corriente"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_RAZCTE, "###,###,###,##0.000000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Capital de Trabajo"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_CAPTRA, "###,###,###,##0.00")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = ""
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ratios de Solvencia (Endeudamiento)":        p_Grid.CellFontBold = True
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Solvecia Total"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_SOLTOT, "###,###,###,##0.000000000")
        
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = ""
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ratios de Rentabilidad":                     p_Grid.CellFontBold = True
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rentabilidad de Ventas"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_RENVTA, "###,###,###,##0.000000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rentabilidad del Patrimonio - ROE"
        p_Grid.Col = 1:                     p_Grid.Text = g_rst_Princi!RATFIN_RENPAT & "%"
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rentabilidad del Activo - ROA"
        p_Grid.Col = 1:                     p_Grid.Text = g_rst_Princi!RATFIN_RENACT & "%"
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = ""
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ciclo del Negocio":                           p_Grid.CellFontBold = True
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rotación de Ctas. X Cobrar"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_ROTCOB, "###,###,###,##0.000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Periodo Promedio de Cobro"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_PERCOB, "###,###,###,##0.000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rotación de Mercaderias "
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_ROTMER, "###,###,###,##0.000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Rotación de Cuentas por Pagar"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_ROTPAG, "###,###,###,##0.000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = ""
        p_Grid.Col = 1:                     p_Grid.Text = ""
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Otros Ratios":                                p_Grid.CellFontBold = True
        p_Grid.Col = 1:                     p_Grid.Text = ""
   
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ratio cuota /Ingreso  Neto"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_CUOING, "###,###,###,##0.000000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "Ratio cuota/excedente Mensual"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_CUOEXC, "###,###,###,##0.000000000")
        
        p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                     p_Grid.Text = "(e) Apalancamiento Financiero < 1"
        p_Grid.Col = 1:                     p_Grid.Text = Format(g_rst_Princi!RATFIN_APAFIN, "###,###,###,##0.000000000")
        
        p_Grid.Redraw = True
        Call gs_UbiIniGrid(p_Grid)
        
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
End Sub
Public Sub fs_DatEEFF(ByVal p_NumSol As String)
   g_str_Parame = "SELECT * FROM MIC_RATFIN WHERE "
   g_str_Parame = g_str_Parame & "RATFIN_NUMSOL = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       
       g_rst_Princi.MoveFirst
       If Not IsNull(g_rst_Princi!RATFIN_COMLIQ) Then pnl_Coment(0).Caption = Trim(g_rst_Princi!RATFIN_COMLIQ)
       If Not IsNull(g_rst_Princi!RATFIN_COMSOL) Then pnl_Coment(1).Caption = Trim(g_rst_Princi!RATFIN_COMSOL)
       If Not IsNull(g_rst_Princi!RATFIN_COMREN) Then pnl_Coment(2).Caption = Trim(g_rst_Princi!RATFIN_COMREN)
       If Not IsNull(g_rst_Princi!RATFIN_COMCIC) Then pnl_Coment(3).Caption = Trim(g_rst_Princi!RATFIN_COMCIC)
       If Not IsNull(g_rst_Princi!RATFIN_COMEND) Then pnl_Coment(4).Caption = Trim(g_rst_Princi!RATFIN_COMEND)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
