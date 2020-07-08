VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_EvaCre_69 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   Icon            =   "EvaCre_frm_535.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9495
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   16748
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
         Height          =   7155
         Left            =   60
         TabIndex        =   22
         Top             =   2280
         Width           =   11505
         _Version        =   65536
         _ExtentX        =   20294
         _ExtentY        =   12621
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
         Begin TabDlg.SSTab tab_TipCli 
            Height          =   7065
            Left            =   30
            TabIndex        =   23
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   12462
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            Tab             =   4
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Titular"
            TabPicture(0)   =   "EvaCre_frm_535.frx":000C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_535.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "EvaCre_frm_535.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos del Crédito"
            TabPicture(3)   =   "EvaCre_frm_535.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Negocio"
            TabPicture(4)   =   "EvaCre_frm_535.frx":007C
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).Control(0)=   "SSPanel11"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "SSPanel8"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "SSPanel7"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).Control(3)=   "SSPanel9"
            Tab(4).Control(3).Enabled=   0   'False
            Tab(4).Control(4)=   "SSPanel10"
            Tab(4).Control(4).Enabled=   0   'False
            Tab(4).Control(5)=   "SSPanel5"
            Tab(4).Control(5).Enabled=   0   'False
            Tab(4).Control(6)=   "SSPanel3"
            Tab(4).Control(6).Enabled=   0   'False
            Tab(4).ControlCount=   7
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6645
               Index           =   0
               Left            =   -74910
               TabIndex        =   35
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
               Height          =   6615
               Index           =   1
               Left            =   -74910
               TabIndex        =   36
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11668
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   2085
               Left            =   60
               TabIndex        =   37
               Top             =   390
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   3678
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
               Begin VB.TextBox txt_CarInf 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   5
                  Top             =   1710
                  Width           =   9375
               End
               Begin VB.TextBox txt_NomInf 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   4
                  Top             =   1380
                  Width           =   9375
               End
               Begin VB.TextBox txt_RefTit 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   3
                  Top             =   1050
                  Width           =   9375
               End
               Begin VB.TextBox txt_DirTit 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   2
                  Top             =   720
                  Width           =   9375
               End
               Begin VB.TextBox txt_RazSoc 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   0
                  Top             =   60
                  Width           =   9375
               End
               Begin VB.TextBox txt_RucTit 
                  Height          =   315
                  Left            =   1800
                  MaxLength       =   11
                  TabIndex        =   1
                  Top             =   390
                  Width           =   3525
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Cargo del Informante"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   43
                  Top             =   1770
                  Width           =   1470
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "R.U.C"
                  Height          =   195
                  Index           =   5
                  Left            =   60
                  TabIndex        =   42
                  Top             =   450
                  Width           =   435
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Razón Social"
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   41
                  Top             =   120
                  Width           =   945
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Referencia"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   40
                  Top             =   1110
                  Width           =   780
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Informante"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   39
                  Top             =   1440
                  Width           =   750
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Dirección"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   38
                  Top             =   780
                  Width           =   675
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   855
               Left            =   60
               TabIndex        =   44
               Top             =   3120
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1508
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
               Begin VB.TextBox txt_CarMat 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   8
                  Top             =   360
                  Width           =   3945
               End
               Begin VB.TextBox txt_CarFac 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   9
                  Top             =   360
                  Width           =   4335
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Características"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   60
                  TabIndex        =   47
                  Top             =   30
                  Width           =   1305
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Material"
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   46
                  Top             =   420
                  Width           =   555
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Fachada"
                  Height          =   195
                  Index           =   4
                  Left            =   6000
                  TabIndex        =   45
                  Top             =   420
                  Width           =   630
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   495
               Left            =   60
               TabIndex        =   48
               Top             =   4020
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   873
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
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   10
                  Top             =   160
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Luz"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   1
                  Left            =   4050
                  TabIndex        =   11
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Agua"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   2
                  Left            =   6660
                  TabIndex        =   12
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Teléfono"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   3
                  Left            =   9330
                  TabIndex        =   13
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Internet"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Servicios"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   49
                  Top             =   160
                  Width           =   645
               End
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   870
               Left            =   60
               TabIndex        =   50
               Top             =   4560
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1535
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
               Begin VB.TextBox txt_ObsTit 
                  Height          =   715
                  Left            =   1800
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   14
                  Text            =   "EvaCre_frm_535.frx":0098
                  Top             =   60
                  Width           =   9435
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Observaciones"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   51
                  Top             =   320
                  Width           =   1065
               End
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   585
               Left            =   60
               TabIndex        =   52
               Top             =   6405
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1032
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
                  Height          =   285
                  Left            =   10770
                  TabIndex        =   18
                  ToolTipText     =   "Adjuntar Croquis del Negocio"
                  Top             =   150
                  Width           =   435
               End
               Begin VB.CommandButton cmd_CroTit 
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
                  Height          =   285
                  Left            =   10320
                  TabIndex        =   17
                  ToolTipText     =   "Adjuntar Croquis del Negocio"
                  Top             =   150
                  Width           =   435
               End
               Begin Threed.SSPanel pnl_ArcCrq 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   16
                  Top             =   150
                  Width           =   8475
                  _Version        =   65536
                  _ExtentX        =   14949
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Croquis del Negocio"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   53
                  Top             =   210
                  Width           =   1425
               End
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   555
               Left            =   60
               TabIndex        =   54
               Top             =   2520
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   979
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
               Begin VB.TextBox txt_OtrEst 
                  Height          =   315
                  Left            =   6810
                  TabIndex        =   7
                  Top             =   120
                  Width           =   4365
               End
               Begin VB.ComboBox cmb_TipEst 
                  Height          =   315
                  Left            =   1800
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   120
                  Width           =   3945
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Otros"
                  Height          =   195
                  Index           =   8
                  Left            =   6000
                  TabIndex        =   56
                  Top             =   180
                  Width           =   375
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Establecimiento"
                  Height          =   195
                  Index           =   2
                  Left            =   60
                  TabIndex        =   55
                  Top             =   180
                  Width           =   1110
               End
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   885
               Left            =   60
               TabIndex        =   57
               Top             =   5475
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1561
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
               Begin VB.TextBox txt_ResTit 
                  Height          =   735
                  Left            =   1800
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   15
                  Text            =   "EvaCre_frm_535.frx":009F
                  Top             =   60
                  Width           =   9435
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Resumen"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   58
                  Top             =   330
                  Width           =   675
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6615
               Index           =   2
               Left            =   -74910
               TabIndex        =   59
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11668
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
               Height          =   6615
               Index           =   3
               Left            =   -74910
               TabIndex        =   60
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11668
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
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   60
         TabIndex        =   24
         Top             =   1470
         Width           =   11505
         _Version        =   65536
         _ExtentX        =   20294
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
            TabIndex        =   25
            Top             =   390
            Width           =   9975
            _Version        =   65536
            _ExtentX        =   17595
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
            TabIndex        =   26
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
            TabIndex        =   27
            Top             =   60
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
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
            Left            =   7740
            TabIndex        =   30
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud:"
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   120
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   450
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   31
         Top             =   60
         Width           =   11505
         _Version        =   65536
         _ExtentX        =   20294
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
            TabIndex        =   32
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
         End
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   660
            TabIndex        =   33
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Verificaciones del Negocio MicroEmpresario"
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   10050
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "EvaCre_frm_535.frx":00A6
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   34
         Top             =   780
         Width           =   11505
         _Version        =   65536
         _ExtentX        =   20294
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
            Left            =   10890
            Picture         =   "EvaCre_frm_535.frx":03B0
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_535.frx":07F2
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_69"
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

Private Sub cmb_TipEst_Click()
   Call gs_SetFocus(txt_OtrEst)
End Sub

Private Sub cmb_TipEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipEst_Click
   End If
End Sub

Private Sub cmd_CroTit_Click()
  On Error GoTo cmd_BusArc_Error
   
   dlg_Guarda.Filter = "Todos los archivos (*.*)|*.*"
   dlg_Guarda.ShowOpen
   pnl_ArcCrq.Caption = UCase(dlg_Guarda.FileName)
   
   If Trim(pnl_ArcCrq.Caption) <> "" Then
      cmd_VerArc.Enabled = True
   End If
   
   Call gs_SetFocus(cmd_Grabar)
   Exit Sub
   
cmd_BusArc_Error:
   pnl_ArcCrq.Caption = ""
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_Contad     As Integer
Dim r_str_NomArc     As String
Dim r_str_ExtArc     As String
Dim r_str_ArcNue     As String
Dim r_str_CadAux     As String
   
   If Len(Trim(pnl_ArcCrq.Caption)) > 0 Then
      r_str_NomArc = Mid(Trim(pnl_ArcCrq.Caption), InStrRev(pnl_ArcCrq.Caption, "\") + 1)
      r_str_ExtArc = Mid(r_str_NomArc, InStrRev(r_str_NomArc, ".") + 1)
      r_str_CadAux = Mid(r_str_NomArc, 1, InStr(r_str_NomArc, "." & r_str_ExtArc) - 1)
      r_str_ArcNue = moddat_g_str_NumSol & "_Croquis_" & r_str_CadAux & "." & r_str_ExtArc
   End If
   
   If Trim(txt_RazSoc.Text) = "" Then
      MsgBox "Debe ingresar la Razón Social del Negocio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RazSoc)
      Exit Sub
   End If
'   If Trim(txt_RucTit.Text) = "" Then
'      MsgBox "Debe ingresar el RUC del Titular.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_RucTit)
'      Exit Sub
'   End If
   If cmb_TipEst.ListIndex = -1 Then
      MsgBox "Debe ingresar el Tipo de Establecimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEst)
      Exit Sub
   End If
   
'   If Not gf_Valida_RUC(Trim(txt_RucTit.Text), Mid(Trim(txt_RucTit.Text), Len(Trim(txt_RucTit.Text)), 1)) Then
'      MsgBox "El Número de RUC no es válido.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(txt_RucTit)
'      Exit Sub
'   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
       Screen.MousePointer = 11
       g_str_Parame = "USP_MIC_DATNEG ("
       g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
       g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & "', "
       g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
       g_str_Parame = g_str_Parame & "'" & txt_RucTit.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_RazSoc.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_DirTit.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_RefTit.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_NomInf.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_CarInf.Text & "', "
       g_str_Parame = g_str_Parame & "'" & cmb_TipEst.ItemData(cmb_TipEst.ListIndex) & "', "
       g_str_Parame = g_str_Parame & "'" & txt_OtrEst.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_CarMat.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_CarFac.Text & "', "
         
       For r_int_Contad = 0 To 3
         g_str_Parame = g_str_Parame & "'" & IIf(sck_SerVic.Item(r_int_Contad).Value = "Verdadero", 1, 0) & "', "
       Next r_int_Contad
       
       g_str_Parame = g_str_Parame & "'" & txt_ObsTit.Text & "', "
       g_str_Parame = g_str_Parame & "'" & txt_ResTit.Text & "', "
       
       If InStr(Trim(pnl_ArcCrq.Caption), "\") > 0 Then
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_ArcNue) & "', "
       Else
         g_str_Parame = g_str_Parame & "'" & Trim(pnl_ArcCrq.Caption) & "', "
       End If

       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
       g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
       g_str_Parame = g_str_Parame & moddat_g_int_FlgAct & ") "
       
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
      If r_str_ArcNue <> "" And InStr(r_str_ArcNue, "\") > 0 Then
          If gf_Existe_Archivo(moddat_g_str_RutAnx, r_str_ArcNue) Then
              Kill moddat_g_str_RutAnx & "\" & r_str_ArcNue
          End If
          FileCopy Trim(pnl_ArcCrq.Caption), moddat_g_str_RutAnx & "\" & r_str_ArcNue
      End If
   End If
   MsgBox "Se grabó satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
    
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerArc_Click()
   If pnl_ArcCrq.Caption <> "" Then
      If InStr(Trim(pnl_ArcCrq.Caption), "\") > 0 Then
         ShellExecute Me.hwnd, "Open", pnl_ArcCrq.Caption, "", "", 1
      Else
         ShellExecute Me.hwnd, "Open", moddat_g_str_RutAnx & "\" & pnl_ArcCrq.Caption, "", "", 1
      End If
   End If
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Información del Inmueble
   Call modmip_gs_DatCre(grd_Listad(3), r_arr_Mtz)
   
   Call fs_Activa(True)
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub
Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT DATNEG_NUMSOL, DATNEG_TIPDOC, DATNEG_NUMDOC, DATNEG_RUCNEG, DATNEG_RAZSOC, DATNEG_DIRNEG, DATNEG_REFERE, "
   g_str_Parame = g_str_Parame & "         DATNEG_NOMINF, DATNEG_CARINF, DATNEG_TIPLOC, DATNEG_OTREST, DATNEG_CARACT_MAT, DATNEG_CARACT_FAC, DATNEG_FLGSER_LUZ, "
   g_str_Parame = g_str_Parame & "         DATNEG_FLGSER_AGU, DATNEG_FLGSER_TEL, DATNEG_FLGSER_INT, DATNEG_ARCCRQ, DATNEG_OBSERV, DATNEG_RESUME "
   g_str_Parame = g_str_Parame & "    FROM MIC_DATNEG WHERE DATNEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        moddat_g_int_FlgAct = 1
        Exit Sub
    End If
    moddat_g_int_FlgAct = 2
    g_rst_Princi.MoveFirst
    
    If Not IsNull(g_rst_Princi!DATNEG_RAZSOC) Then
         txt_RazSoc.Text = g_rst_Princi!DATNEG_RAZSOC
    End If
    If Not IsNull(g_rst_Princi!DATNEG_RUCNEG) Then
        txt_RucTit.Text = g_rst_Princi!DATNEG_RUCNEG
    End If
    If Not IsNull(g_rst_Princi!DATNEG_DIRNEG) Then
        txt_DirTit.Text = g_rst_Princi!DATNEG_DIRNEG
    End If
    If Not IsNull(g_rst_Princi!DATNEG_REFERE) Then
        txt_RefTit.Text = g_rst_Princi!DATNEG_REFERE
    End If
    If Not IsNull(g_rst_Princi!DATNEG_NOMINF) Then
        txt_NomInf.Text = g_rst_Princi!DATNEG_NOMINF
    End If
    If Not IsNull(g_rst_Princi!DATNEG_CARINF) Then
        txt_CarInf.Text = g_rst_Princi!DATNEG_CARINF
    End If
    If Not IsNull(g_rst_Princi!DATNEG_TIPLOC) Then
        cmb_TipEst.Text = moddat_gf_Consulta_ParDes("208", g_rst_Princi!DATNEG_TIPLOC)
    End If
    If Not IsNull(g_rst_Princi!DATNEG_OTREST) Then
        txt_OtrEst.Text = g_rst_Princi!DATNEG_OTREST
    End If
    If Not IsNull(g_rst_Princi!DATNEG_CARACT_MAT) Then
        txt_CarMat.Text = g_rst_Princi!DATNEG_CARACT_MAT
    End If
    If Not IsNull(g_rst_Princi!DATNEG_CARACT_FAC) Then
        txt_CarFac.Text = g_rst_Princi!DATNEG_CARACT_FAC
    End If
    If Not IsNull(g_rst_Princi!DATNEG_FLGSER_LUZ) Then
        sck_SerVic.Item(0).Value = g_rst_Princi!DATNEG_FLGSER_LUZ
    End If
    If Not IsNull(g_rst_Princi!DATNEG_FLGSER_AGU) Then
        sck_SerVic.Item(1) = g_rst_Princi!DATNEG_FLGSER_AGU
    End If
    If Not IsNull(g_rst_Princi!DATNEG_FLGSER_TEL) Then
        sck_SerVic.Item(2) = g_rst_Princi!DATNEG_FLGSER_TEL
    End If
    If Not IsNull(g_rst_Princi!DATNEG_FLGSER_INT) Then
        sck_SerVic.Item(3) = g_rst_Princi!DATNEG_FLGSER_INT
    End If
    If Not IsNull(g_rst_Princi!DATNEG_OBSERV) Then
        txt_ObsTit.Text = g_rst_Princi!DATNEG_OBSERV
    End If
    If Not IsNull(g_rst_Princi!DATNEG_RESUME) Then
        txt_ResTit.Text = g_rst_Princi!DATNEG_RESUME
    End If
    If Not IsNull(g_rst_Princi!DATNEG_ARCCRQ) Then
      pnl_ArcCrq.Caption = g_rst_Princi!DATNEG_ARCCRQ
      cmd_VerArc.Enabled = True
    End If
End Sub
Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipEst, 1, "208")
   'Inicializando Grid de Cliente y de Cónyuge
    For r_int_Contad = 0 To 3
       grd_Listad(r_int_Contad).ColWidth(0) = 2900:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       
       Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
    Next r_int_Contad
   
   cmd_VerArc.Enabled = False
End Sub
Private Sub fs_Limpia()
Dim r_int_Contad     As Integer
   
   txt_RazSoc.Text = ""
   txt_RucTit.Text = ""
   txt_DirTit.Text = ""
   txt_RefTit.Text = ""
   txt_NomInf.Text = ""
   txt_CarInf.Text = ""
   cmb_TipEst.ListIndex = -1
   txt_OtrEst.Text = ""
   txt_CarMat.Text = ""
   txt_CarFac.Text = ""
   
   For r_int_Contad = 0 To 3
      sck_SerVic.Item(r_int_Contad).Value = False
   Next r_int_Contad
   
   txt_ObsTit.Text = ""
   pnl_ArcCrq.Caption = ""
   txt_ResTit.Text = ""
End Sub
Private Sub fs_Activa(ByVal p_Habilita As Integer)
Dim r_int_Contad  As Integer

    txt_RazSoc.Enabled = p_Habilita
    txt_RucTit.Enabled = p_Habilita
    txt_DirTit.Enabled = p_Habilita
    txt_RefTit.Enabled = p_Habilita
    txt_NomInf.Enabled = p_Habilita
    txt_CarInf.Enabled = p_Habilita
    cmb_TipEst.Enabled = p_Habilita
    txt_OtrEst.Enabled = p_Habilita
    txt_CarMat.Enabled = p_Habilita
    txt_CarFac.Enabled = p_Habilita
    
    For r_int_Contad = 0 To 3
        sck_SerVic.Item(r_int_Contad).Enabled = p_Habilita
    Next r_int_Contad
    
    txt_ObsTit.Enabled = p_Habilita
    pnl_ArcCrq.Enabled = p_Habilita
    cmd_CroTit.Enabled = p_Habilita
    txt_ResTit.Enabled = p_Habilita

End Sub

Private Sub sck_SerVic_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      Select Case Index
         Case 0: Call gs_SetFocus(sck_SerVic(1))
         Case 1: Call gs_SetFocus(sck_SerVic(2))
         Case 2: Call gs_SetFocus(sck_SerVic(3))
         Case 3: Call gs_SetFocus(txt_ObsTit)
      End Select
   End If
End Sub


Private Sub txt_CarFac_GotFocus()
   Call gs_SelecTodo(txt_CarFac)
End Sub

Private Sub txt_CarFac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(sck_SerVic(0))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_CarInf_GotFocus()
   Call gs_SelecTodo(txt_CarInf)
End Sub

Private Sub txt_CarInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipEst)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_CarMat_GotFocus()
   Call gs_SelecTodo(txt_CarMat)
End Sub

Private Sub txt_CarMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CarFac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_DirTit_GotFocus()
   Call gs_SelecTodo(txt_DirTit)
End Sub

Private Sub txt_DirTit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RefTit)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_NomInf_GotFocus()
   Call gs_SelecTodo(txt_NomInf)
End Sub

Private Sub txt_NomInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CarInf)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_ObsTit_GotFocus()
   Call gs_SelecTodo(txt_ObsTit)
End Sub

Private Sub txt_ObsTit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ResTit)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_OtrEst_GotFocus()
   Call gs_SelecTodo(txt_OtrEst)
End Sub

Private Sub txt_OtrEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CarMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RucTit)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,@#$%&;:()/º")
   End If
End Sub

Private Sub txt_RefTit_GotFocus()
   Call gs_SelecTodo(txt_RefTit)
End Sub

Private Sub txt_RefTit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomInf)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_ResTit_GotFocus()
   Call gs_SelecTodo(txt_ResTit)
End Sub

Private Sub txt_ResTit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_CroTit)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_RucTit_GotFocus()
   Call gs_SelecTodo(txt_RucTit)
End Sub

Private Sub txt_RucTit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirTit)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub
