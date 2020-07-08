VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_EvaCre_66 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   7065
   ClientLeft      =   1740
   ClientTop       =   2715
   ClientWidth     =   11235
   Icon            =   "EvaCre_frm_073.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7065
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   12462
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
         Height          =   4755
         Left            =   30
         TabIndex        =   9
         Top             =   2250
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   8387
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
            Height          =   4635
            Left            =   30
            TabIndex        =   10
            Top             =   60
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   8176
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Titular"
            TabPicture(0)   =   "EvaCre_frm_073.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel3"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_073.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel5"
            Tab(1).Control(1)=   "SSPanel13"
            Tab(1).ControlCount=   2
            Begin Threed.SSPanel SSPanel3 
               Height          =   2085
               Left            =   30
               TabIndex        =   11
               Top             =   360
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
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
               Begin VB.CommandButton cmd_Tit_EvaPri 
                  Caption         =   "..."
                  Height          =   345
                  Left            =   10530
                  TabIndex        =   60
                  ToolTipText     =   "Evaluación de Empresas Empleadoras"
                  Top             =   1380
                  Width           =   405
               End
               Begin VB.TextBox txt_Tit_ObsPri 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   0
                  Text            =   "EvaCre_frm_073.frx":0044
                  Top             =   60
                  Width           =   8595
               End
               Begin VB.ComboBox cmb_Tit_TipPri 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   720
                  Width           =   8595
               End
               Begin Threed.SSPanel pnl_Tit_EmpPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   27
                  Top             =   1380
                  Width           =   8145
                  _Version        =   65536
                  _ExtentX        =   14367
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
               Begin Threed.SSPanel pnl_Tit_CalPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   29
                  Top             =   1710
                  Width           =   2985
                  _Version        =   65536
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "TOP"
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
               Begin Threed.SSPanel pnl_Tit_CodPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   52
                  Top             =   1050
                  Width           =   8595
                  _Version        =   65536
                  _ExtentX        =   15161
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
               Begin VB.Label Label11 
                  Caption         =   "Actividad Principal:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   53
                  Top             =   1050
                  Width           =   1935
               End
               Begin VB.Label Label3 
                  Caption         =   "Calificación de Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   30
                  Top             =   1710
                  Width           =   2145
               End
               Begin VB.Label Label5 
                  Caption         =   "Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   28
                  Top             =   1380
                  Width           =   1935
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Verificación Act. Eco. Principal:"
                  Height          =   495
                  Index           =   0
                  Left            =   60
                  TabIndex        =   26
                  Top             =   60
                  Width           =   2175
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Verif. Act. Eco. Princ.:"
                  Height          =   315
                  Index           =   5
                  Left            =   60
                  TabIndex        =   12
                  Top             =   720
                  Width           =   2115
               End
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   2085
               Left            =   30
               TabIndex        =   31
               Top             =   2490
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
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
               Begin VB.CommandButton cmd_Tit_EvaSec 
                  Caption         =   "..."
                  Height          =   345
                  Left            =   10530
                  TabIndex        =   61
                  ToolTipText     =   "Evaluación de Empresas Empleadoras"
                  Top             =   1380
                  Width           =   405
               End
               Begin VB.ComboBox cmb_Tit_TipSec 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   3
                  Top             =   720
                  Width           =   8595
               End
               Begin VB.TextBox txt_Tit_ObsSec 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   2
                  Text            =   "EvaCre_frm_073.frx":0048
                  Top             =   60
                  Width           =   8595
               End
               Begin Threed.SSPanel pnl_Tit_EmpSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   32
                  Top             =   1380
                  Width           =   8145
                  _Version        =   65536
                  _ExtentX        =   14367
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
               Begin Threed.SSPanel pnl_Tit_CalSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   33
                  Top             =   1710
                  Width           =   2985
                  _Version        =   65536
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "TOP"
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
               Begin Threed.SSPanel pnl_Tit_CodSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   54
                  Top             =   1050
                  Width           =   8595
                  _Version        =   65536
                  _ExtentX        =   15161
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
                  Caption         =   "Actividad Secundaria:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   55
                  Top             =   1050
                  Width           =   1935
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Verif. Act. Eco. Secund.:"
                  Height          =   315
                  Index           =   3
                  Left            =   60
                  TabIndex        =   37
                  Top             =   720
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Verificación Act. Eco. Secundaria:"
                  Height          =   495
                  Index           =   2
                  Left            =   60
                  TabIndex        =   36
                  Top             =   60
                  Width           =   2175
               End
               Begin VB.Label Label6 
                  Caption         =   "Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   35
                  Top             =   1380
                  Width           =   1935
               End
               Begin VB.Label Label4 
                  Caption         =   "Calificación de Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   34
                  Top             =   1710
                  Width           =   2145
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   2085
               Left            =   -74970
               TabIndex        =   38
               Top             =   360
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
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
               Begin VB.CommandButton cmd_Cyg_EvaPri 
                  Caption         =   "..."
                  Height          =   345
                  Left            =   10530
                  TabIndex        =   62
                  ToolTipText     =   "Evaluación de Empresas Empleadoras"
                  Top             =   1380
                  Width           =   405
               End
               Begin VB.ComboBox cmb_Cyg_TipPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   5
                  Text            =   "Combo1"
                  Top             =   720
                  Width           =   8595
               End
               Begin VB.TextBox txt_Cyg_ObsPri 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   4
                  Text            =   "EvaCre_frm_073.frx":004C
                  Top             =   60
                  Width           =   8595
               End
               Begin Threed.SSPanel pnl_Cyg_EmpPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   39
                  Top             =   1380
                  Width           =   8145
                  _Version        =   65536
                  _ExtentX        =   14367
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
               Begin Threed.SSPanel pnl_Cyg_CalPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   40
                  Top             =   1710
                  Width           =   2985
                  _Version        =   65536
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "TOP"
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
               Begin Threed.SSPanel pnl_Cyg_CodPri 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   58
                  Top             =   1050
                  Width           =   8595
                  _Version        =   65536
                  _ExtentX        =   15161
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "pnl_Cyg_CodPri"
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
               Begin VB.Label Label14 
                  Caption         =   "Actividad Principal:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   59
                  Top             =   1050
                  Width           =   1935
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Verif. Act. Eco. Princ.:"
                  Height          =   315
                  Index           =   4
                  Left            =   60
                  TabIndex        =   44
                  Top             =   720
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Verificación Act. Eco. Principal:"
                  Height          =   495
                  Index           =   1
                  Left            =   60
                  TabIndex        =   43
                  Top             =   60
                  Width           =   2175
               End
               Begin VB.Label Label8 
                  Caption         =   "Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   42
                  Top             =   1380
                  Width           =   1935
               End
               Begin VB.Label Label7 
                  Caption         =   "Calificación de Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   41
                  Top             =   1710
                  Width           =   2145
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   2085
               Left            =   -74970
               TabIndex        =   45
               Top             =   2490
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
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
               Begin VB.CommandButton cmd_Cyg_EvaSec 
                  Caption         =   "..."
                  Height          =   345
                  Left            =   10530
                  TabIndex        =   63
                  ToolTipText     =   "Evaluación de Empresas Empleadoras"
                  Top             =   1380
                  Width           =   405
               End
               Begin VB.TextBox txt_Cyg_ObsSec 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   6
                  Text            =   "EvaCre_frm_073.frx":0050
                  Top             =   60
                  Width           =   8595
               End
               Begin VB.ComboBox cmb_Cyg_TipSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   7
                  Text            =   "Combo1"
                  Top             =   720
                  Width           =   8595
               End
               Begin Threed.SSPanel pnl_Cyg_EmpSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   46
                  Top             =   1380
                  Width           =   8145
                  _Version        =   65536
                  _ExtentX        =   14367
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "EMPRESA XXXX"
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
               Begin Threed.SSPanel pnl_Cyg_CalSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   47
                  Top             =   1710
                  Width           =   2985
                  _Version        =   65536
                  _ExtentX        =   5265
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "TOP"
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
               Begin Threed.SSPanel pnl_Cyg_CodSec 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   56
                  Top             =   1050
                  Width           =   8595
                  _Version        =   65536
                  _ExtentX        =   15161
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "pnl_Cyg_CodSec"
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
               Begin VB.Label Label13 
                  Caption         =   "Actividad Secundaria:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   57
                  Top             =   1050
                  Width           =   1935
               End
               Begin VB.Label Label10 
                  Caption         =   "Calificación de Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   51
                  Top             =   1710
                  Width           =   2145
               End
               Begin VB.Label Label9 
                  Caption         =   "Empleador:"
                  Height          =   315
                  Left            =   60
                  TabIndex        =   50
                  Top             =   1380
                  Width           =   1935
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Verificación Act. Eco. Secundaria:"
                  Height          =   495
                  Index           =   7
                  Left            =   60
                  TabIndex        =   49
                  Top             =   60
                  Width           =   2175
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Verif. Act. Eco. Secund.:"
                  Height          =   315
                  Index           =   6
                  Left            =   60
                  TabIndex        =   48
                  Top             =   720
                  Width           =   2115
               End
            End
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            TabIndex        =   14
            Top             =   390
            Width           =   9675
            _Version        =   65536
            _ExtentX        =   17066
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   15
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         End
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9090
            TabIndex        =   16
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   7710
            TabIndex        =   17
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   660
            TabIndex        =   22
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Verificaciones Laborales"
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
            Picture         =   "EvaCre_frm_073.frx":0054
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   23
         Top             =   750
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_073.frx":035E
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "EvaCre_frm_073.frx":07A0
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_DatCyg           As Integer

Dim l_int_Tit_ActPri       As Integer
Dim l_int_Tit_ActSec       As Integer
Dim l_int_Cyg_ActPri       As Integer
Dim l_int_Cyg_ActSec       As Integer

Dim l_int_Tit_TDoPri       As Integer
Dim l_str_Tit_NDoPri       As String
Dim l_int_Tit_TDoSec       As Integer
Dim l_str_Tit_NDoSec       As String
Dim l_int_Cyg_TDoPri       As Integer
Dim l_str_Cyg_NDoPri       As String
Dim l_int_Cyg_TDoSec       As Integer
Dim l_str_Cyg_NDoSec       As String

Dim l_int_Tit_CalPri       As Integer
Dim l_int_Tit_CalSec       As Integer
Dim l_int_Cyg_CalPri       As Integer
Dim l_int_Cyg_CalSec       As Integer

Dim l_int_Tit_CodPri       As Integer
Dim l_int_Tit_CodSec       As Integer
Dim l_int_Cyg_CodPri       As Integer
Dim l_int_Cyg_CodSec       As Integer

Private Sub cmb_Cyg_TipPri_Click()
   If l_int_Cyg_ActSec = 2 Then
      Call gs_SetFocus(txt_Tit_ObsSec)
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_Cyg_TipPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_TipPri_Click
   End If
End Sub

Private Sub cmb_Cyg_TipSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_TipSec_Click
   End If
End Sub

Private Sub cmb_Tit_TipPri_Click()
   If l_int_Tit_ActSec = 2 Then
      Call gs_SetFocus(txt_Tit_ObsSec)
   Else
      If l_int_DatCyg = 2 Then
         If l_int_Cyg_ActPri = 1 Then
            tab_TipCli.Tab = 1
            Call gs_SetFocus(txt_Cyg_ObsPri)
         End If
      Else
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_Cyg_TipSec_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Tit_TipPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_TipPri_Click
   End If
End Sub

Private Sub cmb_Tit_TipSec_Click()
   If l_int_DatCyg = 2 Then
      If l_int_Cyg_ActPri = 1 Then
         tab_TipCli.Tab = 1
         Call gs_SetFocus(txt_Cyg_ObsPri)
      End If
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_Tit_TipSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_TipSec_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgGrb = 1 Then
      MsgBox "Primero debe de registrar las verificaciones personales en la evaluación crediticia.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Len(Trim(txt_Tit_ObsPri.Text)) = 0 Then
      MsgBox "Debe ingresar los Comentarios sobre la Verificación de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Tit_ObsPri)
      Exit Sub
   End If
   
   If cmb_Tit_TipPri.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Verificación de Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tit_TipPri)
      Exit Sub
   End If
   
   If txt_Tit_ObsSec.Enabled Then
      If Len(Trim(txt_Tit_ObsSec.Text)) = 0 Then
         MsgBox "Debe ingresar los Comentarios sobre la Verificación de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Tit_ObsSec)
         Exit Sub
      End If
      
      If cmb_Tit_TipSec.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Verificación de Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Tit_TipSec)
         Exit Sub
      End If
   End If
   
   If txt_Cyg_ObsPri.Enabled Then
      If Len(Trim(txt_Cyg_ObsPri.Text)) = 0 Then
         MsgBox "Debe ingresar los Comentarios sobre la Verificación de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         
         tab_TipCli.Tab = 1
         Call gs_SetFocus(txt_Cyg_ObsPri)
         Exit Sub
      End If
      
      If cmb_Cyg_TipPri.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Verificación de Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         
         tab_TipCli.Tab = 1
         Call gs_SetFocus(cmb_Cyg_TipPri)
         Exit Sub
      End If
   End If
   
   If txt_Cyg_ObsSec.Enabled Then
      If Len(Trim(txt_Cyg_ObsSec.Text)) = 0 Then
         MsgBox "Debe ingresar los Comentarios sobre la Verificación de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         
         tab_TipCli.Tab = 1
         Call gs_SetFocus(txt_Cyg_ObsSec)
         Exit Sub
      End If
      
      If cmb_Cyg_TipSec.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Verificación de Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         
         tab_TipCli.Tab = 1
         Call gs_SetFocus(cmb_Cyg_TipSec)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   Call moddat_gs_FecSis
   
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "USP_TRA_EVACRE_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "0, "       'Ingreso 1
      g_str_Parame = g_str_Parame & "0, "       'Ingreso 2
      g_str_Parame = g_str_Parame & "0, "       'Ingreso 3
      g_str_Parame = g_str_Parame & "0, "       'Ingreso 4
      g_str_Parame = g_str_Parame & "0, "       'Cuota Soles
      g_str_Parame = g_str_Parame & "0, "       'Cuota Moneda Préstamo
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Cambio
      g_str_Parame = g_str_Parame & "0, "       'Flag Condicion
      g_str_Parame = g_str_Parame & "'', "      'Observaciones de Evaluación
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Adicional 1
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Adicional 2
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Total
      g_str_Parame = g_str_Parame & "0, "       'Obligaciones Mensuales
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Neto
      g_str_Parame = g_str_Parame & "0, "       'Monto Préstamo Aprobado
      g_str_Parame = g_str_Parame & "0, "       'Plazo Préstamo Aprobado
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Seguro Aprobado
      g_str_Parame = g_str_Parame & "0, "       'Período de Gracia Aprobado
      g_str_Parame = g_str_Parame & "0, "       'Cuotas dobles Aprobado
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Cambio (Para Calificación de Ingresos)
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Calificación de Ingresos
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Calificación de Condiciones de Crédito (Aprobación)
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda
      
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Verificación Domiciliaria
      g_str_Parame = g_str_Parame & "'', "      'Observación de Verificación Domiciliaria
      g_str_Parame = g_str_Parame & "0, "       'Flag de Central de Riesgos
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Reporte de Riesgos
            
      g_str_Parame = g_str_Parame & "0, "       'Nro. de Entidades de Central de Riesgo
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 0
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 1
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 2
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 3
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 4
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Moneda Nacional
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Moneda Extranjera
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 1
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 1
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 1
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 2
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 2
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 2
      
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 3
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 3
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 3
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 4
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 4
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 4
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 5
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 5
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 5
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 6
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 6
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 6
      
      g_str_Parame = g_str_Parame & "'', "      'Comentarios Central de Riesgos
      
      g_str_Parame = g_str_Parame & "'" & txt_Tit_ObsPri.Text & "',"                                     'Verificación Laboral 1
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipPri.ItemData(cmb_Tit_TipPri.ListIndex)) & ", "       'Tipo de Verificación Laboral 1
      g_str_Parame = g_str_Parame & CStr(l_int_Tit_CalPri) & ", "                                        'Clasificación Empleador 1
      
      If txt_Tit_ObsSec.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Tit_ObsSec.Text & "',"                                  'Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipSec.ItemData(cmb_Tit_TipSec.ListIndex)) & ", "    'Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(l_int_Tit_CalSec) & ", "                                     'Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Titular - Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Titular- Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Titular - Clasificación Empleador 2
      End If
      
      g_str_Parame = g_str_Parame & "0, "       'Flag de Central de Riesgos (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Reporte de Riesgos (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Nro. de Entidades de Central de Riesgo (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 0 (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 1 (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 2 (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 3 (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Clasificación 4 (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Moneda Nacional (Cónyuge)
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Moneda Extranjera (Cónyuge)
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 1
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 1
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 1
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 2
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 2
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 2
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 3
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 3
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 3
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 4
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 4
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 4
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 5
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 5
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 5
   
      g_str_Parame = g_str_Parame & "'', "      'Deuda Entidad 6
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda Entidad 6
      g_str_Parame = g_str_Parame & "0, "       'Clasificación Entidad 6
      
      g_str_Parame = g_str_Parame & "'', "      'Comentarios Central de Riesgos (Cónyuge)
   
      If txt_Cyg_ObsPri.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_ObsPri.Text & "',"                                  'Cónyuge - Verificación Laboral 1
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipPri.ItemData(cmb_Cyg_TipPri.ListIndex)) & ", "    'Cónyuge - Tipo de Verificación Laboral 1
         g_str_Parame = g_str_Parame & CStr(l_int_Cyg_CalPri) & ", "                                     'Cónyuge Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Cónyuge - Verificación Laboral 1
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Tipo de Verificación Laboral 1
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Clasificación Empleador 1
      End If
   
      If txt_Cyg_ObsSec.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_ObsSec.Text & "',"                                  'Cónyuge - Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipSec.ItemData(cmb_Cyg_TipSec.ListIndex)) & ", "    'Cónyuge - Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(l_int_Cyg_CalSec) & ", "                                     'Cónyuge - Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Cónyuge - Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Clasificación Empleador 2
      End If
      
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Ingreso 1
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Ingreso 2
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Ingreso 3
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Ingreso 4
      
      g_str_Parame = g_str_Parame & "0, "       'Obligaciones Mensuales 1
      g_str_Parame = g_str_Parame & "0, "       'Obligaciones Mensuales 2
      
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Titular
      g_str_Parame = g_str_Parame & "0, "       'Total Deuda Cónyuge
      
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Neto Titular
      g_str_Parame = g_str_Parame & "0, "       'Ingreso Neto Cónyuge
      
      g_str_Parame = g_str_Parame & "0, "       'Ratio Ingreso Deuda
      g_str_Parame = g_str_Parame & "0, "       'Ratio Inicial Deuda
      
      g_str_Parame = g_str_Parame & "'', "      'Referencias Personales
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   Else
      g_str_Parame = "USP_TRA_EVACRE_ACT_VERLAB ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Tit_ObsPri.Text & "',"                                        'Titular - Verificación Laboral 1
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipPri.ItemData(cmb_Tit_TipPri.ListIndex)) & ", "          'Titular - Tipo de Verificación Laboral 1
      g_str_Parame = g_str_Parame & CStr(l_int_Tit_CalPri) & ", "                                           'Titular - Clasificación Empleador 1
      
      If txt_Tit_ObsSec.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Tit_ObsSec.Text & "',"                                     'Titular - Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipSec.ItemData(cmb_Tit_TipSec.ListIndex)) & ", "       'Titular - Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(l_int_Tit_CalSec) & ", "                                        'Titular - Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Titular - Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Titular- Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Titular - Clasificación Empleador 2
      End If
      
      If txt_Cyg_ObsPri.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_ObsPri.Text & "',"                                  'Cónyuge - Verificación Laboral 1
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipPri.ItemData(cmb_Cyg_TipPri.ListIndex)) & ", "    'Cónyuge - Tipo de Verificación Laboral 1
         g_str_Parame = g_str_Parame & CStr(l_int_Cyg_CalPri) & ", "                                     'Cónyuge Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Cónyuge - Verificación Laboral 1
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Tipo de Verificación Laboral 1
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Clasificación Empleador 1
      End If
   
      If txt_Cyg_ObsSec.Enabled Then
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_ObsSec.Text & "',"                                  'Cónyuge - Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipSec.ItemData(cmb_Cyg_TipSec.ListIndex)) & ", "    'Cónyuge - Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & CStr(l_int_Cyg_CalSec) & ", "                                     'Cónyuge - Clasificación Empleador 2
      Else
         g_str_Parame = g_str_Parame & "'',"       'Cónyuge - Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Tipo de Verificación Laboral 2
         g_str_Parame = g_str_Parame & "0, "       'Cónyuge - Clasificación Empleador 2
      End If
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar procedimiento USP_TRA_EVACRE_INSERTA / USP_TRA_EVACRE_ACT_VERLAB.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 41, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 0
   
   MsgBox "La información fue registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   
   moddat_g_int_FlgAct_2 = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call fs_Limpia
   
   'Buscar Datos del Cliente (Estado Civil)
   l_int_DatCyg = 1
   
   l_int_Tit_ActPri = 1
   l_int_Tit_ActSec = 1
   l_int_Tit_CodPri = 0
   l_int_Tit_CodSec = 0
   
   l_int_Cyg_ActPri = 1
   l_int_Cyg_ActSec = 1
   l_int_Cyg_CodPri = 0
   l_int_Cyg_CodSec = 0
   
   tab_TipCli.TabVisible(1) = False
   
   cmd_Tit_EvaPri.Enabled = False
   
   txt_Tit_ObsSec.Enabled = False
   cmb_Tit_TipSec.Enabled = False
   cmd_Tit_EvaSec.Enabled = False
   
   txt_Cyg_ObsPri.Enabled = False
   cmb_Cyg_TipPri.Enabled = False
   cmd_Cyg_EvaPri.Enabled = False
   
   txt_Cyg_ObsSec.Enabled = False
   cmb_Cyg_TipSec.Enabled = False
   cmd_Cyg_EvaSec.Enabled = False
   
   l_int_Tit_TDoPri = 0:   l_str_Tit_NDoPri = ""
   l_int_Tit_TDoSec = 0:   l_str_Tit_NDoSec = ""
   
   l_int_Cyg_TDoPri = 0:   l_str_Cyg_NDoPri = ""
   l_int_Cyg_TDoSec = 0:   l_str_Cyg_NDoSec = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If (g_rst_Princi!DATGEN_ESTCIV = 2 And g_rst_Princi!DATGEN_REGCYG = 1) Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
         tab_TipCli.TabVisible(1) = True
         
         l_int_DatCyg = 2
         
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para obtener Actividades Económicas del Cliente Titular
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!ActEco_OrdAct = 1 Then
            l_int_Tit_ActPri = 2
            l_int_Tit_CodPri = g_rst_Princi!ACTECO_CODACT
            
            pnl_Tit_CodPri.Caption = moddat_gf_Consulta_ParDes("008", CStr(l_int_Tit_CodPri))
         Else
            l_int_Tit_ActSec = 2
            l_int_Tit_CodSec = g_rst_Princi!ACTECO_CODACT
            
            pnl_Tit_CodSec.Caption = moddat_gf_Consulta_ParDes("008", CStr(l_int_Tit_CodSec))
            
            txt_Tit_ObsSec.Enabled = True
            cmb_Tit_TipSec.Enabled = True
         End If
            
         Select Case g_rst_Princi!ACTECO_CODACT
            Case 11
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  l_int_Tit_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                  l_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  
                  cmd_Tit_EvaPri.Enabled = True
               Else
                  l_int_Tit_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                  l_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  
                  cmd_Tit_EvaSec.Enabled = True
               End If
               
            Case 21
               If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     l_int_Tit_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     l_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     
                     cmd_Tit_EvaPri.Enabled = True
                  Else
                     l_int_Tit_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     l_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     
                     cmd_Tit_EvaSec.Enabled = True
                  End If
               End If
               
            Case 41
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  l_int_Tit_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                  l_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  
                  cmd_Tit_EvaPri.Enabled = True
               Else
                  l_int_Tit_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                  l_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  
                  cmd_Tit_EvaSec.Enabled = True
               End If
         End Select
            
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If l_int_DatCyg = 2 Then
      'Para obtener Actividades Económicas del Cliente Cónyuge
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' "
      g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!ActEco_OrdAct = 1 Then
               l_int_Cyg_ActPri = 2
               l_int_Cyg_CodPri = g_rst_Princi!ACTECO_CODACT
               
               pnl_Cyg_CodPri.Caption = moddat_gf_Consulta_ParDes("008", CStr(l_int_Cyg_CodPri))
               
               txt_Cyg_ObsPri.Enabled = True
               cmb_Cyg_TipPri.Enabled = True
            Else
               l_int_Cyg_ActSec = 2
               l_int_Cyg_CodSec = g_rst_Princi!ACTECO_CODACT
               
               pnl_Cyg_CodSec.Caption = moddat_gf_Consulta_ParDes("008", CStr(l_int_Cyg_CodSec))
               
               txt_Cyg_ObsSec.Enabled = True
               cmb_Cyg_TipSec.Enabled = True
               cmd_Cyg_EvaSec.Enabled = True
            End If
               
            Select Case g_rst_Princi!ACTECO_CODACT
               Case 11
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     l_int_Cyg_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                     l_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                     
                     cmd_Cyg_EvaPri.Enabled = True
                  Else
                     l_int_Cyg_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                     l_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                     
                     cmd_Cyg_EvaSec.Enabled = True
                  End If
                  
               Case 21
                  If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                     If g_rst_Princi!ActEco_OrdAct = 1 Then
                        l_int_Cyg_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        l_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                        
                        cmd_Cyg_EvaPri.Enabled = True
                     Else
                        l_int_Cyg_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        l_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                        
                        cmd_Cyg_EvaSec.Enabled = True
                     End If
                  End If
                  
               Case 41
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     l_int_Cyg_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                     l_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                     
                     cmd_Cyg_EvaPri.Enabled = True
                  Else
                     l_int_Cyg_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                     l_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                     
                     cmd_Cyg_EvaSec.Enabled = True
                  End If
            End Select
               
            g_rst_Princi.MoveNext
         Loop
      End If
   End If
   
   'Buscando Calificación de Empresas (Actividad Principal - Titular)
   pnl_Tit_EmpPri.Caption = ""
   pnl_Tit_CalPri.Caption = ""
   
   l_int_Tit_CalPri = 0
   
   If l_int_Tit_TDoPri > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(l_int_Tit_TDoPri) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & l_str_Tit_NDoPri & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         pnl_Tit_EmpPri.Caption = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         pnl_Tit_CalPri.Caption = moddat_gf_Consulta_ParDes("016", CStr(g_rst_Princi!DATGEN_CLASIF))
         
         l_int_Tit_CalPri = g_rst_Princi!DATGEN_CLASIF
         
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Buscando Calificación de Empresas (Actividad Secundaria - Titular)
   pnl_Tit_EmpSec.Caption = ""
   pnl_Tit_CalSec.Caption = ""
   
   l_int_Tit_CalSec = 0
   
   If l_int_Tit_TDoSec > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(l_int_Tit_TDoSec) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & l_str_Tit_NDoSec & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         pnl_Tit_EmpSec.Caption = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         pnl_Tit_CalSec.Caption = moddat_gf_Consulta_ParDes("016", CStr(g_rst_Princi!DATGEN_CLASIF))
         
         l_int_Tit_CalSec = g_rst_Princi!DATGEN_CLASIF
         
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Buscando Calificación de Empresas (Actividad Principal - Cónyuge)
   pnl_Cyg_EmpPri.Caption = ""
   pnl_Cyg_CalPri.Caption = ""
   
   l_int_Cyg_CalPri = 0
   
   If l_int_Cyg_TDoPri > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(l_int_Cyg_TDoPri) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & l_str_Cyg_NDoPri & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         pnl_Cyg_EmpPri.Caption = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         pnl_Cyg_CalPri.Caption = moddat_gf_Consulta_ParDes("016", CStr(g_rst_Princi!DATGEN_CLASIF))
         
         l_int_Cyg_CalPri = g_rst_Princi!DATGEN_CLASIF
         
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Buscando Calificación de Empresas (Actividad Secundaria - Cónyuge)
   pnl_Cyg_EmpSec.Caption = ""
   pnl_Cyg_CalSec.Caption = ""
   
   l_int_Cyg_CalSec = 0
   
   If l_int_Cyg_TDoSec > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(l_int_Cyg_TDoSec) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & l_str_Cyg_NDoSec & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         pnl_Cyg_EmpSec.Caption = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         pnl_Cyg_CalSec.Caption = moddat_gf_Consulta_ParDes("016", CStr(g_rst_Princi!DATGEN_CLASIF))
         
         l_int_Cyg_CalSec = g_rst_Princi!DATGEN_CLASIF
         
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Buscar Información de Evaluación ya registrada
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      moddat_g_int_FlgGrb = 1
   Else
      moddat_g_int_FlgGrb = 2
   
      If Len(Trim(g_rst_Princi!EVACRE_TIT_LABVE1 & "")) > 0 Then
         txt_Tit_ObsPri.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE1 & "")
         Call gs_BuscarCombo_Item(cmb_Tit_TipPri, g_rst_Princi!EVACRE_TIT_TIPVE1)
      
         If txt_Tit_ObsSec.Enabled Then
            txt_Tit_ObsSec.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE2 & "")
            Call gs_BuscarCombo_Item(cmb_Tit_TipSec, g_rst_Princi!EVACRE_TIT_TIPVE2)
         End If
      
         If txt_Cyg_ObsPri.Enabled Then
            txt_Cyg_ObsPri.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE1 & "")
            Call gs_BuscarCombo_Item(cmb_Cyg_TipPri, g_rst_Princi!EVACRE_CYG_TIPVE1)
         End If
      
         If txt_Cyg_ObsSec.Enabled Then
            txt_Cyg_ObsSec.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE2 & "")
            Call gs_BuscarCombo_Item(cmb_Cyg_TipSec, g_rst_Princi!EVACRE_CYG_TIPVE2)
         End If
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipPri, 1, "068")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipSec, 1, "068")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_TipPri, 1, "068")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_TipSec, 1, "068")
End Sub

Private Sub fs_Limpia()
   txt_Tit_ObsPri.Text = ""
   cmb_Tit_TipPri.ListIndex = -1
   pnl_Tit_CodPri.Caption = ""
   pnl_Tit_EmpPri.Caption = ""
   pnl_Tit_CalPri.Caption = ""
   
   txt_Tit_ObsSec.Text = ""
   cmb_Tit_TipSec.ListIndex = -1
   pnl_Tit_CodSec.Caption = ""
   pnl_Tit_EmpSec.Caption = ""
   pnl_Tit_CalSec.Caption = ""
   
   txt_Cyg_ObsPri.Text = ""
   cmb_Cyg_TipPri.ListIndex = -1
   pnl_Cyg_CodPri.Caption = ""
   pnl_Cyg_EmpPri.Caption = ""
   pnl_Cyg_CalPri.Caption = ""
   
   txt_Cyg_ObsSec.Text = ""
   cmb_Cyg_TipSec.ListIndex = -1
   pnl_Cyg_CodSec.Caption = ""
   pnl_Cyg_EmpSec.Caption = ""
   pnl_Cyg_CalSec.Caption = ""
End Sub

Private Sub txt_Cyg_ObsPri_GotFocus()
   Call gs_SelecTodo(txt_Cyg_ObsPri)
End Sub

Private Sub txt_Cyg_ObsPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_TipPri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Cyg_ObsSec_GotFocus()
   Call gs_SelecTodo(txt_Cyg_ObsSec)
End Sub

Private Sub txt_Cyg_ObsSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_TipSec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Tit_ObsPri_GotFocus()
   Call gs_SelecTodo(txt_Tit_ObsPri)
End Sub

Private Sub txt_Tit_ObsPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_TipPri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Tit_ObsSec_GotFocus()
   Call gs_SelecTodo(txt_Tit_ObsSec)
End Sub

Private Sub txt_Tit_ObsSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_TipSec)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub


