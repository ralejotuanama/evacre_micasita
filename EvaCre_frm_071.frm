VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_64 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   10410
   ClientLeft      =   3375
   ClientTop       =   495
   ClientWidth     =   11250
   Icon            =   "EvaCre_frm_071.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10410
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   18362
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   1095
         Left            =   30
         TabIndex        =   17
         Top             =   8430
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin Threed.SSPanel pnl_Tot_IngLiq 
            Height          =   315
            Left            =   2880
            TabIndex        =   18
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_OblMen 
            Height          =   315
            Left            =   2880
            TabIndex        =   56
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_IngNet 
            Height          =   315
            Left            =   2880
            TabIndex        =   59
            Top             =   720
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_TotDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   62
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_IngDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   96
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "9.99 veces "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_IniDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   98
            Top             =   720
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "999.99 % "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   9930
            TabIndex        =   100
            Top             =   720
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "% "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Endeudam. / Cuota Inicial:"
            Height          =   315
            Index           =   45
            Left            =   6030
            TabIndex        =   99
            Top             =   720
            Width           =   2325
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ingreso / Total Endeudamiento:"
            Height          =   315
            Index           =   44
            Left            =   6030
            TabIndex        =   97
            Top             =   390
            Width           =   2325
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Endeudamiento:"
            Height          =   315
            Index           =   18
            Left            =   6030
            TabIndex        =   64
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   17
            Left            =   7980
            TabIndex        =   63
            Top             =   90
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Ingreso Neto:"
            Height          =   315
            Index           =   16
            Left            =   60
            TabIndex        =   61
            Top             =   720
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   15
            Left            =   2370
            TabIndex        =   60
            Top             =   750
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Obligaciones Mensuales:"
            Height          =   315
            Index           =   14
            Left            =   60
            TabIndex        =   58
            Top             =   390
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   12
            Left            =   2370
            TabIndex        =   57
            Top             =   420
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Ingreso Líquido:"
            Height          =   315
            Index           =   20
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   21
            Left            =   2370
            TabIndex        =   19
            Top             =   90
            Width           =   465
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   7710
            TabIndex        =   27
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   795
         Left            =   30
         TabIndex        =   28
         Top             =   9570
         WhatsThisHelpID =   11175
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   1402
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
         Begin Threed.SSPanel pnl_CuoMPr 
            Height          =   315
            Left            =   2880
            TabIndex        =   29
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   8520
            TabIndex        =   30
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin EditLib.fpDoubleSingle ipp_CuoMen 
            Height          =   315
            Left            =   2880
            TabIndex        =   12
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label8 
            Caption         =   "C. Mensual Aprob.:"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   2085
         End
         Begin VB.Label Label4 
            Caption         =   "C. Mensual Aprob. (M. Prest.):"
            Height          =   315
            Left            =   60
            TabIndex        =   35
            Top             =   390
            Width           =   2235
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Cambio:"
            Height          =   315
            Left            =   6030
            TabIndex        =   34
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   25
            Left            =   2370
            TabIndex        =   33
            Top             =   90
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   26
            Left            =   7980
            TabIndex        =   32
            Top             =   60
            Width           =   465
         End
         Begin VB.Label lbl_MonPre 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Left            =   2370
            TabIndex        =   31
            Top             =   420
            Width           =   465
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   37
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "EvaCre_frm_071.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_071.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Calcul 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_071.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Actualizar Cálculo"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2085
         Left            =   30
         TabIndex        =   38
         Top             =   4170
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.ComboBox cmb_Tit_TipSec 
            Height          =   315
            Left            =   8490
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   2625
         End
         Begin VB.ComboBox cmb_Tit_TipPri 
            Height          =   315
            Left            =   8490
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2625
         End
         Begin EditLib.fpDoubleSingle ipp_Tit_IngPri 
            Height          =   315
            Left            =   2880
            TabIndex        =   0
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_Tit_IngSec 
            Height          =   315
            Left            =   2880
            TabIndex        =   2
            Top             =   390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_Tit_IngAdi 
            Height          =   315
            Left            =   2880
            TabIndex        =   4
            Top             =   720
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_Tit_OblMen 
            Height          =   315
            Left            =   2880
            TabIndex        =   5
            Top             =   1380
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_Tit_IngLiq 
            Height          =   315
            Left            =   2880
            TabIndex        =   50
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tit_IngNet 
            Height          =   315
            Left            =   2880
            TabIndex        =   53
            Top             =   1710
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tit_TotDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   69
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tit_IngDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   94
            Top             =   1380
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "9.99 "
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
            Alignment       =   4
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ingreso / Total Endeudamiento:"
            Height          =   315
            Index           =   43
            Left            =   6030
            TabIndex        =   95
            Top             =   1380
            Width           =   2325
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Endeudamiento:"
            Height          =   315
            Index           =   23
            Left            =   6030
            TabIndex        =   71
            Top             =   1050
            Width           =   1635
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   22
            Left            =   7980
            TabIndex        =   70
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Ingreso (Act. Secund.):"
            Height          =   315
            Index           =   7
            Left            =   6030
            TabIndex        =   68
            Top             =   390
            Width           =   2235
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Ingreso (Act. Princ.):"
            Height          =   315
            Index           =   6
            Left            =   6030
            TabIndex        =   65
            Top             =   60
            Width           =   2145
         End
         Begin VB.Label Label3 
            Caption         =   "Ingreso Neto (Titular):"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   1710
            Width           =   2115
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   8
            Left            =   2370
            TabIndex        =   54
            Top             =   1740
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Líquido Total (Titular):"
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   52
            Top             =   1050
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   3
            Left            =   2370
            TabIndex        =   51
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   2
            Left            =   2370
            TabIndex        =   49
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Obligaciones Mensuales:"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   48
            Top             =   1380
            Width           =   1965
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Act. Econ. Princ. (Titular):"
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   9
            Left            =   2370
            TabIndex        =   43
            Top             =   90
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Act. Econ. Sec. (Titular):"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   42
            Top             =   390
            Width           =   2355
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   10
            Left            =   2370
            TabIndex        =   41
            Top             =   420
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Otros Ingresos (Titular):"
            Height          =   315
            Index           =   11
            Left            =   60
            TabIndex        =   40
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   13
            Left            =   2370
            TabIndex        =   39
            Top             =   750
            Width           =   465
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   45
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
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   330
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Calificación de Ingresos"
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
            Picture         =   "EvaCre_frm_071.frx":0BA2
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   1875
         Left            =   30
         TabIndex        =   66
         Top             =   2250
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   3307
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
            Height          =   1755
            Left            =   60
            TabIndex        =   67
            Top             =   60
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   3096
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
      Begin Threed.SSPanel SSPanel14 
         Height          =   2085
         Left            =   30
         TabIndex        =   72
         Top             =   6300
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin EditLib.fpDoubleSingle ipp_Cyg_IngAdi 
            Height          =   315
            Left            =   2880
            TabIndex        =   10
            Top             =   720
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_Cyg_IngSec 
            Height          =   315
            Left            =   2880
            TabIndex        =   8
            Top             =   390
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_Cyg_IngPri 
            Height          =   315
            Left            =   2880
            TabIndex        =   6
            Top             =   60
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.ComboBox cmb_Cyg_TipPri 
            Height          =   315
            Left            =   8490
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   60
            Width           =   2625
         End
         Begin VB.ComboBox cmb_Cyg_TipSec 
            Height          =   315
            Left            =   8490
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   390
            Width           =   2625
         End
         Begin EditLib.fpDoubleSingle ipp_Cyg_OblMen 
            Height          =   315
            Left            =   2880
            TabIndex        =   11
            Top             =   1380
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_Cyg_IngLiq 
            Height          =   315
            Left            =   2880
            TabIndex        =   73
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cyg_IngNet 
            Height          =   315
            Left            =   2880
            TabIndex        =   74
            Top             =   1710
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cyg_TotDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   75
            Top             =   1050
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cyg_IngDeu 
            Height          =   315
            Left            =   8490
            TabIndex        =   92
            Top             =   1380
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "9.99 veces "
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
            Alignment       =   4
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ingreso / Total Endeudamiento:"
            Height          =   315
            Index           =   42
            Left            =   6030
            TabIndex        =   93
            Top             =   1380
            Width           =   2325
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   41
            Left            =   2370
            TabIndex        =   91
            Top             =   750
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Otros Ingresos (Cónyuge):"
            Height          =   315
            Index           =   40
            Left            =   60
            TabIndex        =   90
            Top             =   720
            Width           =   2085
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   39
            Left            =   2370
            TabIndex        =   89
            Top             =   420
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Act. Econ. Sec. (Cónyuge):"
            Height          =   315
            Index           =   38
            Left            =   60
            TabIndex        =   88
            Top             =   390
            Width           =   2355
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   37
            Left            =   2370
            TabIndex        =   87
            Top             =   90
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Act. Econ. Princ. (Cónyuge):"
            Height          =   315
            Index           =   36
            Left            =   60
            TabIndex        =   86
            Top             =   60
            Width           =   2415
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Obligaciones Mensuales:"
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   85
            Top             =   1380
            Width           =   1965
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   34
            Left            =   2370
            TabIndex        =   84
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   33
            Left            =   2370
            TabIndex        =   83
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Ing. Líquido Total (Cónyuge):"
            Height          =   315
            Index           =   32
            Left            =   60
            TabIndex        =   82
            Top             =   1050
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   31
            Left            =   2370
            TabIndex        =   81
            Top             =   1740
            Width           =   465
         End
         Begin VB.Label Label5 
            Caption         =   "Ingreso Neto (Cónyuge):"
            Height          =   315
            Left            =   60
            TabIndex        =   80
            Top             =   1710
            Width           =   2115
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Ingreso (Act. Princ.):"
            Height          =   315
            Index           =   30
            Left            =   6030
            TabIndex        =   79
            Top             =   60
            Width           =   2145
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Ingreso (Act. Secund.):"
            Height          =   315
            Index           =   29
            Left            =   6030
            TabIndex        =   78
            Top             =   390
            Width           =   2235
         End
         Begin VB.Label lbl_Etique 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   27
            Left            =   7980
            TabIndex        =   77
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total Endeudamiento:"
            Height          =   315
            Index           =   24
            Left            =   6030
            TabIndex        =   76
            Top             =   1050
            Width           =   1635
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_Arr_TNC_Cli()  As String
Dim l_Arr_TC_Cli()   As String
Dim l_Arr_TNC_Cof()  As String
Dim l_Arr_TC_Cof()   As String
Dim l_dbl_CuoRta     As Double
Dim l_dbl_ComVta     As Double
Dim l_dbl_ApoPro     As Double
Dim l_dbl_MtoPre     As Double
Dim l_dbl_TasInt     As Double
Dim l_dbl_MtoAho     As Double
Dim l_dbl_TipCam     As Double
Dim l_int_DatCyg     As Integer
Dim l_int_TipEva     As Integer
Dim l_int_PlaAno     As Integer
Dim l_int_PerGra     As Integer
Dim l_int_CuoDbl     As Integer
Dim l_int_TipSeg     As Integer
Dim l_int_DiaPag     As Integer
Dim l_int_MonAho     As Integer
Dim l_str_EmpSeg     As String
Dim l_str_CodCiu     As String
Dim l_int_TasEsp     As Integer

Private Sub cmd_Calcul_Click()
   Call fs_CalRat
   
   Select Case l_int_TipEva
      Case 1
         ipp_CuoMen.Enabled = False
         ipp_CuoMen.Value = Format(l_dbl_CuoRta / 100 * CDbl(pnl_Tot_IngNet.Caption), "###,##0.00") & " "
         
         'Recalculando Cuota en Moneda de Préstamo
         Call fs_Calcul_CuoMPr
         
      Case 2
         'Recalculando Cuota en Moneda de Préstamo
         Call fs_Calcul_CuoMPr
         
         'Calculando Ratios
         Call fs_CalRat
   End Select
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_Ini_IniDeu    As Double
Dim r_dbl_IngMin        As Double
   
   If ipp_Tit_IngPri.Enabled Then
      If ipp_Tit_IngPri.Value = 0# Then
         MsgBox "Debe ingresar el Monto de Ingresos de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Tit_IngPri)
         Exit Sub
      End If
   End If
   If cmb_Tit_TipPri.Enabled Then
      If cmb_Tit_TipPri.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Ingresos de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Tit_TipPri)
         Exit Sub
      End If
   End If
   If ipp_Tit_IngSec.Enabled Then
      If ipp_Tit_IngSec.Value = 0# Then
         MsgBox "Debe ingresar el Monto de Ingresos de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Tit_IngSec)
         Exit Sub
      End If
   End If
   If cmb_Tit_TipSec.Enabled Then
      If cmb_Tit_TipSec.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Ingresos de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Tit_TipSec)
         Exit Sub
      End If
   End If
   If CDbl(pnl_Tit_TotDeu.Caption) > 0# Then
      If ipp_Tit_OblMen.Enabled Then
         If ipp_Tit_OblMen.Value = 0 Then
            MsgBox "Debe ingresar el Monto de Obligaciones Mensuales.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_OblMen)
            Exit Sub
         End If
      End If
   End If
   If ipp_Cyg_IngPri.Enabled Then
      If ipp_Cyg_IngPri.Value = 0# Then
         MsgBox "Debe ingresar el Monto de Ingresos de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Cyg_IngPri)
         Exit Sub
      End If
   End If
   If cmb_Cyg_TipPri.Enabled Then
      If cmb_Cyg_TipPri.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Ingresos de la Actividad Económica Principal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Cyg_TipPri)
         Exit Sub
      End If
   End If
   If ipp_Cyg_IngSec.Enabled Then
      If ipp_Cyg_IngSec.Value = 0# Then
         MsgBox "Debe ingresar el Monto de Ingresos de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Cyg_IngSec)
         Exit Sub
      End If
   End If
   If cmb_Cyg_TipSec.Enabled Then
      If cmb_Cyg_TipSec.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Ingresos de la Actividad Económica Secundaria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Cyg_TipSec)
         Exit Sub
      End If
   End If
   If CDbl(pnl_Cyg_TotDeu.Caption) > 0# Then
      If ipp_Cyg_OblMen.Enabled Then
         If ipp_Cyg_OblMen.Value = 0 Then
            MsgBox "Debe ingresar el Monto de Obligaciones Mensuales.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Cyg_OblMen)
            Exit Sub
         End If
      End If
   End If
   
   If CDbl(pnl_Tit_IngNet.Caption) < 0 Then
      MsgBox "El ingreso neto del titular no puede ser negativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Tit_OblMen)
      Exit Sub
   End If

   If CDbl(pnl_Cyg_IngNet.Caption) < 0 Then
      MsgBox "El ingreso neto del conyuge no puede ser negativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Cyg_OblMen)
      Exit Sub
   End If
   
   If CDbl(ipp_CuoMen.Value) = 0 Then
      MsgBox "Debe ingresar la Cuota Mensual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_CuoMen)
      Exit Sub
   End If
   If CDbl(pnl_CuoMPr.Caption) = 0 Then
      MsgBox "Debe recalcular la Cuota máxima a prestar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Calcul)
      Exit Sub
   End If
   
   If l_int_TipEva = 1 Then
      r_dbl_IngMin = 0
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "014") Then
         r_dbl_IngMin = moddat_g_arr_Genera(1).Genera_Cantid
      End If
   
      If CDbl(pnl_Tot_IngNet.Caption) < r_dbl_IngMin Then
         MsgBox "El Ingreso Neto es menor al Ingreso Mínimo solicitado para el Producto. (Mínimo permitido S/. " & Format(r_dbl_IngMin, "###,##0.00") & ")", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   If l_int_TipEva = 3 Or l_int_TipEva = 4 Then
      'Obteniendo Parámetro de Relación Deuda / Cuota Inicial
      r_dbl_Ini_IniDeu = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "003") Then
         r_dbl_Ini_IniDeu = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      If CDbl(pnl_Tot_IniDeu.Caption) > r_dbl_Ini_IniDeu Then
         MsgBox "La Relación Total Deuda / Cuota Inicial no se ajusta al Parámetro requerido. (Máximo permitido: " & r_dbl_Ini_IniDeu & "%)", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   Call moddat_gs_FecSis
   
   g_str_Parame = "USP_TRA_EVACRE_ACT_CALING ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
   g_str_Parame = g_str_Parame & CStr(ipp_Tit_IngPri.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Tit_IngSec.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Cyg_IngPri.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Cyg_IngSec.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_CuoMen.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMPr.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_TipCam) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Tit_IngAdi.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Cyg_IngAdi.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_IngLiq.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_OblMen.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_IngNet.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_TotDeu.Caption)) & ", "
   
   If cmb_Tit_TipPri.ListIndex > -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipPri.ItemData(cmb_Tit_TipPri.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   If cmb_Tit_TipSec.ListIndex > -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_TipSec.ItemData(cmb_Tit_TipSec.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   If cmb_Cyg_TipPri.ListIndex > -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipPri.ItemData(cmb_Cyg_TipPri.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   If cmb_Cyg_TipSec.ListIndex > -1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_Cyg_TipSec.ItemData(cmb_Cyg_TipSec.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   g_str_Parame = g_str_Parame & CStr(ipp_Tit_OblMen.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_Cyg_OblMen.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tit_TotDeu.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Cyg_TotDeu.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tit_IngNet.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Cyg_IngNet.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_IngDeu.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_IniDeu.Caption)) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar procedimiento USP_TRA_EVACRE_ACT_CALING.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 19, 0, "", 0, 0) Then
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
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call fs_Limpia
   
   ipp_Tit_IngPri.Enabled = False:  cmb_Tit_TipPri.Enabled = False
   ipp_Tit_IngSec.Enabled = False:  cmb_Tit_TipSec.Enabled = False
   ipp_Cyg_IngPri.Enabled = False:  cmb_Cyg_TipPri.Enabled = False
   ipp_Cyg_IngSec.Enabled = False:  cmb_Cyg_TipSec.Enabled = False
   ipp_Cyg_IngAdi.Enabled = False
   ipp_Cyg_OblMen.Enabled = False
   l_int_DatCyg = 1
   
   'Para obtener Estado Civil de Cliente
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      If (g_rst_Princi!DATGEN_ESTCIV = 2 And g_rst_Princi!DATGEN_REGCYG = 1) Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
         l_int_DatCyg = 2
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
      
      l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
      If CInt(l_str_CodCiu) = 0 Then
         MsgBox "El codigo CIIU del cliente debe estar registrado, favor de coordinarlo con Comercial.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscar datos del Cónyuge
   If l_int_DatCyg = 2 Then
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         If CInt(l_str_CodCiu) <> 7522 And CInt(l_str_CodCiu) <> 7523 Then
            l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
         End If
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Para obtener Actividades Económicas del Cliente Titular
   g_str_Parame = "SELECT COUNT(*) AS NUMACT FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      If g_rst_Princi!NUMACT = 1 Then
         ipp_Tit_IngPri.Enabled = True:   cmb_Tit_TipPri.Enabled = True
      ElseIf g_rst_Princi!NUMACT = 2 Then
         ipp_Tit_IngPri.Enabled = True:   cmb_Tit_TipPri.Enabled = True
         ipp_Tit_IngSec.Enabled = True:   cmb_Tit_TipSec.Enabled = True
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para obtener Actividades Económicas del Cliente Cónyuge
   If l_int_DatCyg = 2 Then
      g_str_Parame = "SELECT COUNT(*) AS NUMACT FROM CLI_ACTECO WHERE "
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' "
      g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         If g_rst_Princi!NUMACT = 1 Then
            ipp_Cyg_IngPri.Enabled = True:   cmb_Cyg_TipPri.Enabled = True
         ElseIf g_rst_Princi!NUMACT = 2 Then
            ipp_Cyg_IngPri.Enabled = True:   cmb_Cyg_TipPri.Enabled = True
            ipp_Cyg_IngSec.Enabled = True:   cmb_Cyg_TipSec.Enabled = True
         End If
         
         ipp_Cyg_IngAdi.Enabled = True
         ipp_Cyg_OblMen.Enabled = True
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Cargando Información de Solicitud de Crédito
   Call fs_DatCre
   Call modmip_gs_DatCre(grd_Listad, r_arr_Mtz)
   
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
      pnl_Tit_TotDeu.Caption = Format(g_rst_Princi!EVACRE_TIT_TOTDMN + g_rst_Princi!EVACRE_TIT_TOTDME, "#,###,##0.00") & " "
      pnl_Cyg_TotDeu.Caption = Format(g_rst_Princi!EVACRE_CYG_TOTDMN + g_rst_Princi!EVACRE_CYG_TOTDME, "#,###,##0.00") & " "
      
      If l_int_TipEva = 1 Then
         'Titular
         ipp_Tit_IngPri.Value = g_rst_Princi!EVACRE_INGDP1
         If Len(Trim(g_rst_Princi!EVACRE_TIPIDP1 & "")) > 0 Then
            Call gs_BuscarCombo_Item(cmb_Tit_TipPri, g_rst_Princi!EVACRE_TIPIDP1)
         End If
         If g_rst_Princi!EVACRE_TIPIDP2 > 0 Then
            ipp_Tit_IngSec.Value = g_rst_Princi!EVACRE_INGDP2
            Call gs_BuscarCombo_Item(cmb_Tit_TipSec, g_rst_Princi!EVACRE_TIPIDP2)
         End If
         ipp_Tit_IngAdi.Value = g_rst_Princi!EVACRE_INGAD1
         If Len(Trim(g_rst_Princi!EVACRE_OMETIT & "")) > 0 Then
            ipp_Tit_OblMen.Value = g_rst_Princi!EVACRE_OMETIT
         End If
         
         'Cónyuge
         If g_rst_Princi!EVACRE_TIPIDP3 > 0 Then
            ipp_Cyg_IngPri.Value = g_rst_Princi!EVACRE_INGDP3
            Call gs_BuscarCombo_Item(cmb_Cyg_TipPri, g_rst_Princi!EVACRE_TIPIDP3)
         End If
         If g_rst_Princi!EVACRE_TIPIDP4 > 0 Then
            ipp_Cyg_IngSec.Value = g_rst_Princi!EVACRE_INGDP4
            Call gs_BuscarCombo_Item(cmb_Cyg_TipSec, g_rst_Princi!EVACRE_TIPIDP4)
         End If
         ipp_Cyg_IngAdi.Value = g_rst_Princi!EVACRE_INGAD2
         If Len(Trim(g_rst_Princi!EVACRE_OMECYG & "")) > 0 Then
            ipp_Cyg_OblMen.Value = g_rst_Princi!EVACRE_OMECYG
         End If
         ipp_CuoMen.Value = g_rst_Princi!EVACRE_CUOSOL
         
      ElseIf l_int_TipEva = 2 Then
         ipp_CuoMen.Value = g_rst_Princi!EVACRE_CUOSOL
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Select Case l_int_TipEva
      Case 1
         'Calculando Ratios
         Call fs_CalRat
         ipp_CuoMen.Enabled = False
         ipp_CuoMen.Value = Format(l_dbl_CuoRta / 100 * CDbl(pnl_Tot_IngNet.Caption), "###,##0.00") & " "
         
         'Recalculando Cuota en Moneda de Préstamo
         Call fs_Calcul_CuoMPr
      
      Case 2
         ipp_Tit_IngPri.Enabled = False
         ipp_Tit_IngSec.Enabled = False
         ipp_Tit_IngAdi.Enabled = False
         cmb_Tit_TipPri.Enabled = False
         cmb_Tit_TipSec.Enabled = False
         ipp_Tit_OblMen.Enabled = False
         ipp_Cyg_IngPri.Enabled = False
         ipp_Cyg_IngSec.Enabled = False
         ipp_Cyg_IngAdi.Enabled = False
         cmb_Cyg_TipPri.Enabled = False
         cmb_Cyg_TipSec.Enabled = False
         ipp_Cyg_OblMen.Enabled = False
         If ipp_CuoMen.Value = 0 Then
            If l_int_MonAho = 1 Then
               ipp_CuoMen.Value = l_dbl_MtoAho
            Else
               ipp_CuoMen.Value = CDbl(Format(l_dbl_MtoAho * l_dbl_TipCam, "###,###,##0.00"))
            End If
         End If
         
         'Recalculando Cuota en Moneda de Préstamo
         Call fs_Calcul_CuoMPr
         
         'Calculando Ratios
         Call fs_CalRat
         
      Case 3, 4
         ipp_Tit_IngPri.Enabled = False
         ipp_Tit_IngSec.Enabled = False
         ipp_Tit_IngAdi.Enabled = False
         cmb_Tit_TipPri.Enabled = False
         cmb_Tit_TipSec.Enabled = False
         ipp_Tit_OblMen.Enabled = False
         ipp_Cyg_IngPri.Enabled = False
         ipp_Cyg_IngSec.Enabled = False
         ipp_Cyg_IngAdi.Enabled = False
         cmb_Cyg_TipPri.Enabled = False
         cmb_Cyg_TipSec.Enabled = False
         ipp_Cyg_OblMen.Enabled = False
         ipp_CuoMen.Enabled = False
         
         Call fs_CalCuo
         
         'Recalculando Cuota en Moneda de Préstamo
         Call fs_Calcul_CuoMPr
         
         'Calculando Ratios
         Call fs_CalRat
         
   End Select
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   l_dbl_CuoRta = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "013") Then
      l_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   grd_Listad.ColWidth(0) = 3000
   grd_Listad.ColWidth(1) = 7940
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipPri, 1, "069")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_TipSec, 1, "069")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_TipPri, 1, "069")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_TipSec, 1, "069")
End Sub

Private Sub fs_Limpia()
   'Titular
   ipp_Tit_IngPri.Value = 0
   ipp_Tit_IngSec.Value = 0
   ipp_Tit_IngAdi.Value = 0
   cmb_Tit_TipPri.ListIndex = -1
   cmb_Tit_TipSec.ListIndex = -1
   pnl_Tit_IngLiq.Caption = "0.00 "
   ipp_Tit_OblMen.Value = 0
   pnl_Tit_TotDeu.Caption = "0.00 "
   pnl_Tit_IngNet.Caption = "0.00 "

   'Cónyuge
   ipp_Cyg_IngPri.Value = 0
   ipp_Cyg_IngSec.Value = 0
   ipp_Cyg_IngAdi.Value = 0
   cmb_Cyg_TipPri.ListIndex = -1
   cmb_Cyg_TipSec.ListIndex = -1
   pnl_Cyg_IngLiq.Caption = "0.00 "
   ipp_Cyg_OblMen.Value = 0
   pnl_Cyg_TotDeu.Caption = "0.00 "
   pnl_Cyg_IngNet.Caption = "0.00 "

   'Total
   pnl_Tot_IngLiq.Caption = "0.00 "
   pnl_Tot_TotDeu.Caption = "0.00 "
   pnl_Tot_OblMen.Caption = "0.00 "
   pnl_Tot_IngDeu.Caption = "0 "
   pnl_Tot_IngNet.Caption = "0.00 "
   pnl_Tot_IniDeu.Caption = "0.00 "
   
   ipp_CuoMen.Value = 0
   pnl_TipCam.Caption = "0.00 "
   pnl_CuoMPr.Caption = "0.00 "
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_CuoMen_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_CuoMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Calcul)
   End If
End Sub

Private Sub ipp_Cyg_IngAdi_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Cyg_IngPri_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Cyg_IngSec_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Cyg_OblMen_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Tit_IngAdi_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Tit_IngPri_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Tit_IngPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_TipPri)
   End If
End Sub

Private Sub cmb_Tit_TipPri_Click()
   If ipp_Tit_IngSec.Enabled Then
      Call gs_SetFocus(ipp_Tit_IngSec)
   Else
      Call gs_SetFocus(ipp_Tit_IngAdi)
   End If
End Sub

Private Sub cmb_Tit_TipPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_TipPri_Click
   End If
End Sub

Private Sub ipp_Tit_IngSec_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Tit_IngSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_TipSec)
   End If
End Sub

Private Sub cmb_Tit_TipSec_Click()
   Call gs_SetFocus(ipp_Tit_IngAdi)
End Sub

Private Sub cmb_Tit_TipSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_TipSec_Click
   End If
End Sub

Private Sub ipp_Tit_IngAdi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_OblMen)
   End If
End Sub

Private Sub ipp_Tit_OblMen_Change()
   Call fs_Limpia_CuoMPr
End Sub

Private Sub ipp_Tit_OblMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_Cyg_IngPri.Enabled Then
         Call gs_SetFocus(ipp_Cyg_IngPri)
      Else
         If ipp_CuoMen.Enabled Then
            Call gs_SetFocus(ipp_CuoMen)
         Else
            Call gs_SetFocus(cmd_Calcul)
         End If
      End If
   End If
End Sub

Private Sub ipp_Cyg_IngPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_TipPri)
   End If
End Sub

Private Sub cmb_Cyg_TipPri_Click()
   If ipp_Cyg_IngSec.Enabled Then
      Call gs_SetFocus(ipp_Cyg_IngSec)
   Else
      Call gs_SetFocus(ipp_Cyg_IngAdi)
   End If
End Sub

Private Sub cmb_Cyg_TipPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_TipPri_Click
   End If
End Sub

Private Sub ipp_Cyg_IngSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_TipSec)
   End If
End Sub

Private Sub cmb_Cyg_TipSec_Click()
   Call gs_SetFocus(ipp_Cyg_IngAdi)
End Sub

Private Sub cmb_Cyg_TipSec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_TipSec_Click
   End If
End Sub

Private Sub ipp_Cyg_IngAdi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_OblMen)
   End If
End Sub

Private Sub ipp_Cyg_OblMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_CuoMen.Enabled Then
         Call gs_SetFocus(ipp_CuoMen)
      Else
         Call gs_SetFocus(cmd_Calcul)
      End If
   End If
End Sub

Private Sub fs_DatCre()
Dim r_dbl_TipCam     As Double
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If moddat_g_int_TipMon = 1 Then
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
      l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_SOL - g_rst_Princi!SOLMAE_MTOGCI
   Else
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
      l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_DOL - g_rst_Princi!SOLMAE_MTOGCI
   End If
   
   l_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
   l_int_TipEva = g_rst_Princi!SOLMAE_TIPEVA
   l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
   l_int_PlaAno = g_rst_Princi!SOLMAE_PLAANO
   l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
   l_int_CuoDbl = g_rst_Princi!SOLMAE_CUOEXT
   l_str_EmpSeg = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
   l_int_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
   l_int_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
   l_int_MonAho = g_rst_Princi!SOLMAE_MONAHO
   l_dbl_MtoAho = g_rst_Princi!SOLMAE_MTOAHO
   l_int_TasEsp = g_rst_Princi!SOLMAE_TASESP
   lbl_MonPre.Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Tipo de Cambio de Moneda del Préstamo
   l_dbl_TipCam = 0
   
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   End If
End Sub

Private Sub fs_Limpia_CuoMPr()
   pnl_CuoMPr.Caption = "0.00 "
End Sub

Private Sub fs_CalRat()
   Dim r_dbl_IniSol     As Double
   
   'Titular
   pnl_Tit_IngLiq.Caption = Format(CDbl(ipp_Tit_IngPri.Value) + CDbl(ipp_Tit_IngSec.Value) + CDbl(ipp_Tit_IngAdi.Value), "###,##0.00") & " "
   pnl_Tit_IngNet.Caption = Format(CDbl(pnl_Tit_IngLiq.Caption) - CDbl(ipp_Tit_OblMen.Value), "###,##0.00") & " "
      
   If CDbl(pnl_Tit_IngNet.Caption) > 0# Then
      pnl_Tit_IngDeu.Caption = Format(CDbl(pnl_Tit_TotDeu.Caption) / CDbl(pnl_Tit_IngNet.Caption), "##0.00") & " "
   Else
      pnl_Tit_IngDeu.Caption = "0.00 "
   End If
   
   'Cónyuge
   pnl_Cyg_IngLiq.Caption = Format(CDbl(ipp_Cyg_IngPri.Value) + CDbl(ipp_Cyg_IngSec.Value) + CDbl(ipp_Cyg_IngAdi.Value), "###,##0.00") & " "
   If CDbl(pnl_Cyg_IngLiq.Caption) - CDbl(ipp_Cyg_OblMen.Value) < 0 Then
      pnl_Cyg_IngNet.Caption = "0.00" & " "
   Else
      pnl_Cyg_IngNet.Caption = Format(CDbl(pnl_Cyg_IngLiq.Caption) - CDbl(ipp_Cyg_OblMen.Value), "###,##0.00") & " "
   End If
      
   If CDbl(pnl_Cyg_IngNet.Caption) > 0# Then
      pnl_Cyg_IngDeu.Caption = Format(CDbl(pnl_Cyg_TotDeu.Caption) / CDbl(pnl_Cyg_IngNet.Caption), "##0.00") & " "
   Else
      pnl_Cyg_IngDeu.Caption = "0.00 "
   End If
   
   'Totales
   pnl_Tot_IngLiq.Caption = Format(CDbl(pnl_Tit_IngLiq.Caption) + CDbl(pnl_Cyg_IngLiq.Caption), "###,##0.00") & " "
   pnl_Tot_TotDeu.Caption = Format(CDbl(pnl_Tit_TotDeu.Caption) + CDbl(pnl_Cyg_TotDeu.Caption), "###,##0.00") & " "
   pnl_Tot_OblMen.Caption = Format(CDbl(ipp_Tit_OblMen.Value) + CDbl(ipp_Cyg_OblMen.Value), "###,##0.00") & "  "
   pnl_Tot_IngNet.Caption = Format(CDbl(pnl_Tot_IngLiq.Caption) - CDbl(pnl_Tot_OblMen.Caption), "###,##0.00") & " "
   
   If CDbl(pnl_Tot_IngNet.Caption) > 0# Then
      pnl_Tot_IngDeu.Caption = Format(CDbl(pnl_Tot_TotDeu.Caption) / CDbl(pnl_Tot_IngNet.Caption), "##0.00") & " "
   Else
      pnl_Tot_IngDeu.Caption = "0.00 "
   End If

   If moddat_g_int_TipMon = 2 Then
      r_dbl_IniSol = CDbl(Format(l_dbl_ApoPro * l_dbl_TipCam, "###,###,##0.00"))
   Else
      r_dbl_IniSol = l_dbl_ApoPro
   End If

   If CDbl(pnl_Tot_TotDeu.Caption) > 0# Then
      pnl_Tot_IniDeu.Caption = Format(CDbl(pnl_Tot_TotDeu.Caption) / r_dbl_IniSol * 100, "###,##0.00") & " "
   Else
      pnl_Tot_IniDeu.Caption = "0.00 "
   End If
End Sub

Private Sub fs_Calcul_CuoMPr()
   '******************************************************************************
   'FECHA: 27/08/2012
   'RAFAEL DURAND BANDA
   '******************************************************************************
   'SE COLOCA TEMPORALMENTE PARA CUMPLIR CON LAS POLITICAS.
   'LUEGO SE MODIFICARA PARA QUE LA CUOTA RENTA SE OBTENGA DE PARAMETROS
   '******************************************************************************
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "020" Then
      If CDbl(pnl_Tot_IngNet.Caption) >= 1000 And CDbl(pnl_Tot_IngNet.Caption) < 2000 Then
         l_dbl_CuoRta = 30
      End If
      If CDbl(pnl_Tot_IngNet.Caption) >= 2000 And CDbl(pnl_Tot_IngNet.Caption) < 4000 Then
         l_dbl_CuoRta = 35
      End If
      If CDbl(pnl_Tot_IngNet.Caption) >= 4000 Then
         'EXCEPCION PARA ESTA SOLICITUD AUTORIZADA POR JULIO RODRIGUEZ EL 20/12/2012
         If moddat_g_str_NumSol = "011001120086" Then
            l_dbl_CuoRta = 43
         Else
            l_dbl_CuoRta = 40
         End If
      End If

      If moddat_g_str_CodPrd = "002" Then
         If moddat_g_str_CodSub = "004" Then
            l_dbl_CuoRta = 30
         End If
      End If
      If moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "020" Then
         If moddat_g_str_CodSub = "005" Or moddat_g_str_CodSub = "014" Or moddat_g_str_CodSub = "015" Or moddat_g_str_CodSub = "016" Then
            l_dbl_CuoRta = 30
         End If
      End If
   End If
   '******************************************************************************
   '******************************************************************************
   '******************************************************************************
   
   If l_int_TipEva = 1 Then
      ipp_CuoMen.Value = Format(l_dbl_CuoRta / 100 * CDbl(pnl_Tot_IngNet.Caption), "###,##0.00") & " "
   Else
      If l_int_TipEva = 2 Or l_int_TipEva = 3 Or l_int_TipEva = 4 Then
         ipp_Tit_IngPri.Value = CDbl(Format((1 / l_dbl_CuoRta * 100 * ipp_CuoMen.Value), "###,###,##0.00"))
         cmb_Tit_TipPri.ListIndex = 1
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      pnl_CuoMPr.Caption = Format(ipp_CuoMen.Value, "###,##0.00") & " "
   Else
      pnl_CuoMPr.Caption = Format(ipp_CuoMen.Value / l_dbl_TipCam, "###,##0.00") & " "
   End If
End Sub

Private Sub fs_CalCuo()
Dim r_dbl_IntGra        As Double
Dim r_dbl_Portes        As Double
Dim r_int_TipVal_Des    As Integer
Dim r_dbl_Import_Des    As Double
Dim r_int_TipVal_Viv    As Integer
Dim r_dbl_Import_Viv    As Double
Dim r_dbl_PorCon        As Double
Dim r_dbl_TopCon        As Double
Dim r_dbl_MtoCon        As Double
Dim r_dbl_MtoNCo        As Double
Dim r_dbl_TasMVi        As Double
Dim r_dbl_ComCof        As Double
Dim r_dbl_TasCof        As Double

'variables nueva para la generacion del cronograma
Dim obj_Cronog          As Object
Dim int_Produc          As Integer
Dim int_CuoDbl          As Integer
Dim dbl_ValInm          As Double
Dim dbl_CuoIni          As Double
Dim dbl_MtoCon          As Double
Dim int_PlaPre          As Integer
Dim dbl_TasInt          As Double
Dim dbl_TasCof          As Double
Dim dbl_ComCof          As Double
Dim dat_FecDes          As Date
Dim int_DiaVct          As Integer
Dim int_PerGra          As Integer
Dim str_PriVct          As String
Dim dbl_Portes          As Double
Dim dbl_SegViv          As Double
Dim int_TipSDe          As Integer
Dim dbl_SegDes          As Double
Dim dbl_CuoMen          As Double
Dim dbl_CuoPbp          As Double
Dim dbl_IngReq          As Double

   'inicializa cuotas calculadas
   r_dbl_TasMVi = 0
   r_dbl_ComCof = 0
   r_dbl_TasCof = 0

   'Determina tasa y comision de cofide
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
      r_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   End If
   
   'Obtiene Tasa de Seguro de Desgravamen y Vivienda
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, l_int_TipSeg, moddat_g_int_TipMon, l_dbl_MtoPre, r_int_TipVal_Des, r_dbl_Import_Des, l_int_TasEsp)
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, 0, moddat_g_int_TipMon, l_dbl_ComVta, r_int_TipVal_Viv, r_dbl_Import_Viv, l_int_TasEsp)
   
   'Obtiene portes
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   Select Case moddat_g_str_CodPrd > 0
      'SE COMENTA ESTA PARTE DEL CODIGO PORQUE EL PRODUCTO ESTA DESCONTINUADO
      'Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd)
      '   r_dbl_PorCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
      '      r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   r_dbl_TopCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
      '      r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_PorCon, r_dbl_TopCon, l_dbl_TipCam, l_dbl_ComVta, l_int_PlaAno * 12, l_int_PerGra, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
      
      'SE COMENTA ESTA PARTE DEL CODIGO PORQUE EL PRODUCTO ESTA DESCONTINUADO
      'Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)
      '   'Para obtener porcentaje de TC
      '   r_dbl_PorCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
      '      r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'Para obtener tope de TC
      '   r_dbl_TopCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
      '      r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'NUEVA rutina de generacion de cronogramas
      '   int_Produc = 1
      '   int_CuoDbl = l_int_CuoDbl
      '   dbl_ValInm = l_dbl_ComVta
      '   dbl_CuoIni = l_dbl_ApoPro
      '   dbl_MtoCon = (l_dbl_ComVta - l_dbl_ApoPro) * (r_dbl_PorCon / 100)
      '   If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
      '   int_PlaPre = l_int_PlaAno * 12
      '   dbl_TasInt = l_dbl_TasInt
      '   dbl_TasCof = r_dbl_TasCof
      '   dbl_ComCof = r_dbl_ComCof
      '   dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
      '   int_DiaVct = l_int_DiaPag
      '   int_PerGra = l_int_PerGra
      '   str_PriVct = ""
      '   dbl_Portes = r_dbl_Portes
      '   dbl_SegViv = r_dbl_Import_Viv
      '   int_TipSDe = l_int_TipSeg - 10
      '   dbl_SegDes = r_dbl_Import_Des
      '
      '   'Calculando cronogramas
      '   Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
      '   Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
      '
      '   dbl_CuoMen = 0
      '   dbl_CuoPbp = 0
      '   dbl_IngReq = 0
      '   Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
      '
      '   'muestra valor cuota
      '   If moddat_g_int_TipMon = 1 Then
      '      ipp_CuoMen.Value = Format(dbl_CuoPbp, "###,###,##0.00") & " "
      '   Else
      '      ipp_CuoMen.Value = Format(dbl_CuoPbp * l_dbl_TipCam, "###,###,##0.00") & " "
      '   End If
   
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = 0
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         If moddat_g_int_TipMon = 1 Then
            ipp_CuoMen.Value = Format(dbl_CuoMen, "###,###,##0.00") & " "
         Else
            ipp_CuoMen.Value = Format(dbl_CuoMen * l_dbl_TipCam, "###,###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(l_dbl_ComVta) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'Muestra valor cuota
         If moddat_g_int_TipMon = 1 Then
            ipp_CuoMen.Value = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         Else
            ipp_CuoMen.Value = Format(dbl_CuoPbp * l_dbl_TipCam, "###,###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = 0
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         If moddat_g_int_TipMon = 1 Then
            ipp_CuoMen.Value = Format(dbl_CuoMen, "###,###,##0.00") & " "
         Else
            ipp_CuoMen.Value = Format(dbl_CuoMen * l_dbl_TipCam, "###,###,##0.00") & " "
         End If
   End Select
End Sub

