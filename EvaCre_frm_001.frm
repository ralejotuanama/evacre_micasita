VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   1800
   ClientTop       =   1875
   ClientWidth     =   11250
   Icon            =   "EvaCre_frm_001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6990
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   12330
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
         Height          =   1095
         Left            =   30
         TabIndex        =   40
         Top             =   3870
         WhatsThisHelpID =   11175
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
         Begin Threed.SSPanel pnl_IngNet 
            Height          =   315
            Left            =   2700
            TabIndex        =   41
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
         Begin Threed.SSPanel pnl_CuoMen 
            Height          =   315
            Left            =   2700
            TabIndex        =   43
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
         Begin Threed.SSPanel pnl_CuoMPr 
            Height          =   315
            Left            =   2700
            TabIndex        =   45
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
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   8430
            TabIndex        =   47
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
         Begin VB.Label Label7 
            Caption         =   "Tipo de Cambio (M.l Prest.):"
            Height          =   315
            Left            =   5640
            TabIndex        =   48
            Top             =   60
            Width           =   2235
         End
         Begin VB.Label Label4 
            Caption         =   "Cuota Mensual Aprob. (M. Prest.):"
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   720
            Width           =   2445
         End
         Begin VB.Label Label8 
            Caption         =   "Cuota Mensual Aprobada (S/.):"
            Height          =   315
            Left            =   60
            TabIndex        =   44
            Top             =   390
            Width           =   2505
         End
         Begin VB.Label Label9 
            Caption         =   "Total Ingreso Líquido Neto (S/.):"
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   2325
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   465
         Left            =   30
         TabIndex        =   38
         Top             =   3360
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   820
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
         Begin EditLib.fpDoubleSingle ipp_OblMen 
            Height          =   315
            Left            =   2700
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
         Begin VB.Label lbl_Etique 
            Caption         =   "Obligaciones Mensuales:"
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   2535
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1125
         Left            =   30
         TabIndex        =   33
         Top             =   5010
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   1984
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
         Begin VB.ComboBox cmb_FlgCon 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   60
            Width           =   885
         End
         Begin VB.TextBox txt_ObsEva 
            Height          =   675
            Left            =   2700
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "EvaCre_frm_001.frx":000C
            Top             =   390
            Width           =   8415
         End
         Begin VB.Label Label12 
            Caption         =   "Aprobación Condicionada:"
            Height          =   315
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   2205
         End
         Begin VB.Label Label13 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   2205
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   13
            Top             =   60
            Width           =   7455
            _Version        =   65536
            _ExtentX        =   13150
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Evaluación de Solicitudes"
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
            Picture         =   "EvaCre_frm_001.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   14
         Top             =   750
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1860
            TabIndex        =   15
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
            Left            =   1860
            TabIndex        =   16
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
            Left            =   9690
            TabIndex        =   17
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
            Left            =   9690
            TabIndex        =   18
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
            Left            =   1860
            TabIndex        =   19
            Top             =   720
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin VB.Label Label6 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8040
            TabIndex        =   24
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   21
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   720
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   25
         Top             =   6180
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   9780
            Picture         =   "EvaCre_frm_001.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10470
            Picture         =   "EvaCre_frm_001.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1425
         Left            =   30
         TabIndex        =   26
         Top             =   1890
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   2514
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
         Begin EditLib.fpDoubleSingle ipp_IngPri_Cli 
            Height          =   315
            Left            =   2700
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
         Begin EditLib.fpDoubleSingle ipp_IngSec_Cli 
            Height          =   315
            Left            =   2700
            TabIndex        =   1
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
         Begin EditLib.fpDoubleSingle ipp_IngPri_Cyg 
            Height          =   315
            Left            =   8430
            TabIndex        =   3
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
         Begin EditLib.fpDoubleSingle ipp_IngSec_Cyg 
            Height          =   315
            Left            =   8430
            TabIndex        =   4
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
         Begin Threed.SSPanel pnl_TotIng 
            Height          =   315
            Left            =   2700
            TabIndex        =   27
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
         Begin EditLib.fpDoubleSingle ipp_IngAdi_Cli 
            Height          =   315
            Left            =   2700
            TabIndex        =   2
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
         Begin EditLib.fpDoubleSingle ipp_IngAdi_Cyg 
            Height          =   315
            Left            =   8430
            TabIndex        =   5
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
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Otras Actividades Cónyuge):"
            Height          =   315
            Index           =   4
            Left            =   5640
            TabIndex        =   37
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Otras Actividades Cliente):"
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   36
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Principal Cliente):"
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   2415
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Secundaria Cliente):"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   31
            Top             =   390
            Width           =   2535
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Principal Cónyuge):"
            Height          =   315
            Index           =   1
            Left            =   5640
            TabIndex        =   30
            Top             =   60
            Width           =   2595
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Secundaria Cónyuge):"
            Height          =   315
            Index           =   2
            Left            =   5640
            TabIndex        =   29
            Top             =   390
            Width           =   2625
         End
         Begin VB.Label Label5 
            Caption         =   "Total ILD Mancomunado:"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   1050
            Width           =   2325
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_CuoRta     As Double
Dim l_dbl_TipCam     As Double

Private Sub cmb_FlgCon_Click()
   Call gs_SetFocus(txt_ObsEva)
End Sub

Private Sub cmb_FlgCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgCon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgActPri_Cli = 1 Then
      If CDbl(ipp_IngPri_Cli.Text) = 0 Then
         MsgBox "Debe ingresar el Ingreso Líquido de la Actividad Económica Principal del Cliente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IngPri_Cli)
         Exit Sub
      End If
   End If

   If moddat_g_int_FlgActSec_Cli = 1 Then
      If CDbl(ipp_IngSec_Cli.Text) = 0 Then
         MsgBox "Debe ingresar el Ingreso Líquido de la Actividad Económica Secundaria del Cliente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IngSec_Cli)
         Exit Sub
      End If
   End If

   If moddat_g_int_FlgActPri_Cyg = 1 Then
      If CDbl(ipp_IngPri_Cyg.Text) = 0 Then
         MsgBox "Debe ingresar el Ingreso Líquido de la Actividad Económica Principal del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IngPri_Cyg)
         Exit Sub
      End If
   End If

   If moddat_g_int_FlgActSec_Cyg = 1 Then
      If CDbl(ipp_IngSec_Cyg.Text) = 0 Then
         MsgBox "Debe ingresar el Ingreso Líquido de la Actividad Económica Secundaria del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IngSec_Cyg)
         Exit Sub
      End If
   End If

   If cmb_FlgCon.ListIndex = -1 Then
      MsgBox "Debe indicar si la Aprobación es condicionada.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgCon)
      Exit Sub
   End If
   
   If cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex) = 1 Then
      If Len(Trim(txt_ObsEva.Text)) = 0 Then
         MsgBox "Debe ingresar las observaciones.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_ObsEva)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de registrar la Evaluación?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVACRE ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngPri_Cli.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngSec_Cli.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngPri_Cyg.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngSec_Cyg.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngAdi_Cli.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IngAdi_Cyg.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotIng.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_OblMen.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngNet.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMen.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TipCam.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgCon.ItemData(cmb_FlgCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_ObsEva.Text & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                        'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGrb = 2
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   Screen.MousePointer = 0
   
   MsgBox "Evaluación registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_IngIns.Caption = moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 21)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgCon, 1, "214")
   
   l_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "013") Then
      l_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   ipp_IngPri_Cli.Value = 0
   ipp_IngSec_Cli.Value = 0
   ipp_IngAdi_Cli.Value = 0
   ipp_IngPri_Cyg.Value = 0
   ipp_IngSec_Cyg.Value = 0
   ipp_IngAdi_Cyg.Value = 0
   
   ipp_OblMen.Value = 0
   
   pnl_TotIng.Caption = "0.00 "
   pnl_IngNet.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   pnl_CuoMen.Caption = "0.00 "
   pnl_CuoMPr.Caption = "0.00 "
   
   txt_ObsEva.Text = ""
   
   
   If moddat_g_int_FlgActSec_Cli = 0 Then
      ipp_IngSec_Cli.Enabled = False
   End If
   
   If moddat_g_int_FlgActPri_Cyg = 0 Then
      ipp_IngPri_Cyg.Enabled = False
      ipp_IngAdi_Cyg.Enabled = False
   End If
   
   If moddat_g_int_FlgActSec_Cyg = 0 Then
      ipp_IngSec_Cyg.Enabled = False
   End If
   
   'Obteniendo Tipo de Cambio de Moneda del Préstamo
   l_dbl_TipCam = 0
   
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   End If
End Sub

Private Sub ipp_IngAdi_Cli_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngAdi_Cli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_IngPri_Cyg.Enabled Then
         Call gs_SetFocus(ipp_IngPri_Cyg)
      Else
         Call gs_SetFocus(ipp_OblMen)
      End If
   End If
End Sub

Private Sub ipp_IngAdi_Cyg_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngAdi_Cyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_OblMen)
   End If
End Sub

Private Sub ipp_IngPri_Cli_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngPri_Cli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_IngSec_Cli.Enabled Then
         Call gs_SetFocus(ipp_IngSec_Cli)
      Else
         Call gs_SetFocus(ipp_IngAdi_Cli)
      End If
   End If
End Sub

Private Sub ipp_IngSec_Cli_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngSec_Cli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IngAdi_Cli)
   End If
End Sub

Private Sub ipp_IngPri_Cyg_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngPri_Cyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_IngSec_Cyg.Enabled Then
         Call gs_SetFocus(ipp_IngSec_Cyg)
      Else
         Call gs_SetFocus(ipp_IngAdi_Cyg)
      End If
   End If
End Sub

Private Sub ipp_IngSec_Cyg_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_IngSec_Cyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IngAdi_Cyg)
   End If
End Sub

Private Sub fs_CalCuo()
   pnl_TotIng.Caption = Format(CDbl(ipp_IngPri_Cli.Text) + CDbl(ipp_IngSec_Cli.Text) + CDbl(ipp_IngAdi_Cli.Text) + CDbl(ipp_IngPri_Cyg.Text) + CDbl(ipp_IngSec_Cyg.Text) + CDbl(ipp_IngAdi_Cyg.Text), "###,###,##0.00") & " "
   pnl_IngNet.Caption = Format(CDbl(pnl_TotIng.Caption) - CDbl(ipp_OblMen.Text), "###,###,##0.00") & " "
   pnl_CuoMen.Caption = Format(CDbl(pnl_IngNet.Caption) * l_dbl_CuoRta / 100, "###,###,##0.00") & " "
   
   If moddat_g_int_TipMon = 1 Then
      pnl_CuoMPr.Caption = pnl_CuoMen.Caption
   Else
      pnl_CuoMPr.Caption = Format(CDbl(pnl_CuoMen.Caption) / l_dbl_TipCam, "###,###,###,##0.00") & " "
   End If
End Sub

Private Sub fs_Buscar_DatEva()
   moddat_g_int_FlgGrb = 1
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   ipp_IngPri_Cli.Value = g_rst_Princi!EVACRE_INGDP1
   ipp_IngSec_Cli.Value = g_rst_Princi!EVACRE_INGDP2
   ipp_IngAdi_Cli.Value = g_rst_Princi!EVACRE_INGAD1
   
   ipp_IngPri_Cyg.Value = g_rst_Princi!EVACRE_INGDP3
   ipp_IngSec_Cyg.Value = g_rst_Princi!EVACRE_INGDP4
   ipp_IngAdi_Cyg.Value = g_rst_Princi!EVACRE_INGAD2
   
   ipp_OblMen.Value = g_rst_Princi!EVACRE_OBLMEN
   
   Call gs_BuscarCombo_Item(cmb_FlgCon, g_rst_Princi!EVACRE_FLGCON)
   
   txt_ObsEva.Text = Trim(g_rst_Princi!EVACRE_OBSEVA & "")
   
   Call fs_CalCuo

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   moddat_g_int_FlgGrb = 2
End Sub

Private Sub ipp_OblMen_Change()
   Call fs_CalCuo
End Sub

Private Sub ipp_OblMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_FlgCon)
   End If
End Sub

Private Sub txt_ObsEva_GotFocus()
   Call gs_SelecTodo(txt_ObsEva)
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

