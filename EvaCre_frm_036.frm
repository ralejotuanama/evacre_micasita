VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   2415
   ClientTop       =   1980
   ClientWidth     =   11250
   Icon            =   "EvaCre_frm_036.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8880
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   15663
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   3345
         Left            =   30
         TabIndex        =   44
         Top             =   1890
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   5900
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
            Height          =   3255
            Index           =   5
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   11
         Top             =   7260
         WhatsThisHelpID =   11175
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
         Begin Threed.SSPanel pnl_IngNet 
            Height          =   315
            Left            =   2700
            TabIndex        =   12
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
         Begin Threed.SSPanel pnl_CuoMPr 
            Height          =   315
            Left            =   8460
            TabIndex        =   13
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
            Left            =   8460
            TabIndex        =   14
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
            Left            =   2700
            TabIndex        =   7
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
         Begin VB.Label Label9 
            Caption         =   "Total Ingreso Líquido Neto (S/.):"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "C. Mensual Aprob. (S/.):"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   390
            Width           =   2085
         End
         Begin VB.Label Label4 
            Caption         =   "C. Mensual Aprob. (M. Prest.):"
            Height          =   315
            Left            =   5640
            TabIndex        =   16
            Top             =   390
            Width           =   2235
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Cambio (M.l Prest.):"
            Height          =   315
            Left            =   5640
            TabIndex        =   15
            Top             =   60
            Width           =   2235
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   465
         Left            =   30
         TabIndex        =   19
         Top             =   6750
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
         Begin EditLib.fpDoubleSingle ipp_MtoDeu 
            Height          =   315
            Left            =   8460
            TabIndex        =   46
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
            Caption         =   "Total Endeudamiento (S/.):"
            Height          =   315
            Index           =   7
            Left            =   5640
            TabIndex        =   47
            Top             =   60
            Width           =   2145
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Obligaciones Mensuales (S/.):"
            Height          =   315
            Index           =   6
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   2535
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   21
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
            TabIndex        =   22
            Top             =   60
            Width           =   9495
            _Version        =   65536
            _ExtentX        =   16748
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Evaluación de Solicitudes - Cálculo de Ingresos"
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
            Picture         =   "EvaCre_frm_036.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   28
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   32
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "F. Ingreso Instancia:"
            Height          =   315
            Left            =   8040
            TabIndex        =   29
            Top             =   390
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   34
         Top             =   8070
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
         Begin VB.CommandButton cmd_Calcul 
            Height          =   675
            Left            =   720
            Picture         =   "EvaCre_frm_036.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "EvaCre_frm_036.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10470
            Picture         =   "EvaCre_frm_036.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   9780
            Picture         =   "EvaCre_frm_036.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1425
         Left            =   30
         TabIndex        =   35
         Top             =   5280
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
            Left            =   8460
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
            Left            =   8460
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
            TabIndex        =   36
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
            Left            =   8460
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
         Begin VB.Label Label5 
            Caption         =   "Total ILD Mancomunado (S/.):"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   1050
            Width           =   2325
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Sec. Cónyuge) (S/.):"
            Height          =   315
            Index           =   2
            Left            =   5640
            TabIndex        =   42
            Top             =   390
            Width           =   2595
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Princ. Cónyuge) (S/.):"
            Height          =   315
            Index           =   1
            Left            =   5640
            TabIndex        =   41
            Top             =   60
            Width           =   2595
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Sec. Cliente) (S/.):"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   40
            Top             =   390
            Width           =   2535
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Act. Eco. Princ. Cliente) (S/.):"
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   39
            Top             =   60
            Width           =   2595
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Otras Activ. Cliente) (S/.):"
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   38
            Top             =   720
            Width           =   2385
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "ILD (Otras Activ. Cónyuge) (S/.):"
            Height          =   315
            Index           =   4
            Left            =   5640
            TabIndex        =   37
            Top             =   720
            Width           =   2475
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CliNCo()   As modcal_g_est_CuoCli
Dim l_dbl_CuoRta     As Double
Dim l_dbl_MtoPre     As Double
Dim l_dbl_ComVta     As Double
Dim l_int_TipEva     As Integer
Dim l_dbl_TasInt     As Double
Dim l_int_PlaAno     As Integer
Dim l_int_PerGra     As Integer
Dim l_str_EmpSeg     As String
Dim l_int_TipSeg     As Integer
Dim l_int_DiaPag     As Integer
Dim l_int_MonAho     As Integer
Dim l_dbl_MtoAho     As Double
Dim l_dbl_TipCam     As Double
Dim l_dbl_ApoPro     As Double

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

   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, l_int_TipSeg, moddat_g_int_TipMon, l_dbl_MtoPre, r_int_TipVal_Des, r_dbl_Import_Des)
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, 0, moddat_g_int_TipMon, l_dbl_ComVta, r_int_TipVal_Viv, r_dbl_Import_Viv)
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   Select Case moddat_g_str_CodPrd
      Case "001"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
   
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         
         Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_PorCon, r_dbl_TopCon, l_dbl_TipCam, l_dbl_ComVta, l_int_PlaAno * 12, l_int_PerGra, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
      Case "002"
         Call gs_Cronog_MiCasita(l_arr_CliNCo(), l_dbl_ComVta, l_dbl_MtoPre, l_int_PlaAno * 12, 2, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, l_int_PerGra)

      Case "003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
   
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         
         Call gs_Cronog_CME_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_PorCon, r_dbl_TopCon, l_dbl_ComVta, l_int_PlaAno * 12, l_int_PerGra, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
      Case "004"
         r_dbl_TopCon = 0

         'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
         '   r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         'End If
      
         Call gs_Cronog_Mihogar_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_TopCon, l_dbl_ComVta, l_int_PlaAno * 12, l_int_PerGra, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   End Select
End Sub

Private Sub cmd_Calcul_Click()
   Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\winnt\system32\calc.exe", vbNormalFocus)
   
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicación", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_Aho_CuoMin    As Double
   Dim r_dbl_CuoAho        As Double
   Dim r_dbl_CuoRta        As Double
   Dim r_dbl_IngMin        As Double
   Dim r_dbl_ApoPro        As Double
   Dim r_dbl_Ini_IniDeu    As Double
   
   'Obteniendo Ingreso Mínimo de Parámetro por Producto
   r_dbl_IngMin = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "014") Then
      r_dbl_IngMin = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   If l_int_TipEva = 1 Then
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
      
      If CDbl(pnl_IngNet.Caption) < r_dbl_IngMin Then
         MsgBox "El Ingreso Neto es menor al Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   ElseIf l_int_TipEva = 2 Then
      If CDbl(ipp_CuoMen.Text) = 0 Then
         MsgBox "Debe ingresar el importe de la Cuota Mensual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_CuoMen)
         Exit Sub
      End If
      
      r_dbl_Aho_CuoMin = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "004") Then
         r_dbl_Aho_CuoMin = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      If moddat_g_int_TipMon = 2 Then
         r_dbl_CuoAho = CDbl(Format(ipp_CuoMen.Value / l_dbl_TipCam, "###,###,##0.00"))
      Else
         r_dbl_CuoAho = CDbl(Format(ipp_CuoMen.Value, "###,###,##0.00"))
      End If
      
      If r_dbl_CuoAho < r_dbl_Aho_CuoMin Then
         MsgBox "El Importe de la Cuota Mensual Ahorrada no cumple con el mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_CuoMen)
         Exit Sub
      End If
      
      r_dbl_CuoRta = CDbl(Format(ipp_CuoMen.Value / CDbl(pnl_IngNet.Caption) * 100, "##0.00"))
      
      If r_dbl_CuoRta > l_dbl_CuoRta Then
         MsgBox "La Relación Cuota / Renta excede el parámetro.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_CuoMen)
         Exit Sub
      End If
      
      If CDbl(pnl_IngNet.Caption) < r_dbl_IngMin Then
         MsgBox "El Ingreso Neto es menor al Ingreso Mínimo solicitado para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   ElseIf l_int_TipEva = 3 Then
      'Verificando Nivel de Endeudamiento
      If moddat_g_int_TipMon = 1 Then
         r_dbl_ApoPro = l_dbl_ApoPro
      Else
         r_dbl_ApoPro = l_dbl_ApoPro * l_dbl_TipCam
      End If
      r_dbl_ApoPro = CDbl(Format(r_dbl_ApoPro, "###,##0.00"))
      
      'Obteniendo Parámetro de Relación Deuda / Cuota Inicial
      r_dbl_Ini_IniDeu = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "003") Then
         r_dbl_Ini_IniDeu = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      If CDbl(ipp_MtoDeu.Text) / r_dbl_ApoPro * 100 > r_dbl_Ini_IniDeu Then
         MsgBox "La Relación Total Deuda / Aporte Inicial no se ajusta al Parámetro requerido." & Format(CDbl(ipp_MtoDeu.Text) / r_dbl_ApoPro * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
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
      
      g_str_Parame = "USP_TRA_EVACRE_REGING ("
      
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
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CuoMen.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TipCam.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDeu.Text)) & ", "
      
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

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_01.Show 1
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
   Call fs_Limpia
   Call fs_DatCre
   
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   l_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "013") Then
      l_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   grd_Listad(5).ColWidth(0) = 3000
   grd_Listad(5).ColWidth(1) = 7940

   grd_Listad(5).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(5).ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad(5))
End Sub

Private Sub fs_Limpia()
   ipp_IngPri_Cli.Value = 0
   ipp_IngSec_Cli.Value = 0
   ipp_IngAdi_Cli.Value = 0
   ipp_IngPri_Cyg.Value = 0
   ipp_IngSec_Cyg.Value = 0
   ipp_IngAdi_Cyg.Value = 0
   
   ipp_OblMen.Value = 0
   ipp_MtoDeu.Value = 0
   
   pnl_TotIng.Caption = "0.00 "
   pnl_IngNet.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   ipp_CuoMen.Value = 0
   pnl_CuoMPr.Caption = "0.00 "
   
   
   'Obteniendo Tipo de Cambio de Moneda del Préstamo
   l_dbl_TipCam = 0
   
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   End If
End Sub

Private Sub fs_DatCre()
   Dim r_dbl_TipCam     As Double
   
   Call gs_LimpiaGrid(grd_Listad(5))
   
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
   
   grd_Listad(5).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Sub-Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Evaluación"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Préstamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tasa de Interés"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo (US$)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo (S/.)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Cambio Referencial"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL / g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Plazo (Años)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Período de Gracia (Meses)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Cuotas Extraordinarias"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Compañía de Seguros"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Día de Pago"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Institución Financiera de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Moneda de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!SOLMAE_MONAHO)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Meses Ahorrados"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   If moddat_g_int_TipMon = 1 Then
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
      l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_SOL
   Else
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
      l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_DOL
   End If
   
   l_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
   
   l_int_TipEva = g_rst_Princi!SOLMAE_TIPEVA
   l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
   l_int_PlaAno = g_rst_Princi!SOLMAE_PLAANO
   l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
   l_str_EmpSeg = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
   l_int_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
   l_int_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
   l_int_MonAho = g_rst_Princi!SOLMAE_MONAHO
   l_dbl_MtoAho = g_rst_Princi!SOLMAE_MTOAHO
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If l_int_TipEva = 1 Then
      'Tipo de Evaluación Normal
      
      ipp_CuoMen.Enabled = False
      ipp_MtoDeu.Enabled = False
      
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
   ElseIf l_int_TipEva = 2 Then
      'Tipo de Evaluación Ahorro Programado
      
      ipp_CuoMen.Enabled = True
   
      ipp_IngPri_Cli.Enabled = False
      ipp_IngSec_Cli.Enabled = False
      ipp_IngAdi_Cli.Enabled = False
   
      ipp_IngPri_Cyg.Enabled = False
      ipp_IngSec_Cyg.Enabled = False
      ipp_IngAdi_Cyg.Enabled = False
      
      ipp_OblMen.Enabled = False
      ipp_MtoDeu.Enabled = False
      
      'Se asume Monto de Ahorro como Cuota Máxima Mensual
      r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, l_int_MonAho)
      
      If l_int_MonAho = 1 Then
         ipp_CuoMen.Value = l_dbl_MtoAho
      Else
         ipp_CuoMen.Value = CDbl(Format(l_dbl_MtoAho * r_dbl_TipCam, "###,###,##0.00"))
      End If
   ElseIf l_int_TipEva = 3 Then
      'Tipo de Evaluación Inicial 30%-35%
      ipp_CuoMen.Enabled = False
   
      ipp_IngPri_Cli.Enabled = False
      ipp_IngSec_Cli.Enabled = False
      ipp_IngAdi_Cli.Enabled = False
   
      ipp_IngPri_Cyg.Enabled = False
      ipp_IngSec_Cyg.Enabled = False
      ipp_IngAdi_Cyg.Enabled = False
      
      ipp_OblMen.Enabled = False
      ipp_MtoDeu.Enabled = True
      
      'Calculando Cuota Mensual
      Call fs_CalCuo
      
      'Se asume Cuota Mensual como Cuota Máxima Mensual
      If moddat_g_int_TipMon = 1 Then
         ipp_CuoMen.Value = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
      Else
         ipp_CuoMen.Value = Format(l_arr_CliNCo(2).CuoCli_ValCuo * l_dbl_TipCam, "###,###,##0.00") & " "
      End If
   ElseIf l_int_TipEva = 4 Then
      'Tipo de Evaluación Inicial 50% (S/Evaluación)
      ipp_CuoMen.Enabled = False
   
      ipp_IngPri_Cli.Enabled = False
      ipp_IngSec_Cli.Enabled = False
      ipp_IngAdi_Cli.Enabled = False
   
      ipp_IngPri_Cyg.Enabled = False
      ipp_IngSec_Cyg.Enabled = False
      ipp_IngAdi_Cyg.Enabled = False
      
      ipp_OblMen.Enabled = False
      ipp_MtoDeu.Enabled = False
      
      'Calculando Cuota Mensual
      Call fs_CalCuo
      
      'Se asume Cuota Mensual como Cuota Máxima Mensual
      If moddat_g_int_TipMon = 1 Then
         ipp_CuoMen.Value = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,###,##0.00") & " "
      Else
         ipp_CuoMen.Value = Format(l_arr_CliNCo(2).CuoCli_ValCuo * l_dbl_TipCam, "###,###,##0.00") & " "
      End If
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
   ipp_MtoDeu.Value = g_rst_Princi!EVACRE_MTODEU
   
   ipp_CuoMen.Value = g_rst_Princi!EVACRE_CUOSOL

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   moddat_g_int_FlgGrb = 2
End Sub

Private Sub ipp_CuoMen_Change()
   If l_int_TipEva = 2 Or l_int_TipEva = 3 Or l_int_TipEva = 4 Then
      'If moddat_g_int_TipMon = 1 Then
      '   ipp_IngPri_Cli.Value = CDbl(Format(1 / l_dbl_CuoRta * 100 * ipp_CuoMen.Value, "###,###,##0.00"))
      'Else
         ipp_IngPri_Cli.Value = CDbl(Format((1 / l_dbl_CuoRta * 100 * ipp_CuoMen.Value) + (ipp_CuoMen.Value * 1.5 / 100), "###,###,##0.00"))
      'End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      pnl_CuoMPr.Caption = Format(ipp_CuoMen.Value, "###,###,##0.00") & " "
   Else
      pnl_CuoMPr.Caption = Format(CDbl(ipp_CuoMen.Text) / l_dbl_TipCam, "###,###,###,##0.00") & " "
   End If
End Sub

Private Sub ipp_CuoMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_IngAdi_Cli_Change()
   Call fs_Calcul_IngNet
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
   Call fs_Calcul_IngNet
End Sub

Private Sub ipp_IngAdi_Cyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_OblMen)
   End If
End Sub

Private Sub ipp_IngPri_Cli_Change()
   Call fs_Calcul_IngNet
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
   Call fs_Calcul_IngNet
End Sub

Private Sub ipp_IngSec_Cli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IngAdi_Cli)
   End If
End Sub

Private Sub ipp_IngPri_Cyg_Change()
   Call fs_Calcul_IngNet
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
   Call fs_Calcul_IngNet
End Sub

Private Sub ipp_IngSec_Cyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IngAdi_Cyg)
   End If
End Sub

Private Sub fs_Calcul_IngNet()
   pnl_TotIng.Caption = Format(CDbl(ipp_IngPri_Cli.Text) + CDbl(ipp_IngSec_Cli.Text) + CDbl(ipp_IngAdi_Cli.Text) + CDbl(ipp_IngPri_Cyg.Text) + CDbl(ipp_IngSec_Cyg.Text) + CDbl(ipp_IngAdi_Cyg.Text), "###,###,##0.00") & " "
   pnl_IngNet.Caption = Format(CDbl(pnl_TotIng.Caption) - CDbl(ipp_OblMen.Text), "###,###,##0.00") & " "
   
   If l_int_TipEva = 1 Then
      ipp_CuoMen.Value = CDbl(Format(CDbl(pnl_IngNet.Caption) * l_dbl_CuoRta / 100, "###,###,##0.00"))
   End If
End Sub

Private Sub ipp_MtoDeu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_OblMen_Change()
   Call fs_Calcul_IngNet
End Sub

Private Sub ipp_OblMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_CuoMen.Enabled Then
         Call gs_SetFocus(ipp_CuoMen)
      Else
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub


