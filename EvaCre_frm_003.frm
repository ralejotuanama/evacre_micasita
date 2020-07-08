VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_10 
   Caption         =   "Form1"
   ClientHeight    =   10080
   ClientLeft      =   525
   ClientTop       =   480
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   14415
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _Version        =   65536
      _ExtentX        =   25426
      _ExtentY        =   17754
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
      Begin Threed.SSPanel SSPanel57 
         Height          =   1155
         Left            =   30
         TabIndex        =   149
         Top             =   8010
         Width           =   6465
         _Version        =   65536
         _ExtentX        =   11404
         _ExtentY        =   2037
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisObs 
            Height          =   825
            Left            =   30
            TabIndex        =   150
            Top             =   300
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   1455
            _Version        =   393216
            Rows            =   21
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel58 
            Height          =   285
            Left            =   60
            TabIndex        =   151
            Top             =   30
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Observ."
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
         Begin Threed.SSPanel SSPanel60 
            Height          =   285
            Left            =   1770
            TabIndex        =   152
            Top             =   30
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Observac"
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
         Begin Threed.SSPanel SSPanel61 
            Height          =   285
            Left            =   3930
            TabIndex        =   153
            Top             =   30
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Descargo"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   2295
         Left            =   30
         TabIndex        =   128
         Top             =   2220
         Width           =   6465
         _Version        =   65536
         _ExtentX        =   11404
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   129
            Top             =   60
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Index           =   0
            Left            =   1620
            TabIndex        =   130
            Top             =   390
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   1620
            TabIndex        =   131
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   90
            Left            =   30
            TabIndex        =   135
            Top             =   1080
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Index           =   1
            Left            =   1620
            TabIndex        =   136
            Top             =   1560
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
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
         Begin Threed.SSPanel SSPanel24 
            Height          =   315
            Left            =   1620
            TabIndex        =   137
            Top             =   1890
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   145
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "PRINCIPAL"
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   1
            Left            =   2880
            TabIndex        =   146
            Top             =   1230
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
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
         Begin Threed.SSPanel pnl_Ocupac 
            Height          =   315
            Index           =   3
            Left            =   1620
            TabIndex        =   147
            Top             =   1230
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SECUNDARIA"
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
            Caption         =   "Activ. Económica:"
            Height          =   315
            Left            =   60
            TabIndex        =   140
            Top             =   1230
            Width           =   1515
         End
         Begin VB.Label Label12 
            Caption         =   "Empresa:"
            Height          =   315
            Left            =   60
            TabIndex        =   139
            Top             =   1560
            Width           =   1515
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Inicio/Ingreso:"
            Height          =   315
            Left            =   60
            TabIndex        =   138
            Top             =   1890
            Width           =   1515
         End
         Begin VB.Label Label8 
            Caption         =   "Activ. Económica:"
            Height          =   315
            Left            =   60
            TabIndex        =   134
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Empresa:"
            Height          =   315
            Left            =   60
            TabIndex        =   133
            Top             =   390
            Width           =   1515
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Inicio/Ingreso:"
            Height          =   315
            Left            =   60
            TabIndex        =   132
            Top             =   720
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel40 
         Height          =   2325
         Left            =   6540
         TabIndex        =   81
         Top             =   6840
         Width           =   7785
         _Version        =   65536
         _ExtentX        =   13732
         _ExtentY        =   4101
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
         Begin VB.CommandButton Command15 
            Caption         =   "x"
            Height          =   315
            Left            =   2790
            TabIndex        =   158
            Top             =   60
            Width           =   345
         End
         Begin VB.CommandButton Command14 
            Caption         =   "x"
            Height          =   315
            Left            =   2790
            TabIndex        =   156
            Top             =   720
            Width           =   345
         End
         Begin VB.CommandButton Command13 
            Caption         =   "x"
            Height          =   315
            Left            =   2790
            TabIndex        =   155
            Top             =   390
            Width           =   345
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1830
            TabIndex        =   112
            Top             =   1890
            Width           =   1755
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Generar Alternativa"
            Height          =   375
            Left            =   60
            TabIndex        =   111
            Top             =   1890
            Width           =   1755
         End
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   1620
            TabIndex        =   85
            Top             =   60
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "70"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpDoubleSingle ipp_ComVta_Dolare 
            Height          =   315
            Left            =   1620
            TabIndex        =   88
            Top             =   390
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
         Begin Threed.SSPanel SSPanel41 
            Height          =   315
            Left            =   3900
            TabIndex        =   89
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel42 
            Height          =   315
            Left            =   6180
            TabIndex        =   91
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel44 
            Height          =   315
            Left            =   3900
            TabIndex        =   94
            Top             =   720
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel45 
            Height          =   315
            Left            =   6180
            TabIndex        =   96
            Top             =   720
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel46 
            Height          =   315
            Left            =   1620
            TabIndex        =   99
            Top             =   1050
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel47 
            Height          =   315
            Left            =   3900
            TabIndex        =   100
            Top             =   1050
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel48 
            Height          =   315
            Left            =   6180
            TabIndex        =   102
            Top             =   1050
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel49 
            Height          =   315
            Left            =   1620
            TabIndex        =   105
            Top             =   1380
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel50 
            Height          =   315
            Left            =   3900
            TabIndex        =   106
            Top             =   1380
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel51 
            Height          =   315
            Left            =   6180
            TabIndex        =   108
            Top             =   1380
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel52 
            Height          =   315
            Left            =   6930
            TabIndex        =   110
            Top             =   60
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "100.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel54 
            Height          =   90
            Left            =   30
            TabIndex        =   120
            Top             =   1740
            Width           =   7695
            _Version        =   65536
            _ExtentX        =   13573
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin EditLib.fpDoubleSingle fpDoubleSingle1 
            Height          =   315
            Left            =   1620
            TabIndex        =   157
            Top             =   720
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
         Begin VB.Label Label60 
            Caption         =   "Cuota/Renta %:"
            Height          =   285
            Left            =   5730
            TabIndex        =   119
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label59 
            Caption         =   "MP:"
            Height          =   315
            Left            =   5730
            TabIndex        =   109
            Top             =   1380
            Width           =   405
         End
         Begin VB.Label Label58 
            Caption         =   "S/."
            Height          =   315
            Left            =   3540
            TabIndex        =   107
            Top             =   1380
            Width           =   405
         End
         Begin VB.Label Label57 
            Caption         =   "Ult. Cuota US$:"
            Height          =   285
            Left            =   60
            TabIndex        =   104
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Label56 
            Caption         =   "MP:"
            Height          =   315
            Left            =   5730
            TabIndex        =   103
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label Label55 
            Caption         =   "S/."
            Height          =   315
            Left            =   3540
            TabIndex        =   101
            Top             =   1050
            Width           =   405
         End
         Begin VB.Label Label54 
            Caption         =   "1ra. Cuota US$:"
            Height          =   285
            Left            =   60
            TabIndex        =   98
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label53 
            Caption         =   "MP:"
            Height          =   315
            Left            =   5730
            TabIndex        =   97
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label52 
            Caption         =   "S/."
            Height          =   315
            Left            =   3540
            TabIndex        =   95
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label51 
            Caption         =   "C. Mensual US$:"
            Height          =   285
            Left            =   60
            TabIndex        =   93
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label49 
            Caption         =   "MP:"
            Height          =   315
            Left            =   5730
            TabIndex        =   92
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label46 
            Caption         =   "S/."
            Height          =   315
            Left            =   3540
            TabIndex        =   90
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label47 
            Caption         =   "Monto Crédito US$:"
            Height          =   285
            Left            =   60
            TabIndex        =   87
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label50 
            Caption         =   "Plazo (Meses):"
            Height          =   285
            Left            =   60
            TabIndex        =   86
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel26 
         Height          =   1455
         Left            =   30
         TabIndex        =   39
         Top             =   4560
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   2566
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
         Begin Threed.SSPanel SSPanel27 
            Height          =   315
            Left            =   1980
            TabIndex        =   40
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel28 
            Height          =   315
            Left            =   1980
            TabIndex        =   42
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel29 
            Height          =   315
            Left            =   1980
            TabIndex        =   44
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel30 
            Height          =   315
            Left            =   1980
            TabIndex        =   46
            Top             =   1050
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin VB.Label Label41 
            Caption         =   "%"
            Height          =   315
            Left            =   1650
            TabIndex        =   67
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label Label40 
            Caption         =   "%"
            Height          =   315
            Left            =   1650
            TabIndex        =   66
            Top             =   720
            Width           =   315
         End
         Begin VB.Label Label39 
            Caption         =   "S/."
            Height          =   315
            Left            =   1650
            TabIndex        =   65
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label38 
            Caption         =   "S/."
            Height          =   315
            Left            =   1650
            TabIndex        =   64
            Top             =   60
            Width           =   315
         End
         Begin VB.Label Label23 
            Caption         =   "T. Seguro Vivienda:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "T. Seguro Desg.:"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "T. Cambio M. Prest.:"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label20 
            Caption         =   "T. Cambio US$:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel17 
         Height          =   1695
         Left            =   6540
         TabIndex        =   33
         Top             =   5100
         Width           =   7785
         _Version        =   65536
         _ExtentX        =   13732
         _ExtentY        =   2990
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
            Caption         =   "&Nueva Alternativa"
            Height          =   375
            Left            =   60
            TabIndex        =   84
            Top             =   1260
            Width           =   1755
         End
         Begin VB.CommandButton cmd_BorAct 
            Caption         =   "&Borrar Alternativa"
            Height          =   375
            Left            =   1830
            TabIndex        =   83
            Top             =   1260
            Width           =   1755
         End
         Begin VB.CommandButton cmd_EdiAct 
            Caption         =   "&Editar Alternativa"
            Height          =   375
            Left            =   3600
            TabIndex        =   82
            Top             =   1260
            Width           =   1755
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   885
            Left            =   30
            TabIndex        =   34
            Top             =   330
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Alt."
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
            Left            =   2640
            TabIndex        =   36
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto a Otorgar S/."
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
         Begin Threed.SSPanel SSPanel21 
            Height          =   285
            Left            =   1560
            TabIndex        =   37
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo"
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
         Begin Threed.SSPanel SSPanel23 
            Height          =   285
            Left            =   4380
            TabIndex        =   38
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "C. Mens. S/."
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
         Begin Threed.SSPanel SSPanel39 
            Height          =   285
            Left            =   5910
            TabIndex        =   80
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota/Renta (%)"
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
      Begin Threed.SSPanel SSPanel16 
         Height          =   795
         Left            =   60
         TabIndex        =   30
         Top             =   9210
         Width           =   14295
         _Version        =   65536
         _ExtentX        =   25215
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
         Begin VB.CommandButton Command12 
            Height          =   675
            Left            =   60
            Picture         =   "EvaCre_frm_003.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Solicitud de Crédito"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton Command8 
            Height          =   675
            Left            =   2820
            Picture         =   "EvaCre_frm_003.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   13530
            Picture         =   "EvaCre_frm_003.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Cancelar"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton Command7 
            Height          =   675
            Left            =   2130
            Picture         =   "EvaCre_frm_003.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Cálculo ILD"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton Command6 
            Height          =   675
            Left            =   1440
            Picture         =   "EvaCre_frm_003.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Verificación Telefónica"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton Command4 
            Height          =   675
            Left            =   750
            Picture         =   "EvaCre_frm_003.frx":11A2
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Solicitud de Crédito"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12810
            Picture         =   "EvaCre_frm_003.frx":14AC
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Inversiones, Inmuebles, Autos y Enseres"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   12090
            Picture         =   "EvaCre_frm_003.frx":18EE
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   60
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1455
         Left            =   3270
         TabIndex        =   21
         Top             =   4560
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   2566
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
         Begin VB.CommandButton Command11 
            Caption         =   "..."
            Height          =   315
            Left            =   2790
            TabIndex        =   127
            Top             =   1050
            Width           =   345
         End
         Begin VB.CommandButton Command10 
            Caption         =   "..."
            Height          =   315
            Left            =   2790
            TabIndex        =   126
            Top             =   720
            Width           =   345
         End
         Begin VB.CommandButton Command9 
            Caption         =   "..."
            Height          =   315
            Left            =   2790
            TabIndex        =   125
            Top             =   390
            Width           =   345
         End
         Begin VB.CommandButton Command5 
            Caption         =   "..."
            Height          =   315
            Left            =   2790
            TabIndex        =   124
            Top             =   60
            Width           =   345
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   315
            Left            =   1710
            TabIndex        =   22
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "2,500.00 "
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   315
            Left            =   1710
            TabIndex        =   24
            Top             =   390
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   315
            Left            =   1710
            TabIndex        =   26
            Top             =   720
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   315
            Left            =   1710
            TabIndex        =   28
            Top             =   1050
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
         Begin VB.Label Label35 
            Caption         =   "S/."
            Height          =   315
            Left            =   1350
            TabIndex        =   63
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label Label34 
            Caption         =   "S/."
            Height          =   315
            Left            =   1350
            TabIndex        =   62
            Top             =   720
            Width           =   315
         End
         Begin VB.Label Label33 
            Caption         =   "S/."
            Height          =   315
            Left            =   1350
            TabIndex        =   61
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label32 
            Caption         =   "S/."
            Height          =   315
            Left            =   1350
            TabIndex        =   60
            Top             =   60
            Width           =   315
         End
         Begin VB.Label Label17 
            Caption         =   "Cónyuge ILD 2:"
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   1050
            Width           =   1200
         End
         Begin VB.Label Label16 
            Caption         =   "Cónyuge ILD 1:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label15 
            Caption         =   "Titular ILD 2:"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   390
            Width           =   1200
         End
         Begin VB.Label Label14 
            Caption         =   "Titular ILD 1:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1200
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14295
         _Version        =   65536
         _ExtentX        =   25215
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
            TabIndex        =   2
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Créditos"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   6510
            TabIndex        =   3
            Top             =   30
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel53 
            Height          =   315
            Left            =   6510
            TabIndex        =   118
            Top             =   300
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "EvaCre_frm_003.frx":1BF8
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   4
         Top             =   750
         Width           =   14295
         _Version        =   65536
         _ExtentX        =   25215
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
            Left            =   8970
            TabIndex        =   6
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            Top             =   1050
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   8
            Top             =   720
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   390
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8970
            TabIndex        =   10
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   8970
            TabIndex        =   11
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   315
            Left            =   8970
            TabIndex        =   19
            Top             =   1050
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
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
         Begin VB.Label Label5 
            Caption         =   "Situac. Crediticia:"
            Height          =   315
            Left            =   7410
            TabIndex        =   20
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7410
            TabIndex        =   17
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Left            =   7410
            TabIndex        =   13
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7410
            TabIndex        =   12
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel31 
         Height          =   1905
         Left            =   30
         TabIndex        =   48
         Top             =   6060
         Width           =   6465
         _Version        =   65536
         _ExtentX        =   11404
         _ExtentY        =   3360
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
         Begin Threed.SSPanel SSPanel32 
            Height          =   315
            Left            =   1950
            TabIndex        =   49
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel33 
            Height          =   315
            Left            =   1950
            TabIndex        =   50
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel34 
            Height          =   315
            Left            =   1950
            TabIndex        =   51
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel35 
            Height          =   315
            Left            =   1950
            TabIndex        =   52
            Top             =   1530
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   3660
            TabIndex        =   68
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   315
            Left            =   3660
            TabIndex        =   69
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   315
            Left            =   3660
            TabIndex        =   70
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel36 
            Height          =   315
            Left            =   5250
            TabIndex        =   74
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "30,000.00 "
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
         Begin Threed.SSPanel SSPanel37 
            Height          =   315
            Left            =   5250
            TabIndex        =   75
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel38 
            Height          =   315
            Left            =   5250
            TabIndex        =   76
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   315
            Left            =   1950
            TabIndex        =   121
            Top             =   1050
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin Threed.SSPanel SSPanel56 
            Height          =   90
            Left            =   30
            TabIndex        =   148
            Top             =   1410
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin VB.Label Label18 
            Caption         =   "Total ILD:"
            Height          =   315
            Left            =   60
            TabIndex        =   123
            Top             =   1050
            Width           =   1200
         End
         Begin VB.Label Label36 
            Caption         =   "S/."
            Height          =   315
            Left            =   1620
            TabIndex        =   122
            Top             =   1050
            Width           =   315
         End
         Begin VB.Label Label45 
            Caption         =   "MP:"
            Height          =   315
            Left            =   4920
            TabIndex        =   79
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label44 
            Caption         =   "MP:"
            Height          =   315
            Left            =   4920
            TabIndex        =   78
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label43 
            Caption         =   "MP:"
            Height          =   315
            Left            =   4920
            TabIndex        =   77
            Top             =   60
            Width           =   405
         End
         Begin VB.Label Label42 
            Caption         =   "S/."
            Height          =   315
            Left            =   3240
            TabIndex        =   73
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label37 
            Caption         =   "S/."
            Height          =   315
            Left            =   3240
            TabIndex        =   72
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label19 
            Caption         =   "S/."
            Height          =   315
            Left            =   3240
            TabIndex        =   71
            Top             =   60
            Width           =   405
         End
         Begin VB.Label Label31 
            Caption         =   "US$"
            Height          =   315
            Left            =   1560
            TabIndex        =   59
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label30 
            Caption         =   "US$"
            Height          =   315
            Left            =   1560
            TabIndex        =   58
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label29 
            Caption         =   "US$"
            Height          =   315
            Left            =   1560
            TabIndex        =   57
            Top             =   60
            Width           =   405
         End
         Begin VB.Label Label28 
            Caption         =   "Valor Inmueble:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label27 
            Caption         =   "Aporte Propio:"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label26 
            Caption         =   "Período de Gracia:"
            Height          =   315
            Left            =   60
            TabIndex        =   54
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label25 
            Caption         =   "C. Mensual Sug.:"
            Height          =   315
            Left            =   60
            TabIndex        =   53
            Top             =   720
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel25 
         Height          =   2835
         Left            =   6540
         TabIndex        =   141
         Top             =   2220
         Width           =   7785
         _Version        =   65536
         _ExtentX        =   13732
         _ExtentY        =   5001
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
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2445
            Left            =   30
            TabIndex        =   142
            Top             =   330
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4313
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   4210752
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel55 
            Height          =   285
            Left            =   60
            TabIndex        =   143
            Top             =   60
            Width           =   5805
            _Version        =   65536
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Verificación"
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
         Begin Threed.SSPanel SSPanel59 
            Height          =   285
            Left            =   5850
            TabIndex        =   144
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Resultado"
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
   End
End
Attribute VB_Name = "frm_EvaCre_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

