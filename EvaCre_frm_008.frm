VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaEmp_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   4830
   ClientTop       =   645
   ClientWidth     =   9675
   Icon            =   "EvaCre_frm_008.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9675
      _Version        =   65536
      _ExtentX        =   17066
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   24
         Top             =   750
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1530
            TabIndex        =   25
            Top             =   390
            Width           =   7995
            _Version        =   65536
            _ExtentX        =   14102
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1530
            TabIndex        =   26
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin VB.Label Label26 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label28 
            Caption         =   "Doc. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   60
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
            Left            =   690
            TabIndex        =   18
            Top             =   60
            Width           =   3885
            _Version        =   65536
            _ExtentX        =   6853
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Empresas"
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
            Picture         =   "EvaCre_frm_008.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1695
         Left            =   0
         TabIndex        =   19
         Top             =   1560
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1305
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1710
            TabIndex        =   20
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Emisión"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Evaluación"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   4110
            TabIndex        =   22
            Top             =   60
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación Obtenida"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   2910
            TabIndex        =   23
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Evaluación"
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   5085
         Left            =   30
         TabIndex        =   29
         Top             =   4110
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
         _ExtentY        =   8969
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
         Begin VB.ComboBox cmb_NumTra 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   900
            Width           =   7995
         End
         Begin VB.ComboBox cmb_Magnit 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1230
            Width           =   7995
         End
         Begin VB.ComboBox cmb_ClaSBS 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   3330
            Width           =   7995
         End
         Begin VB.TextBox txt_Observ 
            Height          =   885
            Left            =   1530
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "EvaCre_frm_008.frx":0316
            Top             =   3810
            Width           =   7995
         End
         Begin VB.ComboBox cmb_Clasif 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   4710
            Width           =   7995
         End
         Begin EditLib.fpDateTime ipp_FecRep 
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   2010
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   1
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_NumEva 
            Height          =   315
            Left            =   1530
            TabIndex        =   30
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
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
         End
         Begin Threed.SSPanel pnl_FecEmi 
            Height          =   315
            Left            =   1530
            TabIndex        =   31
            Top             =   390
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
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
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   90
            Left            =   30
            TabIndex        =   32
            Top             =   1590
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
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
         Begin EditLib.fpLongInteger ipp_NumEnt 
            Height          =   315
            Left            =   1530
            TabIndex        =   8
            Top             =   2340
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "99"
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
         Begin EditLib.fpDoubleSingle ipp_DeuMNc 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   2670
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            MinValue        =   "-9000000000"
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
         Begin EditLib.fpDoubleSingle ipp_DeuMEx 
            Height          =   315
            Left            =   1530
            TabIndex        =   10
            Top             =   3000
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            MinValue        =   "-9000000000"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   90
            Left            =   30
            TabIndex        =   33
            Top             =   3690
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   90
            Left            =   30
            TabIndex        =   34
            Top             =   750
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
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
         Begin VB.Label Label10 
            Caption         =   "Nro. Evaluación:"
            Height          =   285
            Left            =   90
            TabIndex        =   46
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Emisión:"
            Height          =   285
            Left            =   90
            TabIndex        =   45
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label38 
            Caption         =   "Nro. Trabajadores:"
            Height          =   285
            Left            =   90
            TabIndex        =   44
            Top             =   900
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Magnitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   43
            Top             =   1230
            Width           =   1665
         End
         Begin VB.Label Label4 
            Caption         =   "Registro SBS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   42
            Top             =   1710
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Reporte:"
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   2010
            Width           =   1425
         End
         Begin VB.Label Label6 
            Caption         =   "Nro. Entidades:"
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   2340
            Width           =   1395
         End
         Begin VB.Label Label41 
            Caption         =   "Importe MN:"
            Height          =   285
            Left            =   90
            TabIndex        =   39
            Top             =   2670
            Width           =   1665
         End
         Begin VB.Label Label7 
            Caption         =   "Importe ME:"
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   3000
            Width           =   1665
         End
         Begin VB.Label Label8 
            Caption         =   "Calificación SBS:"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   3330
            Width           =   1665
         End
         Begin VB.Label Label9 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   3810
            Width           =   1605
         End
         Begin VB.Label Label11 
            Caption         =   "Clasificación:"
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   4710
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   765
         Left            =   30
         TabIndex        =   47
         Top             =   9240
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
            Left            =   8160
            Picture         =   "EvaCre_frm_008.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   8880
            Picture         =   "EvaCre_frm_008.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   765
         Left            =   30
         TabIndex        =   48
         Top             =   3300
         Width           =   9585
         _Version        =   65536
         _ExtentX        =   16907
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
         Begin VB.CommandButton cmd_VerEva 
            Height          =   675
            Left            =   8190
            Picture         =   "EvaCre_frm_008.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Consultar Evaluación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RegEva 
            Height          =   675
            Left            =   7500
            Picture         =   "EvaCre_frm_008.frx":0D70
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Registrar Evaluación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   8880
            Picture         =   "EvaCre_frm_008.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueEva 
            Height          =   675
            Left            =   6810
            Picture         =   "EvaCre_frm_008.frx":14BC
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nueva Evaluación"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_EvaEmp_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgCon     As Integer

Private Sub cmb_ClaSbs_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_ClaSbs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaSbs_Click
   End If
End Sub

Private Sub cmb_Clasif_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Clasif_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Clasif_Click
   End If
End Sub

Private Sub cmb_Magnit_Click()
   Call gs_SetFocus(ipp_FecRep)
End Sub

Private Sub cmb_Magnit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Magnit_Click
   End If
End Sub

Private Sub cmb_NumTra_Click()
   Call gs_SetFocus(cmb_Magnit)
End Sub

Private Sub cmb_NumTra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NumTra_Click
   End If
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_NumTra.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Número de Trabajadores.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NumTra)
      Exit Sub
   End If

   If cmb_Magnit.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Magnitud.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Magnit)
      Exit Sub
   End If
   
   If cmb_ClaSbs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Calificación SBS.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaSbs)
      Exit Sub
   End If
   
   If cmb_Clasif.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Clasif)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de registrar la Evaluación?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_MODIFICA_EMP_DATEVA ("
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TDoEmp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NDoEmp & "', "
      g_str_Parame = g_str_Parame & CStr(CLng(pnl_NumEva.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_Clasif.ItemData(cmb_Clasif.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_NumTra.ItemData(cmb_NumTra.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Left(cmb_Magnit.Text, 1) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecRep.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumEnt.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DeuMNc.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DeuMEx.Value) & ", "
      g_str_Parame = g_str_Parame & Left(cmb_ClaSbs.Text, 1) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                        'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_EMP_DATEVA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   Screen.MousePointer = 0
   
   MsgBox "Evaluación registrada correctamente.", vbInformation, modgen_g_str_NomPlt
      
   Call fs_Limpia
   Call fs_Activa(True)
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_NueEva_Click()
   l_int_FlgCon = 1
   
   grd_Listad.Row = 0
   grd_Listad.Col = 2
   
   If Len(Trim(grd_Listad.Text)) = 0 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "Tiene una Evaluación Pendiente. No se puede generar una nueva evaluación.", vbInformation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   'Generar Nueva Evaluación
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_INSERTA_EMP_DATEVA ("
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TDoEmp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NDoEmp & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                        'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_EMP_DATEVA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   Screen.MousePointer = 0
   
   MsgBox "Se genero nueva Evaluación.", vbInformation, modgen_g_str_NomPlt
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_RegEva_Click()
   l_int_FlgCon = 1
   
   grd_Listad.Col = 2
   
   If Len(Trim(grd_Listad.Text)) > 0 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "Ya registro la Evaluación. No se puede modificar una Evaluación ya registrada.", vbInformation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If

   grd_Listad.Col = 0
   pnl_NumEva.Caption = grd_Listad.Text

   grd_Listad.Col = 1
   pnl_FecEmi.Caption = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_NumTra)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerEva_Click()
   grd_Listad.Col = 2
   
   If Len(Trim(grd_Listad.Text)) = 0 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se ha registrado la evaluación. No hay datos por consultar.", vbInformation, modgen_g_str_NomPlt
      
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If

   grd_Listad.Col = 0
   pnl_NumEva.Caption = grd_Listad.Text

   grd_Listad.Col = 1
   pnl_FecEmi.Caption = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM EMP_DATEVA WHERE "
   g_str_Parame = g_str_Parame & "DATEVA_EMPTDO = " & CStr(moddat_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATEVA_EMPNDO = '" & moddat_g_str_NDoEmp & "' AND "
   g_str_Parame = g_str_Parame & "DATEVA_NUMEVA = " & CStr(CLng(pnl_NumEva.Caption)) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Call gs_BuscarCombo_Item(cmb_NumTra, g_rst_Princi!DATEVA_NUMTRA)
   
   Call gs_BuscarCombo_Text(cmb_Magnit, g_rst_Princi!DATEVA_MAGSBS, 1)
   
   ipp_FecRep.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATEVA_FRGSBS))
   ipp_NumEnt.Value = g_rst_Princi!DATEVA_ENTREP
   ipp_DeuMNc.Value = g_rst_Princi!DATEVA_DEUMNC
   ipp_DeuMEx.Value = g_rst_Princi!DATEVA_DEUMEX
   
   Call gs_BuscarCombo_Text(cmb_ClaSbs, Format(g_rst_Princi!DATEVA_CLASBS, "0"), 1)
   
   txt_Observ.Text = Trim(g_rst_Princi!DATEVA_OBSERV)
   
   Call gs_BuscarCombo_Item(cmb_Clasif, g_rst_Princi!DATEVA_CLASIF)
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   l_int_FlgCon = 2
      
   grd_Listad.Enabled = False
   cmd_NueEva.Enabled = False
   cmd_RegEva.Enabled = False
   cmd_VerEva.Enabled = False
   txt_Observ.Enabled = True
   cmd_Cancel.Enabled = True
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call fs_Inicia
   
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call moddat_gf_RazSoc_NomCom
   
   pnl_DocIde.Caption = CStr(moddat_g_int_TDoEmp) & " - " & moddat_g_str_NDoEmp
   pnl_RazSoc.Caption = moddat_g_str_RazSoc
   
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt & " - Evaluación de Empresas"
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Magnit, 1, "054")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NumTra, 1, "017")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClaSbs, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Clasif, 1, "016")
   
   cmb_Clasif.RemoveItem cmb_Clasif.ListCount - 1
   
   grd_Listad.ColWidth(0) = 1620
   grd_Listad.ColWidth(1) = 1200
   grd_Listad.ColWidth(2) = 1200
   grd_Listad.ColWidth(3) = 5130
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   pnl_NumEva.Caption = ""
   pnl_FecEmi.Caption = ""
   
   cmb_NumTra.ListIndex = -1
   cmb_Magnit.ListIndex = -1
   
   ipp_FecRep.Text = Format(date, "dd/mm/yyyy")
   ipp_NumEnt.Value = 0
   ipp_DeuMNc.Value = 0
   ipp_DeuMEx.Value = 0
   
   txt_Observ.Text = ""
   
   cmb_ClaSbs.ListIndex = -1
   cmb_Clasif.ListIndex = -1
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   grd_Listad.Enabled = p_Habilita
   cmd_NueEva.Enabled = p_Habilita
   cmd_RegEva.Enabled = p_Habilita
   cmd_VerEva.Enabled = p_Habilita
   
   cmb_NumTra.Enabled = Not p_Habilita
   cmb_Magnit.Enabled = Not p_Habilita
   ipp_FecRep.Enabled = Not p_Habilita
   ipp_NumEnt.Enabled = Not p_Habilita
   ipp_DeuMNc.Enabled = Not p_Habilita
   ipp_DeuMEx.Enabled = Not p_Habilita
   cmb_ClaSbs.Enabled = Not p_Habilita
   txt_Observ.Enabled = Not p_Habilita
   cmb_Clasif.Enabled = Not p_Habilita
   
   cmd_Grabar.Enabled = Not p_Habilita
   cmd_Cancel.Enabled = Not p_Habilita
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM EMP_DATEVA WHERE "
   g_str_Parame = g_str_Parame & "DATEVA_EMPTDO = " & CStr(moddat_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATEVA_EMPNDO = '" & moddat_g_str_NDoEmp & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DATEVA_NUMEVA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontraron Evaluaciones registradas para esta Clasificación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!DATEVA_NUMEVA, "000000")
   
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATEVA_FECEMI))
   
      If g_rst_Princi!DATEVA_FECEVA > 0 Then
         grd_Listad.Col = 2
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATEVA_FECEVA))
         
         grd_Listad.Col = 3
         grd_Listad.Text = moddat_gf_Consulta_ParDes("016", CStr(g_rst_Princi!DATEVA_CLASIF))
      End If
   
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerEva_Click
End Sub

Private Sub ipp_DeuMEx_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ClaSbs)
   End If
End Sub

Private Sub ipp_DeuMNc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DeuMEx)
   End If
End Sub

Private Sub ipp_FecRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumEnt)
   End If
End Sub

Private Sub ipp_NumEnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DeuMNc)
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If l_int_FlgCon = 2 Then
      KeyAscii = 0
   Else
      If KeyAscii = 13 Then
         Call gs_SetFocus(cmb_Clasif)
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
      End If
   End If
End Sub

