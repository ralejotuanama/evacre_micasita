VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_65 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   9165
   ClientLeft      =   2010
   ClientTop       =   2505
   ClientWidth     =   11250
   Icon            =   "EvaCre_frm_072.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9210
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   16245
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
      Begin Threed.SSPanel SSPanel12 
         Height          =   2055
         Left            =   30
         TabIndex        =   9
         Top             =   7080
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   3625
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
         Begin VB.ComboBox cmb_TasEsp 
            Height          =   315
            Left            =   9420
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1050
            Width           =   1605
         End
         Begin VB.ComboBox cmb_CuoDbl 
            Height          =   315
            Left            =   6060
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   720
            Width           =   1905
         End
         Begin VB.ComboBox cmb_TipSeg 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   5355
         End
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   2610
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2610
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
            MaxValue        =   "20"
            MinValue        =   "5"
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
         Begin EditLib.fpLongInteger ipp_PerGra 
            Height          =   315
            Left            =   2610
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
            MaxValue        =   "12"
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
         Begin Threed.SSPanel pnl_CuoMPr_Cal 
            Height          =   315
            Left            =   2610
            TabIndex        =   10
            Top             =   1380
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
         Begin Threed.SSPanel pnl_CuoSol_Cal 
            Height          =   315
            Left            =   2610
            TabIndex        =   11
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
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Especial:"
            Height          =   195
            Left            =   8220
            TabIndex        =   38
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label17 
            Caption         =   "Cuotas Extraordinarias:"
            Height          =   285
            Left            =   4320
            TabIndex        =   37
            Top             =   750
            Width           =   1725
         End
         Begin VB.Label Label15 
            Caption         =   "Cuota Mensual:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   1740
            Width           =   1365
         End
         Begin VB.Label Label30 
            Caption         =   "Cuota Mensual (M. Prest.):"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   1410
            Width           =   2025
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo Seguro Desgraven:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   1080
            Width           =   2205
         End
         Begin VB.Label Label25 
            Caption         =   "Período de Gracia:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo (En Años):"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   1665
         End
         Begin VB.Label Label27 
            Caption         =   "Monto Préstamo:"
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   255
            Left            =   2100
            TabIndex        =   13
            Top             =   1740
            Width           =   435
         End
         Begin VB.Label lbl_MonPre 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Left            =   2100
            TabIndex        =   12
            Top             =   1410
            Width           =   435
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   4335
         Left            =   30
         TabIndex        =   20
         Top             =   2220
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   7646
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
            Height          =   4185
            Left            =   90
            TabIndex        =   21
            Top             =   90
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   7382
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   22
         Top             =   720
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
            Picture         =   "EvaCre_frm_072.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "EvaCre_frm_072.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Recalc 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_072.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Actualizar Cálculo"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   23
         Top             =   6600
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   2580
            TabIndex        =   24
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
         Begin VB.Label Label5 
            Caption         =   "Tipo Cambio (M. Prest.):"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   1935
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   26
         Top             =   1410
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   32
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   7710
            TabIndex        =   30
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   33
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
            TabIndex        =   34
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
            TabIndex        =   35
            Top             =   330
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Condiciones Crediticias"
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
            Picture         =   "EvaCre_frm_072.frx":0BA2
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_65"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String
Dim l_arr_CliNCo()      As modcal_g_est_CuoCli
Dim l_dbl_MtoPre        As Double
Dim l_dbl_ComVta        As Double
Dim l_int_TipEva        As Integer
Dim l_dbl_TasInt        As Double
Dim l_int_PlaAno        As Integer
Dim l_int_PerGra        As Integer
Dim l_int_CuoExt        As Integer
Dim l_str_EmpSeg        As String
Dim l_int_TipSeg        As Integer
Dim l_int_TasEsp        As Integer
Dim l_int_DiaPag        As Integer
Dim l_int_MonAho        As Integer
Dim l_dbl_MtoAho        As Double
Dim l_dbl_TipCam        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_PlzMax        As Double
Dim l_int_EdaMax        As Integer
Dim l_dbl_IngNet        As Double
Dim l_int_ComRta        As Integer
Dim l_int_GraMax        As Integer
Dim l_dbl_CuoIni        As Double
Dim l_str_CodCiu        As String

Private Sub cmd_Grabar_Click()
Dim r_dbl_Ini_PlaMin    As Double
Dim r_dbl_Ini_PlaMax    As Double
Dim r_int_EdaAct        As Integer
Dim r_int_EdaCli        As Integer
Dim r_int_EdaCyg        As Integer

   If ipp_MtoPre.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
      Exit Sub
   End If
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   'Calculando Edad del Cliente
   r_int_EdaCli = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), date), 2))
   
   r_int_EdaCyg = 0
   If l_int_ComRta = 1 Then
      r_int_EdaCyg = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), date), 2))
   End If
   If r_int_EdaCli > r_int_EdaCyg Then
      r_int_EdaAct = r_int_EdaCli
   Else
      r_int_EdaAct = r_int_EdaCyg
   End If
   
   'Validando Edad del Cliente + Plazo del Préstamo para Cobertura de Seguro
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > l_int_EdaMax Then
      MsgBox "La Edad del Cliente o de su Cónyuge más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(l_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If cmb_TipSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipSeg)
      Exit Sub
   End If
   If cmb_TasEsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar una tasa especial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TasEsp)
      Exit Sub
   End If
   If l_int_ComRta = 1 Then
      If cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex) <> 12 Then
         MsgBox "El Tipo de Seguro debe ser Mancomunado porque el Cliente complementa renta con el Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipSeg)
         Exit Sub
      End If
   End If
   
   If CDbl(pnl_CuoMPr_Cal.Caption) = 0 Then
      MsgBox "Debe calcular la Cuota Mensual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Recalc)
      Exit Sub
   End If
   
   If CDbl(pnl_CuoMPr_Cal.Caption) > l_dbl_CuoApr Then
      MsgBox moddat_g_str_Msje01, vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
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
      
      g_str_Parame = "USP_TRA_EVACRE_ACT_REGCAL ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoPre.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_PlaAno.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_PerGra.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TipCam) & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & CStr(cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex)) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
         
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
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 20, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 0
   MsgBox "Evaluación registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct_2 = 2
   Unload Me
End Sub

Private Sub cmd_Recalc_Click()
   If ipp_MtoPre.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
      Exit Sub
   End If
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If cmb_CuoDbl.ListIndex = -1 Then
      MsgBox "Debe seleccionar si desea cuotas extraordinarias (dobles).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CuoDbl)
      Exit Sub
   End If
   If cmb_TipSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipSeg)
      Exit Sub
   End If
   If cmb_TasEsp.ListIndex = -1 Then
      MsgBox "Debe seleccionar una tipo de tasa especial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TasEsp)
      Exit Sub
   End If
   
   If CInt(l_str_CodCiu) = 0 Then
      MsgBox "El codigo CIIU del cliente debe estar registrado, favor de coordinarlo con Comercial.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 11
   Call fs_CalCuo
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecSol.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Call fs_Limpia
   Call fs_Inicia
   Call fs_DatCre
   Call modmip_gs_DatCre(grd_Listad, r_arr_Mtz)
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   'inicializa variables
   l_dbl_MtoPre = 0
   l_dbl_ComVta = 0
   l_int_TipEva = 0
   l_dbl_TasInt = 0
   l_int_PlaAno = 0
   l_int_PerGra = 0
   l_int_CuoExt = 0
   l_str_EmpSeg = 0
   l_int_TipSeg = 0
   l_int_DiaPag = 0
   l_int_MonAho = 0
   l_dbl_MtoAho = 0
   l_dbl_TipCam = 0
   l_dbl_CuoApr = 0
   l_dbl_PlzMax = 0
   l_int_EdaMax = 0
   l_dbl_IngNet = 0
   l_int_ComRta = 0
   l_int_GraMax = 0
   l_dbl_CuoIni = 0
      
   'inicializa controles
   ipp_MtoPre.Value = 0
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_CuoDbl.ListIndex = -1
   cmb_TipSeg.Clear
   
   pnl_TipCam.Caption = "0.000000 "
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
   lbl_MonPre.Caption = " "
   
   'Obtiene el Tipo de Cambio
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   End If
   
   'Obtiene la Moneda del Préstamo
   lbl_MonPre.Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 3200
   grd_Listad.ColWidth(1) = 7940
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad)

   'Plazo de Crédito
   l_dbl_PlzMax = 0
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
      ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
      l_dbl_PlzMax = moddat_g_arr_Genera(1).Genera_PlzMax
   End If

   'Periodo de Gracia
   l_int_GraMax = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
      ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
      ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      l_int_GraMax = moddat_g_arr_Genera(1).Genera_ValMax
   End If

   'Obteniendo Edad Máxima permitida para el Cliente
   l_int_EdaMax = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      l_int_EdaMax = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Cuotas Extraordinarias (Dobles)
   Call moddat_gs_Carga_LisIte_Combo(cmb_CuoDbl, 1, "277")
   
   'Tasa Especial
   Call moddat_gs_Carga_LisIte_Combo(cmb_TasEsp, 1, "522")
End Sub

Private Sub fs_DatCre()
   'Buscar Fechas de Nacimiento de Cliente Titular
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      moddat_g_str_FecNac_Tit = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      
      l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
      If CInt(l_str_CodCiu) = 0 Then
         MsgBox "El codigo CIIU del titular debe estar registrado, favor de coordinarlo con Comercial.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscar Fechas de Nacimiento de Cliente Cónyuge
   If moddat_g_int_CygTDo > 0 Then
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         moddat_g_str_FecNac_Cyg = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
         If CInt(l_str_CodCiu) <> 7522 And CInt(l_str_CodCiu) <> 7523 Then
            If CInt(g_rst_Princi!DATGEN_CODCIU) > 0 Then
               l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
            End If
         End If
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   'Cargando Información de Solicitud de Crédito
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
   
   'carga variables
   If moddat_g_int_TipMon = 1 Then
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
      l_dbl_CuoIni = g_rst_Princi!SOLMAE_APOPRO_SOL - g_rst_Princi!SOLMAE_MTOGCI
   Else
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
      l_dbl_CuoIni = g_rst_Princi!SOLMAE_APOPRO_DOL - g_rst_Princi!SOLMAE_MTOGCI
   End If
   l_int_CuoExt = g_rst_Princi!SOLMAE_CUOEXT
   l_int_PlaAno = g_rst_Princi!SOLMAE_PLAANO
   l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
   l_int_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
   l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
   l_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
   l_int_TipEva = g_rst_Princi!SOLMAE_TIPEVA
   l_str_EmpSeg = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
   l_int_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
   l_int_TasEsp = g_rst_Princi!SOLMAE_TASESP
   l_int_MonAho = g_rst_Princi!SOLMAE_MONAHO
   l_dbl_MtoAho = g_rst_Princi!SOLMAE_MTOAHO
      
   'cierra recordset
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'carga controles
   ipp_MtoPre.Value = l_dbl_MtoPre
   ipp_PlaAno.Value = l_int_PlaAno
   ipp_PerGra.Value = l_int_PerGra
   Call gs_BuscarCombo_Item(cmb_CuoDbl, l_int_CuoExt)
   Call moddat_gs_Carga_TipSeg(cmb_TipSeg, l_str_EmpSeg)
   Call gs_BuscarCombo_Item(cmb_TipSeg, l_int_TipSeg)
   Call gs_BuscarCombo_Item(cmb_TasEsp, l_int_TasEsp)
End Sub

Private Sub fs_Buscar_DatEva()
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
   l_dbl_IngNet = g_rst_Princi!EVACRE_INGNET
   l_dbl_CuoApr = g_rst_Princi!EVACRE_CUOMPR
   
   l_int_ComRta = 2
   If g_rst_Princi!EVACRE_INGDP3 > 0 Or g_rst_Princi!EVACRE_INGAD2 > 0 Then
      l_int_ComRta = 1
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Total Ingreso Neto"
   grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:                 grd_Listad.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Cuota Máx. Aprob."
   grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:                 grd_Listad.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Cuota Máx. Aprob. (M. Prest.)"
   grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:                 grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2)
   
   If moddat_g_int_TipMon <> 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                       grd_Listad.Text = "Tipo Cambio (Cálculo Ingresos)"
      grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:              grd_Listad.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Fecha Cálculo de Ingresos"
   grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:                 grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_FECING))
   
   Call gs_UbiIniGrid(grd_Listad)
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      ipp_MtoPre.Value = g_rst_Princi!EVACRE_MTOPRE_CAL
      ipp_PlaAno.Value = g_rst_Princi!EVACRE_PLAANO_CAL
      ipp_PerGra.Value = g_rst_Princi!EVACRE_PERGRA_CAL
      Call gs_BuscarCombo_Item(cmb_CuoDbl, g_rst_Princi!EVACRE_CUODBL_CAL)
      Call gs_BuscarCombo_Item(cmb_TipSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_CalCuo
End Sub

Private Sub fs_CalCuo()
Dim r_dbl_IntGra        As Double
Dim r_dbl_Portes        As Double
Dim r_dbl_TasMVi        As Double
Dim r_dbl_ComCof        As Double
Dim r_dbl_TasCof        As Double
Dim r_int_TipVal_Des    As Integer
Dim r_dbl_Import_Des    As Double
Dim r_int_TipVal_Viv    As Integer
Dim r_dbl_Import_Viv    As Double
Dim r_dbl_PorCon        As Double
Dim r_dbl_TopCon        As Double
Dim r_dbl_MtoCon        As Double
Dim r_dbl_MtoNCo        As Double
Dim r_dbl_CuoMen        As Double
Dim r_int_EdaAct        As Integer
Dim r_int_EdaCli        As Integer
Dim r_int_EdaCyg        As Integer
Dim r_dbl_CuoFin        As Double
Dim r_dbl_MtoPre        As Double
Dim r_dbl_MtoPre_Max    As Double
Dim r_dbl_CuoMpr_Max    As Double

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

   If ipp_MtoPre.Value = 0 Then
      Exit Sub
   End If
   If ipp_PlaAno.Value = 0 Then
      Exit Sub
   End If
   If cmb_CuoDbl.ListIndex = -1 Then
      Exit Sub
   End If
   If cmb_TipSeg.ListIndex = -1 Then
      Exit Sub
   End If
   If cmb_TasEsp.ListIndex = -1 Then
      Exit Sub
   End If
   
   If l_dbl_CuoApr = 0 Then
      Exit Sub
   End If
   
   'inicializa cuotas calculadas
   r_dbl_TasMVi = 0
   r_dbl_ComCof = 0
   r_dbl_TasCof = 0
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
   
   'Determina tasa y comision de cofide
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
      r_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   End If
   
   'Determina Plazo Máximo para dar Crédito por Cobertura de Seguro
   r_int_EdaCli = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), date), 2))
   r_int_EdaCyg = 0
   If l_int_ComRta = 1 Then
      r_int_EdaCyg = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), date), 2))
   End If
   If r_int_EdaCli > r_int_EdaCyg Then
      r_int_EdaAct = r_int_EdaCli
   Else
      r_int_EdaAct = r_int_EdaCyg
   End If
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > l_int_EdaMax Then
      l_dbl_PlzMax = l_int_EdaMax - r_int_EdaAct
   End If
      
   'Obtiene Tasa de Seguro de Desgravamen y Vivienda
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex), moddat_g_int_TipMon, ipp_MtoPre.Value, r_int_TipVal_Des, r_dbl_Import_Des, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, 0, moddat_g_int_TipMon, l_dbl_ComVta, r_int_TipVal_Viv, r_dbl_Import_Viv, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   
   'Obtiene portes
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Obtiene monto del prestamo
   r_dbl_MtoPre = l_dbl_ComVta - l_dbl_CuoIni
   If CDbl(ipp_MtoPre.Text) <> r_dbl_MtoPre Then
      l_dbl_CuoIni = l_dbl_ComVta - CDbl(ipp_MtoPre.Text)
   End If
   
   Select Case moddat_g_str_CodPrd > 0
      'SE COMENTA ESTA PARTE DEL CODIGO PORQUE EL PRODUCTO ESTA DESCONTINUADO
      'Case InStr(moddat_g_str_Agr1CRC, moddat_g_str_CodPrd)
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
      '   Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), ipp_MtoPre.Value, r_dbl_PorCon, r_dbl_TopCon, l_dbl_TipCam, l_dbl_ComVta, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
      
      'SE COMENTA ESTA PARTE DEL CODIGO PORQUE EL PRODUCTO ESTA DESCONTINUADO
      'Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)
      ''Para obtener porcentaje de TC
      'r_dbl_PorCon = 0
      'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
      '   r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
      'End If
      '
      ''Para obtener tope de TC
      'r_dbl_TopCon = 0
      'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
      '   r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      'End If
      '
      ''NUEVA rutina de generacion de cronogramas
      'int_Produc = 1
      'int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
      'dbl_ValInm = l_dbl_ComVta
      'dbl_CuoIni = l_dbl_CuoIni
      'dbl_MtoCon = ipp_MtoPre.Value * (r_dbl_PorCon / 100)
      'If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
      'int_PlaPre = CInt(ipp_PlaAno.Text) * 12
      'dbl_TasInt = l_dbl_TasInt
      'dbl_TasCof = r_dbl_TasCof
      'dbl_ComCof = r_dbl_ComCof
      'dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
      'int_DiaVct = l_int_DiaPag
      'int_PerGra = CInt(ipp_PerGra.Text)
      'str_PriVct = ""
      'dbl_Portes = r_dbl_Portes
      'dbl_SegViv = r_dbl_Import_Viv
      'int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
      'dbl_SegDes = r_dbl_Import_Des
      '
      ''Calculando cronogramas
      'Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
      'Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
      '
      'dbl_CuoMen = 0
      'dbl_CuoPbp = 0
      'dbl_IngReq = 0
      'Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
      '
      ''Muestra valor de la cuota
      'pnl_CuoMPr_Cal.Caption = Format(dbl_CuoPbp, "###,##0.00") & " "
      'If moddat_g_int_TipMon = 1 Then
      '   pnl_CuoSol_Cal.Caption = pnl_CuoMPr_Cal.Caption
      'Else
      '   pnl_CuoSol_Cal.Caption = Format(pnl_CuoMPr_Cal.Caption * l_dbl_TipCam, "###,##0.00") & " "
      'End If
      
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_CuoIni
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor de la cuota
         pnl_CuoMPr_Cal.Caption = Format(dbl_CuoMen, "###,##0.00") & " "
         If moddat_g_int_TipMon = 1 Then
            pnl_CuoSol_Cal.Caption = pnl_CuoMPr_Cal.Caption
         Else
            pnl_CuoSol_Cal.Caption = Format(pnl_CuoMPr_Cal.Caption * l_dbl_TipCam, "###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)  '"004", "006", "007", "009", "010", "013", "014", "015", "016", "017", "018"
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(l_dbl_ComVta) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_CuoIni
         dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'Muestra valor de la cuota
         pnl_CuoMPr_Cal.Caption = Format(dbl_CuoPbp, "###,##0.00") & " "
         If moddat_g_int_TipMon = 1 Then
            pnl_CuoSol_Cal.Caption = pnl_CuoMPr_Cal.Caption
         Else
            pnl_CuoSol_Cal.Caption = Format(pnl_CuoMPr_Cal.Caption * l_dbl_TipCam, "###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = l_dbl_ComVta
         dbl_CuoIni = l_dbl_CuoIni
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         pnl_CuoMPr_Cal.Caption = Format(dbl_CuoMen, "###,##0.00") & " "
         If moddat_g_int_TipMon = 1 Then
            pnl_CuoSol_Cal.Caption = pnl_CuoMPr_Cal.Caption
         Else
            pnl_CuoSol_Cal.Caption = Format(pnl_CuoMPr_Cal.Caption * l_dbl_TipCam, "###,##0.00") & " "
         End If
   End Select
   
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_MtoPre_Change()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_PlaAno_Change()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CuoDbl)
   End If
End Sub

Private Sub cmb_CuoDbl_Click()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
   Call gs_SetFocus(cmb_TipSeg)
End Sub

Private Sub cmb_CuoDbl_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_CuoDbl.ListIndex > -1 Then
         Call gs_SetFocus(cmb_TipSeg)
      End If
   End If
End Sub

Private Sub cmb_TipSeg_Click()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
   Call gs_SetFocus(cmb_TasEsp)
End Sub

Private Sub cmb_TipSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipSeg.ListIndex > -1 Then
         Call gs_SetFocus(cmb_TasEsp)
      End If
   End If
End Sub

Private Sub cmb_TasEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub


