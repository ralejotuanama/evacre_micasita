VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_53 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9210
   ClientLeft      =   7590
   ClientTop       =   1065
   ClientWidth     =   11250
   Icon            =   "EvaCre_frm_046.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9210
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   1
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
         Begin VB.ComboBox cmb_TipSeg 
            Height          =   315
            Left            =   2580
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1050
            Width           =   8535
         End
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   2580
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2580
            TabIndex        =   4
            Top             =   390
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            Left            =   2580
            TabIndex        =   5
            Top             =   720
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
            TabIndex        =   6
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
            TabIndex        =   7
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
         Begin VB.Label lbl_MonPre 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Left            =   2100
            TabIndex        =   33
            Top             =   1410
            Width           =   435
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   255
            Left            =   2100
            TabIndex        =   32
            Top             =   1740
            Width           =   435
         End
         Begin VB.Label Label27 
            Caption         =   "Monto Pr�stamo:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1815
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo (En A�os):"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Per�odo de Gracia:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo Seguro Desgraven:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   1050
            Width           =   2205
         End
         Begin VB.Label Label30 
            Caption         =   "Cuota Mensual (M. Prest.):"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   1380
            Width           =   2025
         End
         Begin VB.Label Label15 
            Caption         =   "Cuota Mensual:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   1710
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   4335
         Left            =   30
         TabIndex        =   14
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
            Height          =   4245
            Index           =   5
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   7488
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
         TabIndex        =   16
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
         Begin VB.CommandButton cmd_TipCam 
            Height          =   585
            Left            =   1230
            Picture         =   "EvaCre_frm_046.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Consulta Tipo de Cambio"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Calcul 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_046.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_046.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Simulaci�n de Cr�ditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Recalc 
            Height          =   585
            Left            =   2430
            Picture         =   "EvaCre_frm_046.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Actualizar C�lculo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "EvaCre_frm_046.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   1830
            Picture         =   "EvaCre_frm_046.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   21
            Top             =   60
            Width           =   1935
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   25
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   7710
            TabIndex        =   28
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   0
         TabIndex        =   29
         Top             =   0
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
            Left            =   720
            TabIndex        =   30
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cr�ditos Hipotecarios - Evaluaci�n Crediticia"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Left            =   720
            TabIndex        =   31
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "C�lculo del Pr�stamo"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "EvaCre_frm_046.frx":15F0
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_53"
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
Dim l_dbl_CuoApr     As Double
Dim l_dbl_PlzMax     As Double
Dim l_int_EdaMax     As Integer
Dim l_dbl_IngNet     As Double
Dim l_int_ComRta     As Integer
Dim l_int_GraMax     As Integer

Private Sub cmb_TipSeg_Click()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "

   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_TipSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipSeg_Click
   End If
End Sub

Private Sub cmd_Calcul_Click()
   Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\winnt\system32\calc.exe", vbNormalFocus)
   
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicaci�n", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_Ini_PlaMin    As Double
   Dim r_dbl_Ini_PlaMax    As Double
   Dim r_int_EdaAct        As Integer
   Dim r_int_EdaCli        As Integer
   Dim r_int_EdaCyg        As Integer

   If ipp_MtoPre.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Pr�stamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
      
      Exit Sub
   End If
   
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Par�metros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   
   'Calculando Edad del Cliente
   r_int_EdaCli = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), Date), 2))
   
   r_int_EdaCyg = 0
   If l_int_ComRta = 1 Then
      r_int_EdaCyg = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), Date), 2))
   End If
   
   If r_int_EdaCli > r_int_EdaCyg Then
      r_int_EdaAct = r_int_EdaCli
   Else
      r_int_EdaAct = r_int_EdaCyg
   End If
   
   'Validando Edad del Cliente + Plazo del Pr�stamo para Cobertura de Seguro
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > l_int_EdaMax Then
      MsgBox "La Edad del Cliente o de su C�nyuge m�s el Plazo del Pr�stamo excede el par�metro permitido. El Plazo m�ximo podr�a ser de " & CStr(l_int_EdaMax - r_int_EdaAct) & " a�os.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If cmb_TipSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipSeg)
      Exit Sub
   End If
   
   If l_int_ComRta = 1 Then
      If cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex) <> 12 Then
         MsgBox "El Tipo de Seguro debe ser Mancomunado porque el Cliente complementa renta con el C�nyuge.", vbExclamation, modgen_g_str_NomPlt
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
      MsgBox "La Cuota calculada es mayor a la Cuota Aprobada.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
   
      Exit Sub
   End If

   If MsgBox("�Est� seguro de registrar la Evaluaci�n?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVACRE_REGCAL ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoPre.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_PlaAno.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_PerGra.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TipCam) & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'C�digo Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'C�digo Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGrb = 2
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
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
   
   MsgBox "Evaluaci�n registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   
   moddat_g_int_FlgAct_1 = 2
   
   Unload Me
End Sub

Private Sub cmd_Recalc_Click()
   If ipp_MtoPre.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Pr�stamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoPre)
      
      Exit Sub
   End If
   
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Par�metros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If cmb_TipSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipSeg)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_CalCuo
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opci�n.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub cmd_TipCam_Click()
   frm_ConTCa_01.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecSol.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_DatCre
   
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad(5).ColWidth(0) = 3200
   grd_Listad(5).ColWidth(1) = 7940

   grd_Listad(5).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(5).ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad(5))

   'Plazo de Cr�dito
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

   'Obteniendo Edad M�xima permitida para el Cliente
   l_int_EdaMax = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      l_int_EdaMax = moddat_g_arr_Genera(1).Genera_Cantid
   End If
End Sub

Private Sub fs_Limpia()
   pnl_TipCam.Caption = "0.0000 "
   
   ipp_MtoPre.Value = 0
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_TipSeg.Clear
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
   
   'Obteniendo Tipo de Cambio de Moneda del Pr�stamo
   l_dbl_TipCam = 0
   
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
   End If
   
   lbl_MonPre.Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
End Sub

Private Sub fs_DatCre()
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
   grd_Listad(5).Text = "Tipo de Evaluaci�n"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Pr�stamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tasa de Inter�s"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Valor de Compra Venta"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Aporte Propio"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Monto Pr�stamo"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
      Else
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Valor de Compra Venta"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Aporte Propio"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Monto Pr�stamo"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
      End If
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Plazo (A�os)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Per�odo de Gracia (Meses)"
   
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
      grd_Listad(5).Text = "Compa��a de Seguros"
   
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
      grd_Listad(5).Text = "D�a de Pago"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Instituci�n Financiera de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto M�nimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!SOLMAE_MONAHO) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
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
   Else
      l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
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
   
   ipp_MtoPre.Value = l_dbl_MtoPre
   ipp_PlaAno.Value = l_int_PlaAno
   ipp_PerGra.Value = l_int_PerGra
   
   Call moddat_gs_Carga_TipSeg(cmb_TipSeg, l_str_EmpSeg)
   Call gs_BuscarCombo_Item(cmb_TipSeg, l_int_TipSeg)
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
   Dim r_dbl_CuoMen        As Double
   Dim r_dbl_SegViv        As Double
   Dim r_int_EdaAct        As Integer
   Dim r_int_EdaCli        As Integer
   Dim r_int_EdaCyg        As Integer
   Dim r_dbl_CuoFin        As Double
   Dim r_dbl_MtoPre_Max    As Double
   Dim r_dbl_CuoMpr_Max    As Double

   If ipp_MtoPre.Value = 0 Then
      Exit Sub
   End If
   
   If ipp_PlaAno.Value = 0 Then
      Exit Sub
   End If
   
   If cmb_TipSeg.ListIndex = -1 Then
      Exit Sub
   End If

   If l_dbl_CuoApr = 0 Then
      Exit Sub
   End If
   
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "

   'Determinando Plazo M�ximo para dar Cr�dito por Cobertura de Seguro
   r_int_EdaCli = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), Date), 2))
   
   r_int_EdaCyg = 0
   If l_int_ComRta = 1 Then
      r_int_EdaCyg = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), Date), 2))
   End If
   
   If r_int_EdaCli > r_int_EdaCyg Then
      r_int_EdaAct = r_int_EdaCli
   Else
      r_int_EdaAct = r_int_EdaCyg
   End If
   
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > l_int_EdaMax Then
      l_dbl_PlzMax = l_int_EdaMax - r_int_EdaAct
   End If
   
   'Obteniendo Tasa de Seguro de Desgravamen
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex), moddat_g_int_TipMon, ipp_MtoPre.Value, r_int_TipVal_Des, r_dbl_Import_Des)
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, 0, moddat_g_int_TipMon, l_dbl_ComVta, r_int_TipVal_Viv, r_dbl_Import_Viv)
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   If r_int_TipVal_Viv = 1 Then
      r_dbl_SegViv = r_dbl_Import_Viv / 100 * l_dbl_ComVta * 0.72
   Else
      r_dbl_SegViv = r_dbl_Import_Viv
   End If
   r_dbl_SegViv = CDbl(Format(r_dbl_SegViv, "###0.00"))

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
         
         Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), ipp_MtoPre.Value, r_dbl_PorCon, r_dbl_TopCon, l_dbl_TipCam, l_dbl_ComVta, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
         
      Case "002", "006"
         Call gs_Cronog_MiCasita(l_arr_CliNCo(), l_dbl_ComVta, ipp_MtoPre.Value, ipp_PlaAno.Value * 12, 2, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, ipp_PerGra.Value)
         
      Case "003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
   
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         
         Call gs_Cronog_CME_NC(l_arr_CliNCo(), ipp_MtoPre.Value, r_dbl_PorCon, r_dbl_TopCon, l_dbl_ComVta, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
         
      Case "004"
         r_dbl_TopCon = 0
      
         Call gs_Cronog_Mihogar_NC(l_arr_CliNCo(), ipp_MtoPre.Value, r_dbl_TopCon, l_dbl_ComVta, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
      Case "007", "009"
         r_dbl_TopCon = 0
      
         Call gs_Cronog_Mihogar_NC(l_arr_CliNCo(), ipp_MtoPre.Value, r_dbl_TopCon, l_dbl_ComVta, ipp_PlaAno.Value * 12, ipp_PerGra.Value, l_dbl_TasInt, r_dbl_Import_Des, 1, r_dbl_Import_Viv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), l_int_DiaPag, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   End Select
   
   pnl_CuoMPr_Cal.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo, "###,##0.00") & " "
   
   If moddat_g_int_TipMon = 1 Then
      pnl_CuoSol_Cal.Caption = pnl_CuoMPr_Cal.Caption
   Else
      pnl_CuoSol_Cal.Caption = Format(pnl_CuoMPr_Cal.Caption * l_dbl_TipCam, "###,##0.00") & " "
   End If
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
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 2
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).Text = "Total Ingreso Neto"
   
   grd_Listad(5).Col = 1
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).Text = "Cuota M�x. Aprob."
   
   grd_Listad(5).Col = 1
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).Text = "Cuota M�x. Aprob. (M. Prest.)"
   
   grd_Listad(5).Col = 1
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2)
   
   If moddat_g_int_TipMon <> 1 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).Text = "Tipo Cambio (C�lculo Ingresos)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
   End If
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).Text = "Fecha C�lculo de Ingresos"
   
   grd_Listad(5).Col = 1
   grd_Listad(5).CellForeColor = modgen_g_con_ColAzu
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_FECING))
   
   Call gs_UbiIniGrid(grd_Listad(5))
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      ipp_MtoPre.Value = g_rst_Princi!EVACRE_MTOPRE_CAL
      ipp_PlaAno.Value = g_rst_Princi!EVACRE_PLAANO_CAL
      ipp_PerGra.Value = g_rst_Princi!EVACRE_PERGRA_CAL
      
      Call gs_BuscarCombo_Item(cmb_TipSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_CalCuo
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

Private Sub ipp_PerGra_Change()
   pnl_CuoMPr_Cal.Caption = "0.00 "
   pnl_CuoSol_Cal.Caption = "0.00 "
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipSeg)
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




