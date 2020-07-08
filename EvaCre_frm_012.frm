VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   1800
   ClientTop       =   1050
   ClientWidth     =   11655
   Icon            =   "EvaCre_frm_012.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   16325
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   1095
         Left            =   30
         TabIndex        =   83
         Top             =   7320
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.ComboBox cmb_AccNue 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   720
            Width           =   9555
         End
         Begin VB.ComboBox cmb_RelLab 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   390
            Width           =   9555
         End
         Begin VB.TextBox txt_CodSbs 
            Height          =   315
            Left            =   1950
            MaxLength       =   12
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label34 
            Caption         =   "Vínculo Accionista:"
            Height          =   315
            Left            =   90
            TabIndex        =   86
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label33 
            Caption         =   "Vínculo Laboral:"
            Height          =   315
            Left            =   90
            TabIndex        =   85
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label32 
            Caption         =   "Código Deudor SBS:"
            Height          =   285
            Left            =   90
            TabIndex        =   84
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2085
         Left            =   30
         TabIndex        =   71
         Top             =   5190
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo 
            Height          =   315
            Left            =   1950
            MaxLength       =   12
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1950
            TabIndex        =   32
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8190
            TabIndex        =   31
            Text            =   "cmb_PrvDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1950
            TabIndex        =   30
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   8190
            MaxLength       =   15
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1950
            MaxLength       =   120
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6210
            TabIndex        =   81
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label27 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   90
            TabIndex        =   80
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   90
            TabIndex        =   79
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6210
            TabIndex        =   78
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   90
            TabIndex        =   77
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6210
            TabIndex        =   76
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   90
            TabIndex        =   75
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6210
            TabIndex        =   74
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   90
            TabIndex        =   73
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   72
            Top             =   60
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   45
         Top             =   8460
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_LisRec 
            Height          =   675
            Left            =   30
            Picture         =   "EvaCre_frm_012.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LisOpe 
            Height          =   675
            Left            =   720
            Picture         =   "EvaCre_frm_012.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatCyg 
            Height          =   675
            Left            =   9450
            Picture         =   "EvaCre_frm_012.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "EvaCre_frm_012.frx":0797
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10140
            Picture         =   "EvaCre_frm_012.frx":0BD9
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   675
            Left            =   8760
            Picture         =   "EvaCre_frm_012.frx":101B
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4395
         Left            =   30
         TabIndex        =   46
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   7752
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
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   4020
            Width           =   3315
         End
         Begin VB.CheckBox chk_DirEle 
            Caption         =   "Autoriz. Corresp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9900
            TabIndex        =   16
            Top             =   3420
            Width           =   1485
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   3360
            Width           =   1665
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   1950
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   3360
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   8190
            TabIndex        =   13
            Text            =   "cmb_Profes"
            Top             =   3030
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3030
            Width           =   3315
         End
         Begin VB.ComboBox cmb_RegCyg 
            Height          =   315
            Left            =   8190
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_EstCiv 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   9
            Text            =   "cmb_DstNac"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   1950
            TabIndex        =   8
            Text            =   "cmb_PrvNac"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   7
            Text            =   "cmb_DptNac"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   1950
            TabIndex        =   6
            Text            =   "cmb_Paises"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodSex 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin EditLib.fpLongInteger ipp_DepEc1 
            Height          =   315
            Left            =   8190
            TabIndex        =   18
            Top             =   3690
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   1950
            TabIndex        =   5
            Top             =   1710
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin EditLib.fpLongInteger ipp_DepEc2 
            Height          =   315
            Left            =   8820
            TabIndex        =   19
            Top             =   3690
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
         Begin EditLib.fpLongInteger ipp_DepEc3 
            Height          =   315
            Left            =   9480
            TabIndex        =   20
            Top             =   3690
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
         Begin EditLib.fpLongInteger ipp_DepEc4 
            Height          =   315
            Left            =   10170
            TabIndex        =   21
            Top             =   3720
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
         Begin EditLib.fpLongInteger ipp_DepEc5 
            Height          =   315
            Left            =   10800
            TabIndex        =   22
            Top             =   3720
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
         Begin EditLib.fpLongInteger ipp_NumDep 
            Height          =   315
            Left            =   1950
            TabIndex        =   17
            Top             =   3690
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
         Begin Threed.SSPanel pnl_EdaCli 
            Height          =   315
            Left            =   3300
            TabIndex        =   47
            Top             =   1710
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin Threed.SSPanel pnl_MesCli 
            Height          =   315
            Left            =   4230
            TabIndex        =   48
            Top             =   1710
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1950
            TabIndex        =   88
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
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
            Caption         =   "Registra Activ. Econom.:"
            Height          =   315
            Left            =   90
            TabIndex        =   87
            Top             =   4020
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Documento Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   82
            Top             =   60
            Width           =   1725
         End
         Begin VB.Label Label38 
            Caption         =   "Nro. Depend. Econom.:"
            Height          =   285
            Left            =   90
            TabIndex        =   68
            Top             =   3690
            Width           =   2055
         End
         Begin VB.Label Label18 
            Caption         =   "Edades Depend. Econom.:"
            Height          =   285
            Left            =   6210
            TabIndex        =   67
            Top             =   3690
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6210
            TabIndex        =   66
            Top             =   3360
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   90
            TabIndex        =   65
            Top             =   3360
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión:"
            Height          =   315
            Left            =   6210
            TabIndex        =   64
            Top             =   3030
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   90
            TabIndex        =   63
            Top             =   3030
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Régimen Conyugal:"
            Height          =   315
            Left            =   6210
            TabIndex        =   62
            Top             =   2700
            Width           =   1905
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   2700
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   60
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   59
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   58
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   57
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label6 
            Caption         =   "Sexo:"
            Height          =   315
            Left            =   90
            TabIndex        =   55
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6210
            TabIndex        =   51
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label30 
            Caption         =   "Años"
            Height          =   285
            Left            =   3840
            TabIndex        =   50
            Top             =   1770
            Width           =   555
         End
         Begin VB.Label Label31 
            Caption         =   "Meses"
            Height          =   285
            Left            =   4770
            TabIndex        =   49
            Top             =   1770
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   69
         Top             =   30
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
            TabIndex        =   70
            Top             =   60
            Width           =   3255
            _Version        =   65536
            _ExtentX        =   5741
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
            Picture         =   "EvaCre_frm_012.frx":1325
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Profes()      As moddat_tpo_Genera
Dim l_arr_Paises()      As moddat_tpo_Genera
Dim l_int_FlgCmb        As Integer
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_str_DptNac        As String
Dim l_str_PrvNac        As String
Dim l_str_DstNac        As String
Dim l_str_Paises        As String
Dim l_str_Profes        As String

Private Sub cmb_ActEco_Click()
   Call gs_SetFocus(cmb_TipVia)
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
   End If
End Sub

Private Sub cmd_ActEco_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   modatecli_g_int_Tip_ActEco = 1
   
   frm_MntCli_04.Show 1
End Sub

Private Sub cmd_DatCyg_Click()
   frm_MntCli_03.Show 1
End Sub

Private Sub cmd_LisOpe_Click()
   frm_LisOpe_01.Show 1
End Sub

Private Sub cmd_LisRec_Click()
   frm_LisRec_01.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " - Mantenimiento de Clientes"
   
   Call fs_Inicio
   Call fs_Limpia
   
   Call fs_Buscar
   
   Call gs_SetFocus(txt_ApePat)

   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodSex, 1, "207")
   Call moddat_gs_Carga_LisIte_Combo(cmb_EstCiv, 1, "205")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RegCyg, 1, "206")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_AccNue, 1, "052")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RelLab, 1, "053")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "214")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
      
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub

Private Sub fs_Buscar()
   Dim r_int_FlgFun     As Integer

   'Inicializando y/o Limpiando Arreglos
   Call modatecli_gs_Limpia_DatGen(1)     'Cliente Titular - Datos Generales
   Call modatecli_gs_Limpia_DatGen(2)     'Cónyuge - Datos Generales
   
   ReDim modatecli_g_arr_Tit_ActEco(0)    'Cliente Titular - Datos Económicos
   ReDim modatecli_g_arr_Cyg_ActEco(0)    'Cónyuge - Datos Económicos

   'Datos Actividades Económicas Cliente Titular
   modatecli_g_str_CodCiu_Tit = ""
   modatecli_g_str_GirCom_Tit = ""
   modatecli_g_str_SecEco_Tit = ""
   modatecli_g_int_TDoEmp_Tit = 0
   modatecli_g_str_NDoEmp_Tit = ""
   modatecli_g_int_ActPri_Tit = 0
   modatecli_g_int_ActSec_Tit = 0
   
   'Datos Actividades Económicas Cliente Cónyuge
   modatecli_g_str_CodCiu_Cyg = ""
   modatecli_g_str_GirCom_Cyg = ""
   modatecli_g_str_SecEco_Cyg = ""
   modatecli_g_int_TDoEmp_Cyg = 0
   modatecli_g_str_NDoEmp_Cyg = ""
   modatecli_g_int_ActSec_Cyg = 0
   
   atecli_int_CliReg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Titular)
   atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
   'Inicializando Arreglos de Solicitudes Rechazadas
   ReDim modatecli_g_arr_LisRec(0)
   ReDim modatecli_g_arr_CygRec(0)

   'Inicializando Arreglos de Operaciones Vigentes
   ReDim modatecli_g_arr_TitOpe(0)
   ReDim modatecli_g_arr_CygOpe(0)

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_ActEcoTit = 1
   modatecli_g_int_CygDatGen = 1
   
   'Inicializando Variables de DOI Cónyuge
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   'Mostrando Mensaje si Cliente tiene Solicitud de Crédito en Evaluación
   r_int_FlgFun = atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
   r_int_FlgFun = atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
   
   'Buscando Solicitudes Rechazadas
   Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_LisRec) > 0 Then
      cmd_LisRec.Enabled = True
   End If
   
   'Buscando Operaciones
   Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_TitOpe) > 0 Then
      cmd_LisOpe.Enabled = True
   End If
   
   'Buscando Información de Cliente Titular
   Call atecli_gs_Buscar_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   moddat_g_str_NomCli = Trim(modatecli_g_arr_DatGen(1).DatGen_ApePat) & " " & Trim(modatecli_g_arr_DatGen(1).DatGen_ApeMat) & " " & Trim(modatecli_g_arr_DatGen(1).DatGen_Nombre)
   
   
   'Si el Titular está registrado como Casado buscar información del Cónyuge
   If moddat_g_int_CygTDo > 0 Then
      Call atecli_gs_Buscar_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      moddat_g_str_CygNom = Trim(modatecli_g_arr_DatGen(2).DatGen_ApePat) & " " & Trim(modatecli_g_arr_DatGen(2).DatGen_ApeMat) & " " & Trim(modatecli_g_arr_DatGen(2).DatGen_Nombre)
   End If
   
   'Asignado DOI a Controles
   pnl_DocIde.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc
   
   'Si se encontro Cliente en Base de Datos Asignar Información de Cliente Titular a Controles
   If atecli_int_CliReg = 2 Then
      Call fs_Arreglo_DatCli
   End If
End Sub

Private Sub fs_Arreglo_DatCli()
   txt_ApePat.Text = modatecli_g_arr_DatGen(1).DatGen_ApePat
   txt_ApeMat.Text = modatecli_g_arr_DatGen(1).DatGen_ApeMat
   txt_Nombre.Text = modatecli_g_arr_DatGen(1).DatGen_Nombre
   
   Call gs_BuscarCombo_Item(cmb_CodSex, modatecli_g_arr_DatGen(1).DatGen_CodSex)
   ipp_FecNac.Text = modatecli_g_arr_DatGen(1).DatGen_FecNac
   
   cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises(), modatecli_g_arr_DatGen(1).DatGen_Paises) - 1
   
   If modatecli_g_arr_DatGen(1).DatGen_Paises = "004028" Then
      cmb_DptNac.Enabled = True
      cmb_PrvNac.Enabled = True
      cmb_DstNac.Enabled = True
   
      Call gs_BuscarCombo_Item(cmb_DptNac, modatecli_g_arr_DatGen(1).DatGen_DptNac)
   
      Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(modatecli_g_arr_DatGen(1).DatGen_DptNac, "00"))
      Call gs_BuscarCombo_Item(cmb_PrvNac, modatecli_g_arr_DatGen(1).DatGen_PrvNac)

      Call moddat_gs_Carga_Distri(cmb_DstNac, Format(modatecli_g_arr_DatGen(1).DatGen_DptNac, "00"), Format(modatecli_g_arr_DatGen(1).DatGen_PrvNac, "00"))
      Call gs_BuscarCombo_Item(cmb_DstNac, modatecli_g_arr_DatGen(1).DatGen_DstNac)
   End If
   
   Call gs_BuscarCombo_Item(cmb_EstCiv, modatecli_g_arr_DatGen(1).DatGen_EstCiv)
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      cmb_RegCyg.Enabled = True
      Call gs_BuscarCombo_Item(cmb_RegCyg, modatecli_g_arr_DatGen(1).DatGen_RegCyg)
   End If

   Call gs_BuscarCombo_Item(cmb_NivEst, modatecli_g_arr_DatGen(1).DatGen_NivEst)
   
   cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, modatecli_g_arr_DatGen(1).DatGen_Profes) - 1
   
   txt_Celula.Text = modatecli_g_arr_DatGen(1).DatGen_Celula
   txt_DirEle.Text = modatecli_g_arr_DatGen(1).DatGen_DirEle
      
   If modatecli_g_arr_DatGen(1).DatGen_Autori = 1 Then
      chk_DirEle.Value = 1
   End If
      
   ipp_NumDep.Value = modatecli_g_arr_DatGen(1).DatGen_DepEco
      
   If ipp_NumDep.Value > 0 Then
      ipp_DepEc1.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 1, 3))
      ipp_DepEc2.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 4, 3))
      ipp_DepEc3.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 7, 3))
      ipp_DepEc4.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 10, 3))
      ipp_DepEc5.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 13, 3))
   End If
   
   Call gs_BuscarCombo_Item(cmb_ActEco, modatecli_g_arr_DatGen(1).DatGen_ActEco)
   
   Call gs_BuscarCombo_Item(cmb_TipVia, modatecli_g_arr_DatGen(1).DatGen_TipVia)
   txt_NomVia.Text = modatecli_g_arr_DatGen(1).DatGen_NomVia
   txt_Numero.Text = modatecli_g_arr_DatGen(1).DatGen_Numero
   txt_Interi.Text = modatecli_g_arr_DatGen(1).DatGen_IntDpt
   
   Call gs_BuscarCombo_Item(cmb_TipZon, modatecli_g_arr_DatGen(1).DatGen_TipZon)
   txt_NomZon.Text = modatecli_g_arr_DatGen(1).DatGen_NomZon
   
   Call gs_BuscarCombo_Item(cmb_DptDir, modatecli_g_arr_DatGen(1).DatGen_DptDir)

   Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(modatecli_g_arr_DatGen(1).DatGen_DptDir, "00"))
   Call gs_BuscarCombo_Item(cmb_PrvDir, modatecli_g_arr_DatGen(1).DatGen_PrvDir)

   Call moddat_gs_Carga_Distri(cmb_DstDir, Format(modatecli_g_arr_DatGen(1).DatGen_DptDir, "00"), Format(modatecli_g_arr_DatGen(1).DatGen_PrvDir, "00"))
   Call gs_BuscarCombo_Item(cmb_DstDir, modatecli_g_arr_DatGen(1).DatGen_DstDir)
   
   txt_Refere.Text = modatecli_g_arr_DatGen(1).DatGen_Refere
   txt_Telefo.Text = modatecli_g_arr_DatGen(1).DatGen_Telefo
      
   txt_CodSbs.Text = modatecli_g_arr_DatGen(1).DatGen_CodSbs
   
   Call gs_BuscarCombo_Text(cmb_RelLab, modatecli_g_arr_DatGen(1).DatGen_RelLab, 1)
   Call gs_BuscarCombo_Text(cmb_AccNue, modatecli_g_arr_DatGen(1).DatGen_AccNue, 1)
      
   'Obteniendo DNI del Cónyuge o Conviviente
   If modatecli_g_arr_DatGen(1).DatGen_CygTDo > 0 Then
      moddat_g_int_CygTDo = modatecli_g_arr_DatGen(1).DatGen_CygTDo
      moddat_g_str_CygNDo = modatecli_g_arr_DatGen(1).DatGen_CygNDo
      
      cmd_DatCyg.Enabled = True
   End If
End Sub

Private Sub fs_Limpia()
   Dim r_int_Contad  As Integer
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   cmb_CodSex.ListIndex = -1
   
   ipp_FecNac.Text = Format(CDate(moddat_g_str_FecSis) - CDate(18 * 365.25), "dd/mm/yyyy")
   pnl_EdaCli.Caption = "0 "
   pnl_MesCli.Caption = "0 "
   
   cmb_Paises.ListIndex = -1
   cmb_DptNac.ListIndex = -1
   cmb_PrvNac.Clear
   cmb_DstNac.Clear
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_EstCiv.ListIndex = -1
   cmb_RegCyg.ListIndex = -1
   cmb_RegCyg.Enabled = False
   cmd_DatCyg.Enabled = False
   cmb_NivEst.ListIndex = -1
   cmb_Profes.ListIndex = -1
   txt_DirEle.Text = ""
   txt_Celula.Text = ""
   
   ipp_NumDep.Value = 0
   ipp_DepEc1.Value = 0
   ipp_DepEc2.Value = 0
   ipp_DepEc3.Value = 0
   ipp_DepEc4.Value = 0
   ipp_DepEc5.Value = 0
   
   ipp_DepEc1.Enabled = False
   ipp_DepEc2.Enabled = False
   ipp_DepEc3.Enabled = False
   ipp_DepEc4.Enabled = False
   ipp_DepEc5.Enabled = False
   
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
   
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_Numero.Text = ""
   txt_Interi.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Telefo.Text = ""
   
   txt_CodSbs.Text = ""
   cmb_RelLab.ListIndex = -1
   cmb_AccNue.ListIndex = -1
   
   cmd_LisRec.Enabled = False
   cmd_LisOpe.Enabled = False
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeCas_GotFocus()
   Call gs_SelecTodo(txt_ApeCas)
End Sub

Private Sub txt_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Celula_GotFocus()
   Call gs_SelecTodo(txt_Celula)
End Sub

Private Sub txt_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DirEle_Change()
   If Len(Trim(txt_DirEle)) > 0 Then
      chk_DirEle.Enabled = True
   Else
      chk_DirEle.Value = 0
      chk_DirEle.Enabled = False
   End If
End Sub

Private Sub txt_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Telefo)
End Sub

Private Sub txt_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CodSbs)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodSex)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumDep)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Numero_GotFocus()
   Call gs_SelecTodo(txt_Numero)
End Sub

Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Interi_GotFocus()
   Call gs_SelecTodo(txt_Interi)
End Sub

Private Sub txt_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_Profes_Change()
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_Click()
   If cmb_Profes.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Celula)
      End If
   End If
End Sub

Private Sub cmb_Profes_GotFocus()
   l_int_FlgCmb = True
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Profes, l_str_Profes)
      l_int_FlgCmb = True
      
      If cmb_Profes.ListIndex > -1 Then
         l_str_Profes = ""
      End If
      
      Call gs_SetFocus(txt_Celula)
   End If
End Sub

Private Sub cmb_CodSex_Click()
   Call gs_SetFocus(ipp_FecNac)
End Sub

Private Sub cmb_CodSex_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodSex_Click
   End If
End Sub

Private Sub cmb_DptNac_Change()
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_Click()
   If cmb_DptNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvNac)
      End If
   End If
End Sub

Private Sub cmb_DptNac_GotFocus()
   l_int_FlgCmb = True
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptNac, l_str_DptNac)
      l_int_FlgCmb = True
      
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      If cmb_DptNac.ListIndex > -1 Then
         l_str_DptNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvNac)
   End If
End Sub

Private Sub cmb_EstCiv_Click()
   cmb_RegCyg.Enabled = False
   Call gs_SetFocus(cmb_NivEst)
   
   If cmb_EstCiv.ListIndex > -1 Then
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         cmb_RegCyg.Enabled = True
         Call gs_SetFocus(cmb_RegCyg)
      Else
         cmb_RegCyg.ListIndex = -1
      End If
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         cmd_DatCyg.Enabled = True
      Else
         Call modatecli_gs_Limpia_DatGen(2)
         ReDim modatecli_g_arr_CygActEco(0)
      
         'Datos Actividades Económicas Cliente Cónyuge
         modatecli_g_str_CodCiu_Cyg = ""
         modatecli_g_str_GirCom_Cyg = ""
         modatecli_g_str_SecEco_Cyg = ""
         modatecli_g_int_TDoEmp_Cyg = 0
         modatecli_g_str_NDoEmp_Cyg = ""
         modatecli_g_int_ActSec_Cyg = 0
   
         atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
         'Inicializando Arreglos de Solicitudes Rechazadas
         ReDim modatecli_g_arr_CygRec(0)

         'Inicializando Arreglos de Operaciones Vigentes
         ReDim modatecli_g_arr_CygOpe(0)

         'Inicializando Flag de Datos Ingresados
         modatecli_g_int_CygDatGen = 1
   
         'Inicializando Variables de DOI Cónyuge
         moddat_g_int_CygTDo = 0
         moddat_g_str_CygNDo = ""
         moddat_g_str_CygNom = ""
         
         cmd_DatCyg.Enabled = False
      End If
   End If
End Sub

Private Sub cmb_EstCiv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EstCiv_Click
   End If
End Sub

Private Sub cmb_NivEst_Click()
   Call gs_SetFocus(cmb_Profes)
End Sub

Private Sub cmb_NivEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivEst_Click
   End If
End Sub

Private Sub cmb_PrvNac_Change()
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_Click()
   If cmb_PrvNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstNac)
      End If
   End If
End Sub

Private Sub cmb_PrvNac_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvNac, l_str_PrvNac)
      l_int_FlgCmb = True
      
      cmb_DstNac.Clear
      If cmb_PrvNac.ListIndex > -1 Then
         l_str_DstNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstNac)
   End If
End Sub

Private Sub cmb_DstNac_Change()
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_Click()
   If cmb_DstNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_DstNac_GotFocus()
   l_int_FlgCmb = True
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstNac, l_str_DstNac)
      l_int_FlgCmb = True
      
      If cmb_DstNac.ListIndex > -1 Then
         l_str_DstNac = ""
      End If
      
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_Paises_Change()
   l_str_Paises = cmb_Paises.Text
End Sub

Private Sub cmb_Paises_Click()
   If cmb_Paises.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
         
         Call gs_SetFocus(cmb_DptNac)
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         End If
      End If
   Else
      cmb_DptNac.ListIndex = -1
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      
      cmb_DptNac.Enabled = False
      cmb_PrvNac.Enabled = False
      cmb_DstNac.Enabled = False
   
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_Paises_GotFocus()
   l_int_FlgCmb = True
   l_str_Paises = cmb_Paises.Text
End Sub

Private Sub cmb_Paises_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Paises, l_str_Paises)
      l_int_FlgCmb = True
      
      cmb_DptNac.Enabled = True
      cmb_PrvNac.Enabled = True
      cmb_DstNac.Enabled = True

      Call gs_SetFocus(cmb_DptNac)
      
      If cmb_Paises.ListIndex > -1 Then
         l_str_Paises = ""
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         End If
      Else
         cmb_DptNac.ListIndex = -1
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
      
         cmb_DptNac.Enabled = False
         cmb_PrvNac.Enabled = False
         cmb_DstNac.Enabled = False
   
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_RegCyg_Click()
   Call gs_SetFocus(cmb_NivEst)
End Sub

Private Sub cmb_RegCyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RegCyg_Click
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub ipp_DepEc1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc2.Enabled Then
         Call gs_SetFocus(ipp_DepEc2)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc3.Enabled Then
         Call gs_SetFocus(ipp_DepEc3)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc4.Enabled Then
         Call gs_SetFocus(ipp_DepEc4)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc5.Enabled Then
         Call gs_SetFocus(ipp_DepEc5)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ActEco)
   End If
End Sub

Private Sub ipp_FecNac_Change()
   pnl_EdaCli.Caption = Left(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
   pnl_MesCli.Caption = Right(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
   End If
End Sub

Private Sub ipp_NumDep_Change()
   If ipp_NumDep.Value = 0 Then
      ipp_DepEc1.Enabled = False
      ipp_DepEc2.Enabled = False
      ipp_DepEc3.Enabled = False
      ipp_DepEc4.Enabled = False
      ipp_DepEc5.Enabled = False
      
      ipp_DepEc1.Value = 0
   Else
      ipp_DepEc1.Enabled = True
      ipp_DepEc2.Enabled = True
      ipp_DepEc3.Enabled = True
      ipp_DepEc4.Enabled = True
      ipp_DepEc5.Enabled = True
      
      If ipp_NumDep.Value < 5 Then
         ipp_DepEc5.Enabled = False
         ipp_DepEc5.Value = 0
      End If
      
      If ipp_NumDep.Value < 4 Then
         ipp_DepEc4.Enabled = False
         ipp_DepEc4.Value = 0
      End If
      
      If ipp_NumDep.Value < 3 Then
         ipp_DepEc3.Enabled = False
         ipp_DepEc3.Value = 0
      End If
      
      If ipp_NumDep.Value < 2 Then
         ipp_DepEc2.Enabled = False
         ipp_DepEc2.Value = 0
      End If
   End If
End Sub

Private Sub ipp_NumDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_NumDep.Value > 0 Then
         Call gs_SetFocus(ipp_DepEc1)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Function ff_Valida() As Integer
   ff_Valida = False
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_ApePat)
      Exit Function
   End If
   
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_ApeMat)
      Exit Function
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Nombre)
      Exit Function
   End If
   
   If cmb_CodSex.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sexo.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_CodSex)
      Exit Function
   Else
      modatecli_g_arr_DatGen(1).DatGen_CodSex = cmb_CodSex.ItemData(cmb_CodSex.ListIndex)
   End If
   
   If Not IsDate(ipp_FecNac.Text) Then
      MsgBox "La Fecha de Nacimiento no es válida.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If CDate(ipp_FecNac.Text) > Date Then
      MsgBox "Debe ingresar una Fecha de Nacimiento valida.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If

   Call moddat_gs_FecSis
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Paises)
      Exit Function
   End If
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
      If cmb_DptNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de Nacimiento.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_DptNac)
         Exit Function
      End If
      
      If cmb_PrvNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de Nacimiento.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_PrvNac)
         Exit Function
      End If
      
      If cmb_DstNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de Nacimiento.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_DstNac)
         Exit Function
      End If
   End If
   
   If cmb_EstCiv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Estado Civil.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_EstCiv)
      Exit Function
   End If
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      If cmb_RegCyg.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Régimen Conyugal.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(cmb_RegCyg)
         Exit Function
      End If
   End If
   
   If cmb_NivEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Estudio.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_NivEst)
      Exit Function
   End If
   
   If cmb_Profes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Profesión u Oficio.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_Profes)
      Exit Function
   End If
   
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipVia)
      Exit Function
   End If
   
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Numero.Text)) = 0 Then
      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_Numero)
      Exit Function
   End If
   
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_TipZon)
      Exit Function
   End If
   
   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_NomZon.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_con_AteCli
         Call gs_SetFocus(txt_NomZon)
         Exit Function
      End If
   End If
   
   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_DptDir)
      Exit Function
   End If
   
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_PrvDir)
      Exit Function
   End If
   
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(cmb_DstDir)
      Exit Function
   End If
   
   moddat_g_str_NomCli = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   
   moddat_g_int_EdaAno = CInt(pnl_EdaCli.Caption)
   moddat_g_int_EdaMes = CInt(pnl_MesCli.Caption)
   
   ff_Valida = True
End Function

