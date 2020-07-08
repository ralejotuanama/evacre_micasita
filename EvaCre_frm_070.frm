VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_EvaCre_63 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9135
   ClientLeft      =   3060
   ClientTop       =   1380
   ClientWidth     =   12690
   Icon            =   "EvaCre_frm_070.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9225
      Left            =   0
      TabIndex        =   40
      Top             =   -90
      Width           =   12705
      _Version        =   65536
      _ExtentX        =   22410
      _ExtentY        =   16272
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
         Height          =   1155
         Left            =   30
         TabIndex        =   89
         Top             =   2250
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
         Begin VB.TextBox txt_ObsVDm 
            Height          =   705
            Left            =   2400
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Text            =   "EvaCre_frm_070.frx":000C
            Top             =   390
            Width           =   10155
         End
         Begin VB.ComboBox cmb_TipVDm 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   10185
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones Verificación Domiciliaria:"
            Height          =   495
            Index           =   4
            Left            =   60
            TabIndex        =   91
            Top             =   390
            Width           =   2175
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo Verificación Domiciliaria:"
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   90
            Top             =   60
            Width           =   2175
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4845
         Left            =   0
         TabIndex        =   87
         Top             =   3450
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
         _ExtentY        =   8546
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
            Height          =   4755
            Left            =   30
            TabIndex        =   88
            Top             =   60
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   8387
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Titular"
            TabPicture(0)   =   "EvaCre_frm_070.frx":0010
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel9"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel7"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel10"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_070.frx":002C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel3"
            Tab(1).Control(1)=   "SSPanel11"
            Tab(1).Control(2)=   "SSPanel12"
            Tab(1).ControlCount=   3
            Begin Threed.SSPanel SSPanel10 
               Height          =   765
               Left            =   30
               TabIndex        =   107
               Top             =   3900
               Width           =   12435
               _Version        =   65536
               _ExtentX        =   21934
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
               Begin VB.TextBox txt_Tit_CRiCom 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   36
                  Text            =   "EvaCre_frm_070.frx":0048
                  Top             =   60
                  Width           =   10035
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Reporte de Central de Riesgos:"
                  Height          =   495
                  Index           =   16
                  Left            =   60
                  TabIndex        =   108
                  Top             =   60
                  Width           =   2175
               End
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   2355
               Left            =   30
               TabIndex        =   101
               Top             =   1500
               Width           =   12435
               _Version        =   65536
               _ExtentX        =   21934
               _ExtentY        =   4154
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
               Begin VB.ComboBox cmb_Tit_ClaEn6 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   1980
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Tit_CodEn6 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   1980
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_CodEn1 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   12
                  Top             =   330
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_ClaEn1 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   330
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Tit_CodEn2 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   660
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_ClaEn2 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   660
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Tit_CodEn3 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   990
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_ClaEn3 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   990
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Tit_CodEn4 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   24
                  Top             =   1320
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_ClaEn4 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   27
                  Top             =   1320
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Tit_CodEn5 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   1650
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Tit_ClaEn5 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   1650
                  Width           =   3135
               End
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn1 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   13
                  Top             =   330
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
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn2 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   17
                  Top             =   660
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
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn3 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   21
                  Top             =   990
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
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn4 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   25
                  Top             =   1320
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
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn5 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   29
                  Top             =   1650
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
               Begin EditLib.fpDoubleSingle ipp_Tit_DeuEn6 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   33
                  Top             =   1980
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn1 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   14
                  Top             =   330
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn2 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   18
                  Top             =   660
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn3 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   22
                  Top             =   990
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn4 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   26
                  Top             =   1320
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn5 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   30
                  Top             =   1650
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
               Begin EditLib.fpDoubleSingle ipp_Tit_LimDeuEn6 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   34
                  Top             =   1980
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
                  AutoSize        =   -1  'True
                  Caption         =   "Clasificación"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   10120
                  TabIndex        =   129
                  Top             =   90
                  Width           =   1080
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Línea Asignada"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7815
                  TabIndex        =   128
                  Top             =   90
                  Width           =   1335
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Línea Utilizada"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   6390
                  TabIndex        =   127
                  Top             =   90
                  Width           =   1290
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Entidad"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   3950
                  TabIndex        =   126
                  Top             =   90
                  Width           =   660
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 6:"
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   109
                  Top             =   2010
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 1:"
                  Height          =   195
                  Index           =   10
                  Left            =   60
                  TabIndex        =   106
                  Top             =   360
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 2:"
                  Height          =   195
                  Index           =   12
                  Left            =   60
                  TabIndex        =   105
                  Top             =   690
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 3:"
                  Height          =   195
                  Index           =   13
                  Left            =   60
                  TabIndex        =   104
                  Top             =   1020
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 4:"
                  Height          =   195
                  Index           =   14
                  Left            =   60
                  TabIndex        =   103
                  Top             =   1350
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 5:"
                  Height          =   195
                  Index           =   15
                  Left            =   60
                  TabIndex        =   102
                  Top             =   1680
                  Width           =   1590
               End
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   1095
               Left            =   30
               TabIndex        =   92
               Top             =   360
               Width           =   12435
               _Version        =   65536
               _ExtentX        =   21934
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
               Begin VB.ComboBox cmb_Tit_CRiFlg 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   60
                  Width           =   735
               End
               Begin EditLib.fpDateTime ipp_Tit_CRiFec 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   3
                  Top             =   60
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
               Begin EditLib.fpLongInteger ipp_Tit_CRiEnt 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   4
                  Top             =   390
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
               Begin EditLib.fpDoubleSingle ipp_Tit_TotDMN 
                  Height          =   315
                  Left            =   2340
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
               Begin EditLib.fpDoubleSingle ipp_Tit_TotDME 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   11
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CRiCl0 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   5
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
                  Text            =   "999.99"
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CRiCl1 
                  Height          =   315
                  Left            =   9180
                  TabIndex        =   6
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CRiCl2 
                  Height          =   315
                  Left            =   9840
                  TabIndex        =   7
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CRiCl3 
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   8
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CRiCl4 
                  Height          =   315
                  Left            =   11160
                  TabIndex        =   9
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
                  AutoSize        =   -1  'True
                  Caption         =   "Clasificación:"
                  Height          =   195
                  Index           =   9
                  Left            =   6300
                  TabIndex        =   98
                  Top             =   450
                  Width           =   930
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Endeudamiento ME:"
                  Height          =   195
                  Index           =   8
                  Left            =   6300
                  TabIndex        =   97
                  Top             =   780
                  Width           =   1845
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Total Endeudamiento MN:"
                  Height          =   315
                  Index           =   7
                  Left            =   60
                  TabIndex        =   96
                  Top             =   720
                  Width           =   2115
               End
               Begin VB.Label Label38 
                  Caption         =   "Nro. Entidades Reportadas:"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   95
                  Top             =   390
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Presenta Información:"
                  Height          =   315
                  Index           =   6
                  Left            =   60
                  TabIndex        =   94
                  Top             =   60
                  Width           =   1815
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Reporte:"
                  Height          =   195
                  Left            =   6300
                  TabIndex        =   93
                  Top             =   120
                  Width           =   1110
               End
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   765
               Left            =   -74970
               TabIndex        =   110
               Top             =   3900
               Width           =   12405
               _Version        =   65536
               _ExtentX        =   21881
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
               Begin VB.TextBox txt_Cyg_CRiCom 
                  Height          =   645
                  Left            =   2340
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   76
                  Text            =   "EvaCre_frm_070.frx":004C
                  Top             =   60
                  Width           =   10005
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Comentarios sobre Reporte de Central de Riesgos:"
                  Height          =   495
                  Index           =   2
                  Left            =   60
                  TabIndex        =   111
                  Top             =   60
                  Width           =   2175
               End
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   2325
               Left            =   -74970
               TabIndex        =   112
               Top             =   1500
               Width           =   12405
               _Version        =   65536
               _ExtentX        =   21881
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
               Begin VB.ComboBox cmb_Cyg_ClaEn5 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   71
                  Top             =   1650
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Cyg_CodEn5 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   68
                  Top             =   1650
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_ClaEn4 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   1320
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Cyg_CodEn4 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   64
                  Top             =   1320
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_ClaEn3 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   63
                  Top             =   990
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Cyg_CodEn3 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   990
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_ClaEn2 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   59
                  Top             =   660
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Cyg_CodEn2 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   56
                  Top             =   660
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_ClaEn1 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   55
                  Top             =   330
                  Width           =   3135
               End
               Begin VB.ComboBox cmb_Cyg_CodEn1 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   330
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_CodEn6 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   72
                  Top             =   1980
                  Width           =   3975
               End
               Begin VB.ComboBox cmb_Cyg_ClaEn6 
                  Height          =   315
                  Left            =   9240
                  Style           =   2  'Dropdown List
                  TabIndex        =   75
                  Top             =   1980
                  Width           =   3135
               End
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn1 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   53
                  Top             =   330
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn2 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   57
                  Top             =   660
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn3 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   61
                  Top             =   990
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn4 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   65
                  Top             =   1320
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn5 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   69
                  Top             =   1650
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_DeuEn6 
                  Height          =   315
                  Left            =   6330
                  TabIndex        =   73
                  Top             =   1980
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn1 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   54
                  Top             =   330
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn2 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   58
                  Top             =   660
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn3 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   62
                  Top             =   990
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn4 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   66
                  Top             =   1320
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn5 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   70
                  Top             =   1650
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_LimDeuEn6 
                  Height          =   315
                  Left            =   7780
                  TabIndex        =   74
                  Top             =   1980
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
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Entidad"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   3950
                  TabIndex        =   133
                  Top             =   90
                  Width           =   660
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Línea Utilizada"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   6390
                  TabIndex        =   132
                  Top             =   90
                  Width           =   1290
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Línea Asignada"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7810
                  TabIndex        =   131
                  Top             =   90
                  Width           =   1335
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Clasificación"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   10120
                  TabIndex        =   130
                  Top             =   90
                  Width           =   1080
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 5:"
                  Height          =   195
                  Index           =   20
                  Left            =   60
                  TabIndex        =   118
                  Top             =   1680
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 4:"
                  Height          =   195
                  Index           =   19
                  Left            =   60
                  TabIndex        =   117
                  Top             =   1350
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 3:"
                  Height          =   195
                  Index           =   18
                  Left            =   60
                  TabIndex        =   116
                  Top             =   1020
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 2:"
                  Height          =   195
                  Index           =   17
                  Left            =   60
                  TabIndex        =   115
                  Top             =   690
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 1:"
                  Height          =   195
                  Index           =   11
                  Left            =   60
                  TabIndex        =   114
                  Top             =   360
                  Width           =   1590
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Deuda Entidad Nro. 6:"
                  Height          =   195
                  Index           =   5
                  Left            =   60
                  TabIndex        =   113
                  Top             =   2010
                  Width           =   1590
               End
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   1095
               Left            =   -74970
               TabIndex        =   119
               Top             =   360
               Width           =   12405
               _Version        =   65536
               _ExtentX        =   21881
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
               Begin VB.ComboBox cmb_Cyg_CRiFlg 
                  Height          =   315
                  Left            =   2340
                  Style           =   2  'Dropdown List
                  TabIndex        =   42
                  Top             =   60
                  Width           =   735
               End
               Begin EditLib.fpDateTime ipp_Cyg_CRiFec 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   43
                  Top             =   60
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
               Begin EditLib.fpLongInteger ipp_Cyg_CRiEnt 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   44
                  Top             =   390
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_TotDMN 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   50
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_TotDME 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   51
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_CRiCl0 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   45
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
                  Text            =   "999.99"
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_CRiCl1 
                  Height          =   315
                  Left            =   9180
                  TabIndex        =   46
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_CRiCl2 
                  Height          =   315
                  Left            =   9840
                  TabIndex        =   47
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_CRiCl3 
                  Height          =   315
                  Left            =   10500
                  TabIndex        =   48
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
               Begin EditLib.fpDoubleSingle ipp_Cyg_CRiCl4 
                  Height          =   315
                  Left            =   11160
                  TabIndex        =   49
                  Top             =   390
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Reporte:"
                  Height          =   195
                  Left            =   6300
                  TabIndex        =   125
                  Top             =   120
                  Width           =   1110
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Presenta Información:"
                  Height          =   315
                  Index           =   24
                  Left            =   60
                  TabIndex        =   124
                  Top             =   60
                  Width           =   1815
               End
               Begin VB.Label Label3 
                  Caption         =   "Nro. Entidades Reportadas:"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   123
                  Top             =   390
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Total Endeudamiento MN:"
                  Height          =   315
                  Index           =   23
                  Left            =   60
                  TabIndex        =   122
                  Top             =   720
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Endeudamiento ME:"
                  Height          =   195
                  Index           =   22
                  Left            =   6300
                  TabIndex        =   121
                  Top             =   780
                  Width           =   1845
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Clasificación:"
                  Height          =   195
                  Index           =   21
                  Left            =   6300
                  TabIndex        =   120
                  Top             =   450
                  Width           =   930
               End
            End
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   41
         Top             =   1440
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            TabIndex        =   77
            Top             =   390
            Width           =   11115
            _Version        =   65536
            _ExtentX        =   19606
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
            TabIndex        =   78
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
            Left            =   10500
            TabIndex        =   79
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
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
            Left            =   8760
            TabIndex        =   82
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   81
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   80
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   83
         Top             =   30
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            TabIndex        =   85
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
            TabIndex        =   86
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Verificaciones Personales"
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
            Picture         =   "EvaCre_frm_070.frx":0050
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   84
         Top             =   750
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
            Left            =   12010
            Picture         =   "EvaCre_frm_070.frx":035A
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_070.frx":079C
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   99
         Top             =   8310
         Width           =   12615
         _Version        =   65536
         _ExtentX        =   22251
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
         Begin VB.TextBox txt_RefPer 
            Height          =   645
            Left            =   2340
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Text            =   "EvaCre_frm_070.frx":0BDE
            Top             =   60
            Width           =   10155
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Comentarios sobre Referencias Personales:"
            Height          =   495
            Index           =   1
            Left            =   60
            TabIndex        =   100
            Top             =   60
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_63"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_DatCyg        As Integer

Dim l_arr_Tit_CodEn1()  As moddat_tpo_Genera
Dim l_arr_Tit_CodEn2()  As moddat_tpo_Genera
Dim l_arr_Tit_CodEn3()  As moddat_tpo_Genera
Dim l_arr_Tit_CodEn4()  As moddat_tpo_Genera
Dim l_arr_Tit_CodEn5()  As moddat_tpo_Genera
Dim l_arr_Tit_CodEn6()  As moddat_tpo_Genera

Dim l_arr_Cyg_CodEn1()  As moddat_tpo_Genera
Dim l_arr_Cyg_CodEn2()  As moddat_tpo_Genera
Dim l_arr_Cyg_CodEn3()  As moddat_tpo_Genera
Dim l_arr_Cyg_CodEn4()  As moddat_tpo_Genera
Dim l_arr_Cyg_CodEn5()  As moddat_tpo_Genera
Dim l_arr_Cyg_CodEn6()  As moddat_tpo_Genera

Private Sub cmb_Cyg_CodEn1_Click()
   If cmb_Cyg_CodEn1.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn1)
   End If
End Sub

Private Sub cmb_Cyg_CodEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn1_Click
   End If
End Sub

Private Sub cmb_Cyg_CodEn2_Click()
   If cmb_Cyg_CodEn2.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn2)
   End If
End Sub

Private Sub cmb_Cyg_CodEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn2_Click
   End If
End Sub

Private Sub cmb_Cyg_CodEn3_Click()
   If cmb_Cyg_CodEn3.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn3)
   End If
End Sub

Private Sub cmb_Cyg_CodEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn3_Click
   End If
End Sub

Private Sub cmb_Cyg_CodEn4_Click()
   If cmb_Cyg_CodEn4.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn4)
   End If
End Sub

Private Sub cmb_Cyg_CodEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn4_Click
   End If
End Sub

Private Sub cmb_Cyg_CodEn5_Click()
   If cmb_Cyg_CodEn5.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn5)
   End If
End Sub

Private Sub cmb_Cyg_CodEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn5_Click
   End If
End Sub

Private Sub cmb_Cyg_CodEn6_Click()
   If cmb_Cyg_CodEn6.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Cyg_DeuEn6)
   End If
End Sub

Private Sub cmb_Cyg_CodEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CodEn6_Click
   End If
End Sub

Private Sub cmb_TipVDm_Click()
   Call gs_SetFocus(txt_ObsVDm)
End Sub

Private Sub cmb_TipVDm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipVDm_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn1_Click()
   If cmb_Tit_CodEn2.Enabled Then
      Call gs_SetFocus(cmb_Tit_CodEn2)
   Else
      Call gs_SetFocus(txt_Tit_CRiCom)
   End If
End Sub

Private Sub cmb_Tit_ClaEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn1_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn2_Click()
   If cmb_Tit_CodEn3.Enabled Then
      Call gs_SetFocus(cmb_Tit_CodEn3)
   Else
      Call gs_SetFocus(txt_Tit_CRiCom)
   End If
End Sub

Private Sub cmb_Tit_ClaEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn2_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn3_Click()
   If cmb_Tit_CodEn4.Enabled Then
      Call gs_SetFocus(cmb_Tit_CodEn4)
   Else
      Call gs_SetFocus(txt_Tit_CRiCom)
   End If
End Sub

Private Sub cmb_Tit_ClaEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn3_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn4_Click()
   If cmb_Tit_CodEn5.Enabled Then
      Call gs_SetFocus(cmb_Tit_CodEn5)
   Else
      Call gs_SetFocus(txt_Tit_CRiCom)
   End If
End Sub

Private Sub cmb_Tit_ClaEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn4_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn5_Click()
   If cmb_Tit_CodEn6.Enabled Then
      Call gs_SetFocus(cmb_Tit_CodEn6)
   Else
      Call gs_SetFocus(txt_Tit_CRiCom)
   End If
End Sub

Private Sub cmb_Tit_ClaEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn5_Click
   End If
End Sub

Private Sub cmb_Tit_ClaEn6_Click()
   Call gs_SetFocus(txt_Tit_CRiCom)
End Sub

Private Sub cmb_Tit_ClaEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_ClaEn6_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn1_Click()
   If cmb_Tit_CodEn1.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn1)
   End If
End Sub

Private Sub cmb_Tit_CodEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn1_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn2_Click()
   If cmb_Tit_CodEn2.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn2)
   End If
End Sub

Private Sub cmb_Tit_CodEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn2_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn3_Click()
   If cmb_Tit_CodEn3.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn3)
   End If
End Sub

Private Sub cmb_Tit_CodEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn3_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn4_Click()
   If cmb_Tit_CodEn4.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn4)
   End If
End Sub

Private Sub cmb_Tit_CodEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn4_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn5_Click()
   If cmb_Tit_CodEn5.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn5)
   End If
End Sub

Private Sub cmb_Tit_CodEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn5_Click
   End If
End Sub

Private Sub cmb_Tit_CodEn6_Click()
   If cmb_Tit_CodEn6.ListIndex > -1 Then
      Call gs_SetFocus(ipp_Tit_DeuEn6)
   End If
End Sub

Private Sub cmb_Tit_CodEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CodEn6_Click
   End If
End Sub

Private Sub cmb_Tit_CRiFlg_Click()
   Call gs_SetFocus(ipp_Tit_CRiFec)
   
   If cmb_Tit_CRiFlg.ListIndex > -1 Then
      If cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex) = 2 Then
         ipp_Tit_CRiEnt.Value = 0
         ipp_Tit_CRiCl0.Value = 0
         ipp_Tit_CRiCl1.Value = 0
         ipp_Tit_CRiCl2.Value = 0
         ipp_Tit_CRiCl3.Value = 0
         ipp_Tit_CRiCl4.Value = 0
         ipp_Tit_TotDMN.Value = 0
         ipp_Tit_TotDME.Value = 0
         
         cmb_Tit_CodEn1.ListIndex = -1
         ipp_Tit_DeuEn1.Value = 0
         ipp_Tit_LimDeuEn1.Value = 0
         cmb_Tit_ClaEn1.ListIndex = -1
         
         cmb_Tit_CodEn2.ListIndex = -1
         ipp_Tit_DeuEn2.Value = 0
         ipp_Tit_LimDeuEn2.Value = 0
         cmb_Tit_ClaEn2.ListIndex = -1
         
         cmb_Tit_CodEn3.ListIndex = -1
         ipp_Tit_DeuEn3.Value = 0
         ipp_Tit_LimDeuEn3.Value = 0
         cmb_Tit_ClaEn3.ListIndex = -1
         
         cmb_Tit_CodEn4.ListIndex = -1
         ipp_Tit_DeuEn4.Value = 0
         ipp_Tit_LimDeuEn4.Value = 0
         cmb_Tit_ClaEn4.ListIndex = -1
         
         cmb_Tit_CodEn5.ListIndex = -1
         ipp_Tit_DeuEn5.Value = 0
         ipp_Tit_LimDeuEn5.Value = 0
         cmb_Tit_ClaEn5.ListIndex = -1
         
         cmb_Tit_CodEn6.ListIndex = -1
         ipp_Tit_DeuEn6.Value = 0
         ipp_Tit_LimDeuEn6.Value = 0
         cmb_Tit_ClaEn6.ListIndex = -1
      
         ipp_Tit_CRiEnt.Enabled = False
         ipp_Tit_CRiCl0.Enabled = False
         ipp_Tit_CRiCl1.Enabled = False
         ipp_Tit_CRiCl2.Enabled = False
         ipp_Tit_CRiCl3.Enabled = False
         ipp_Tit_CRiCl4.Enabled = False
         ipp_Tit_TotDMN.Enabled = False
         ipp_Tit_TotDME.Enabled = False
         cmb_Tit_CodEn1.Enabled = False
         ipp_Tit_DeuEn1.Enabled = False
         ipp_Tit_LimDeuEn1.Enabled = False
         cmb_Tit_ClaEn1.Enabled = False
         cmb_Tit_CodEn2.Enabled = False
         ipp_Tit_DeuEn2.Enabled = False
         ipp_Tit_LimDeuEn2.Enabled = False
         cmb_Tit_ClaEn2.Enabled = False
         cmb_Tit_CodEn3.Enabled = False
         ipp_Tit_DeuEn3.Enabled = False
         ipp_Tit_LimDeuEn3.Enabled = False
         cmb_Tit_ClaEn3.Enabled = False
         cmb_Tit_CodEn4.Enabled = False
         ipp_Tit_DeuEn4.Enabled = False
         ipp_Tit_LimDeuEn4.Enabled = False
         cmb_Tit_ClaEn4.Enabled = False
         cmb_Tit_CodEn5.Enabled = False
         ipp_Tit_DeuEn5.Enabled = False
         ipp_Tit_LimDeuEn5.Enabled = False
         cmb_Tit_ClaEn5.Enabled = False
         cmb_Tit_CodEn6.Enabled = False
         ipp_Tit_DeuEn6.Enabled = False
         ipp_Tit_LimDeuEn6.Enabled = False
         cmb_Tit_ClaEn6.Enabled = False
      Else
         ipp_Tit_CRiEnt.Enabled = True
         ipp_Tit_CRiCl0.Enabled = True
         ipp_Tit_CRiCl1.Enabled = True
         ipp_Tit_CRiCl2.Enabled = True
         ipp_Tit_CRiCl3.Enabled = True
         ipp_Tit_CRiCl4.Enabled = True
         ipp_Tit_TotDMN.Enabled = True
         ipp_Tit_TotDME.Enabled = True
      End If
   End If
End Sub

Private Sub cmb_Tit_CRiFlg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tit_CRiFlg_Click
   End If
End Sub

Private Sub cmb_Cyg_CRiFlg_Click()
   Call gs_SetFocus(ipp_Cyg_CRiFec)
   
   If cmb_Cyg_CRiFlg.ListIndex > -1 Then
      If cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex) = 2 Then
         ipp_Cyg_CRiEnt.Value = 0
         ipp_Cyg_CRiCl0.Value = 0
         ipp_Cyg_CRiCl1.Value = 0
         ipp_Cyg_CRiCl2.Value = 0
         ipp_Cyg_CRiCl3.Value = 0
         ipp_Cyg_CRiCl4.Value = 0
         ipp_Cyg_TotDMN.Value = 0
         ipp_Cyg_TotDME.Value = 0
         
         cmb_Cyg_CodEn1.ListIndex = -1
         ipp_Cyg_DeuEn1.Value = 0
         ipp_Cyg_LimDeuEn1.Value = 0
         cmb_Cyg_ClaEn1.ListIndex = -1
         
         cmb_Cyg_CodEn2.ListIndex = -1
         ipp_Cyg_DeuEn2.Value = 0
         ipp_Cyg_LimDeuEn2.Value = 0
         cmb_Cyg_ClaEn2.ListIndex = -1
         
         cmb_Cyg_CodEn3.ListIndex = -1
         ipp_Cyg_DeuEn3.Value = 0
         ipp_Cyg_LimDeuEn3.Value = 0
         cmb_Cyg_ClaEn3.ListIndex = -1
         
         cmb_Cyg_CodEn4.ListIndex = -1
         ipp_Cyg_DeuEn4.Value = 0
         ipp_Cyg_LimDeuEn4.Value = 0
         cmb_Cyg_ClaEn4.ListIndex = -1
         
         cmb_Cyg_CodEn5.ListIndex = -1
         ipp_Cyg_DeuEn5.Value = 0
         ipp_Cyg_LimDeuEn5.Value = 0
         cmb_Cyg_ClaEn5.ListIndex = -1
         
         cmb_Cyg_CodEn6.ListIndex = -1
         ipp_Cyg_DeuEn6.Value = 0
         ipp_Cyg_LimDeuEn6.Value = 0
         cmb_Cyg_ClaEn6.ListIndex = -1
      
         ipp_Cyg_CRiEnt.Enabled = False
         ipp_Cyg_CRiCl0.Enabled = False
         ipp_Cyg_CRiCl1.Enabled = False
         ipp_Cyg_CRiCl2.Enabled = False
         ipp_Cyg_CRiCl3.Enabled = False
         ipp_Cyg_CRiCl4.Enabled = False
         ipp_Cyg_TotDMN.Enabled = False
         ipp_Cyg_TotDME.Enabled = False
         cmb_Cyg_CodEn1.Enabled = False
         ipp_Cyg_DeuEn1.Enabled = False
         ipp_Cyg_LimDeuEn1.Enabled = False
         cmb_Cyg_ClaEn1.Enabled = False
         cmb_Cyg_CodEn2.Enabled = False
         ipp_Cyg_DeuEn2.Enabled = False
         ipp_Cyg_LimDeuEn2.Enabled = False
         cmb_Cyg_ClaEn2.Enabled = False
         cmb_Cyg_CodEn3.Enabled = False
         ipp_Cyg_DeuEn3.Enabled = False
         ipp_Cyg_LimDeuEn3.Enabled = False
         cmb_Cyg_ClaEn3.Enabled = False
         cmb_Cyg_CodEn4.Enabled = False
         ipp_Cyg_DeuEn4.Enabled = False
         ipp_Cyg_LimDeuEn4.Enabled = False
         cmb_Cyg_ClaEn4.Enabled = False
         cmb_Cyg_CodEn5.Enabled = False
         ipp_Cyg_DeuEn5.Enabled = False
         ipp_Cyg_LimDeuEn5.Enabled = False
         cmb_Cyg_ClaEn5.Enabled = False
         cmb_Cyg_CodEn6.Enabled = False
         ipp_Cyg_DeuEn6.Enabled = False
         ipp_Cyg_LimDeuEn6.Enabled = False
         cmb_Cyg_ClaEn6.Enabled = False
      Else
         ipp_Cyg_CRiEnt.Enabled = True
         ipp_Cyg_CRiCl0.Enabled = True
         ipp_Cyg_CRiCl1.Enabled = True
         ipp_Cyg_CRiCl2.Enabled = True
         ipp_Cyg_CRiCl3.Enabled = True
         ipp_Cyg_CRiCl4.Enabled = True
         ipp_Cyg_TotDMN.Enabled = True
         ipp_Cyg_TotDME.Enabled = True
      End If
   End If
End Sub

Private Sub cmb_Cyg_CRiFlg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_CRiFlg_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipVDm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Verificación Domiciliaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVDm)
      Exit Sub
   End If
   If cmb_Tit_CRiFlg.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Titular presenta Información en Central de Riesgos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tit_CRiFlg)
      Exit Sub
   End If
   If CDate(ipp_Tit_CRiFec.Text) > date Then
      MsgBox "Fecha de Reporte no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Tit_CRiFec)
   End If
   If CDate(ipp_Tit_CRiFec.Text) < date - CDate(60) Then
      MsgBox "Fecha de Reporte no puede ser menor a 60 días.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Tit_CRiFec)
   End If
   
   If cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex) = 1 Then
      If ipp_Tit_CRiEnt.Value = 0 Then
         MsgBox "Debe reportar el Nro. de Entidades con Deuda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Tit_CRiEnt)
         Exit Sub
      End If
      If CDbl(ipp_Tit_CRiCl0.Text) + CDbl(ipp_Tit_CRiCl1.Text) + CDbl(ipp_Tit_CRiCl2.Text) + CDbl(ipp_Tit_CRiCl3.Text) + CDbl(ipp_Tit_CRiCl4.Text) <> 100 Then
         MsgBox "El Total de Clasificación de la deuda del Titular no suma 100%", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_Tit_CRiCl0)
         Exit Sub
      End If
      'If CDbl(ipp_Tit_TotDMN.Text) = 0 And CDbl(ipp_Tit_TotDME.Text) = 0 Then
      '   MsgBox "No ha ingresado el Total de Endeudamiento.", vbExclamation, modgen_g_str_NomPlt
      '   Call gs_SetFocus(ipp_Tit_TotDMN)
      '   Exit Sub
      'End If
      
      If cmb_Tit_CodEn1.Enabled Then
         If cmb_Tit_CodEn1.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn1)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn1.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn1)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn1.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn1)
            Exit Sub
         End If
         If cmb_Tit_ClaEn1.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn1)
            Exit Sub
         End If
      End If
      
      If cmb_Tit_CodEn2.Enabled Then
         If cmb_Tit_CodEn2.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn2)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn2.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn2)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn2.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn2)
            Exit Sub
         End If
         If cmb_Tit_ClaEn2.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn2)
            Exit Sub
         End If
      End If
      
      If cmb_Tit_CodEn3.Enabled Then
         If cmb_Tit_CodEn3.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn3)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn3.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn3)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn3.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn3)
            Exit Sub
         End If
         If cmb_Tit_ClaEn3.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn3)
            Exit Sub
         End If
      End If
      
      If cmb_Tit_CodEn4.Enabled Then
         If cmb_Tit_CodEn4.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn4)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn4.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn4)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn4.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn4)
            Exit Sub
         End If
         If cmb_Tit_ClaEn4.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn4)
            Exit Sub
         End If
      End If
      
      If cmb_Tit_CodEn5.Enabled Then
         If cmb_Tit_CodEn5.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn5)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn5.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn5)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn5.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn5)
            Exit Sub
         End If
         If cmb_Tit_ClaEn5.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn5)
            Exit Sub
         End If
      End If
      
      If cmb_Tit_CodEn6.Enabled Then
         If cmb_Tit_CodEn6.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_CodEn6)
            Exit Sub
         End If
         'If CDbl(ipp_Tit_DeuEn6.Text) = 0 Then
         '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
         '   Call gs_SetFocus(ipp_Tit_DeuEn6)
         '   Exit Sub
         'End If
         If CDbl(ipp_Tit_LimDeuEn6.Text) = 0 Then
            MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_Tit_DeuEn6)
            Exit Sub
         End If
         If cmb_Tit_ClaEn6.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Tit_ClaEn6)
            Exit Sub
         End If
      End If
   End If

   'valida que el total se igual a la suma de las partes
   
   


   If tab_TipCli.TabVisible(1) Then
      If cmb_Cyg_CRiFlg.ListIndex = -1 Then
         MsgBox "Debe seleccionar si el Cónyuge presenta Información en Central de Riesgos.", vbExclamation, modgen_g_str_NomPlt
         tab_TipCli.Tab = 1
         Call gs_SetFocus(cmb_Cyg_CRiFlg)
         Exit Sub
      End If
      If CDate(ipp_Cyg_CRiFec.Text) > date Then
         MsgBox "Fecha de Reporte no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
         tab_TipCli.Tab = 1
         Call gs_SetFocus(ipp_Cyg_CRiFec)
      End If
      If CDate(ipp_Cyg_CRiFec.Text) < date - CDate(60) Then
         MsgBox "Fecha de Reporte no puede ser menor a 60 días.", vbExclamation, modgen_g_str_NomPlt
         tab_TipCli.Tab = 1
         Call gs_SetFocus(ipp_Cyg_CRiFec)
      End If
      If cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex) = 1 Then
         If ipp_Cyg_CRiEnt.Value = 0 Then
            MsgBox "Debe reportar el Nro. de Entidades con Deuda.", vbExclamation, modgen_g_str_NomPlt
            tab_TipCli.Tab = 1
            Call gs_SetFocus(ipp_Cyg_CRiEnt)
            Exit Sub
         End If
         If CDbl(ipp_Cyg_CRiCl0.Text) + CDbl(ipp_Cyg_CRiCl1.Text) + CDbl(ipp_Cyg_CRiCl2.Text) + CDbl(ipp_Cyg_CRiCl3.Text) + CDbl(ipp_Cyg_CRiCl4.Text) <> 100 Then
            MsgBox "El Total de Clasificación de la deuda del Cónyuge no suma 100%", vbExclamation, modgen_g_str_NomPlt
            tab_TipCli.Tab = 1
            Call gs_SetFocus(ipp_Cyg_CRiCl0)
            Exit Sub
         End If
         'If CDbl(ipp_Cyg_TotDMN.Text) = 0 And CDbl(ipp_Cyg_TotDME.Text) = 0 Then
         '   MsgBox "No ha ingresado el Total de Endeudamiento.", vbExclamation, modgen_g_str_NomPlt
         '   tab_TipCli.Tab = 1
         '   Call gs_SetFocus(ipp_Cyg_TotDMN)
         '   Exit Sub
         'End If
         If cmb_Cyg_CodEn1.Enabled Then
            If cmb_Cyg_CodEn1.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn1)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn1.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn1)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn1.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn1)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn1.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn1)
               Exit Sub
            End If
         End If
         
         If cmb_Cyg_CodEn2.Enabled Then
            If cmb_Cyg_CodEn2.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn2)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn2.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn2)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn2.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn2)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn2.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn2)
               Exit Sub
            End If
         End If
         
         If cmb_Cyg_CodEn3.Enabled Then
            If cmb_Cyg_CodEn3.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn3)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn3.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn3)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn3.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn3)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn3.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn3)
               Exit Sub
            End If
         End If
         
         If cmb_Cyg_CodEn4.Enabled Then
            If cmb_Cyg_CodEn4.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn4)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn4.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn4)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn4.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn4)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn4.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn4)
               Exit Sub
            End If
         End If
         
         If cmb_Cyg_CodEn5.Enabled Then
            If cmb_Cyg_CodEn5.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn5)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn5.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn5)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn5.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn5)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn5.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn5)
               Exit Sub
            End If
         End If
         
         If cmb_Cyg_CodEn6.Enabled Then
            If cmb_Cyg_CodEn6.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Entidad.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_CodEn6)
               Exit Sub
            End If
            'If CDbl(ipp_Cyg_DeuEn6.Text) = 0 Then
            '   MsgBox "Debe ingresar el Monto de la Deuda.", vbExclamation, modgen_g_str_NomPlt
            '   tab_TipCli.Tab = 1
            '   Call gs_SetFocus(ipp_Cyg_DeuEn6)
            '   Exit Sub
            'End If
            If CDbl(ipp_Cyg_LimDeuEn6.Text) = 0 Then
               MsgBox "Debe ingresar la Línea Asignada de la Deuda.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Cyg_LimDeuEn6)
               Exit Sub
            End If
            If cmb_Cyg_ClaEn6.ListIndex = -1 Then
               MsgBox "Debe seleccionar la Clasificación.", vbExclamation, modgen_g_str_NomPlt
               tab_TipCli.Tab = 1
               Call gs_SetFocus(cmb_Cyg_ClaEn6)
               Exit Sub
            End If
         End If
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
      g_str_Parame = g_str_Parame & "0, "       'Cuotas Extraordinarias
      g_str_Parame = g_str_Parame & "0, "       'Tipo de Cambio (Para Calificación de Ingresos)
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Calificación de Ingresos
      g_str_Parame = g_str_Parame & "0, "       'Fecha de Calificación de Condiciones de Crédito (Aprobación)
      g_str_Parame = g_str_Parame & "0, "       'Monto Deuda
      
      If cmb_TipVDm.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & CStr(cmb_TipVDm.ItemData(cmb_TipVDm.ListIndex)) & ", "
      End If
      g_str_Parame = g_str_Parame & "'" & txt_ObsVDm.Text & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_Tit_CRiFec.Text), "yyyymmdd") & ", "
      
      If cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiEnt.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl0.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl3.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl4.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_TotDMN.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_TotDME.Value) & ", "
         
         If cmb_Tit_CodEn1.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn1(cmb_Tit_CodEn1.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn1.ItemData(cmb_Tit_ClaEn1.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn2.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn2(cmb_Tit_CodEn2.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn2.ItemData(cmb_Tit_ClaEn2.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn3.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn3(cmb_Tit_CodEn3.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn3.ItemData(cmb_Tit_ClaEn3.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn4.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn4(cmb_Tit_CodEn4.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn4.ItemData(cmb_Tit_ClaEn4.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn5.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn5(cmb_Tit_CodEn5.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn5.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn5.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn5.ItemData(cmb_Tit_ClaEn5.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn6.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn6(cmb_Tit_CodEn6.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn6.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn6.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn6.ItemData(cmb_Tit_ClaEn6.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_Tit_CRiCom.Text & "', "
      
      g_str_Parame = g_str_Parame & "'',"     'Verificación Laboral 1
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Verificación Laboral 1
      g_str_Parame = g_str_Parame & "0, "     'Clasificación Empleador 1
      g_str_Parame = g_str_Parame & "'',"     'Verificación Laboral 2
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Verificación Laboral 2
      g_str_Parame = g_str_Parame & "0, "     'Clasificación Empleador 2
      
      If tab_TipCli.TabVisible(1) Then
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_Cyg_CRiFec.Text), "yyyymmdd") & ", "
         
         If cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiEnt.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl0.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_TotDMN.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_TotDME.Value) & ", "
            
            If cmb_Cyg_CodEn1.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn1(cmb_Cyg_CodEn1.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn1.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn1.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn1.ItemData(cmb_Cyg_ClaEn1.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn2.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn2(cmb_Cyg_CodEn2.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn2.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn2.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn2.ItemData(cmb_Cyg_ClaEn2.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn3.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn3(cmb_Cyg_CodEn3.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn3.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn3.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn3.ItemData(cmb_Cyg_ClaEn3.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn4.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn4(cmb_Cyg_CodEn4.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn4.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn4.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn4.ItemData(cmb_Cyg_ClaEn4.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn5.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn5(cmb_Cyg_CodEn5.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn5.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn5.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn5.ItemData(cmb_Cyg_ClaEn5.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn6.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn6(cmb_Cyg_CodEn6.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn6.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn6.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn6.ItemData(cmb_Cyg_ClaEn6.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_CRiCom.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
      End If
   
      g_str_Parame = g_str_Parame & "'',"     'Cónyuge - Verificación Laboral 1
      g_str_Parame = g_str_Parame & "0, "     'Cónyuge - Tipo de Verificación Laboral 1
      g_str_Parame = g_str_Parame & "0, "     'Cónyuge - Clasificación Empleador 1
      g_str_Parame = g_str_Parame & "'',"     'Cónyuge - Verificación Laboral 2
      g_str_Parame = g_str_Parame & "0, "     'Cónyuge - Tipo de Verificación Laboral 2
      g_str_Parame = g_str_Parame & "0, "     'Cónyuge - Clasificación Empleador 2
      
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Ingreso 1
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Ingreso 2
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Ingreso 3
      g_str_Parame = g_str_Parame & "0, "     'Tipo de Ingreso 4
      
      g_str_Parame = g_str_Parame & "0, "     'Obligaciones Mensuales 1
      g_str_Parame = g_str_Parame & "0, "     'Obligaciones Mensuales 2
      
      g_str_Parame = g_str_Parame & "0, "     'Total Deuda Titular
      g_str_Parame = g_str_Parame & "0, "     'Total Deuda Cónyuge
      
      g_str_Parame = g_str_Parame & "0, "     'Ingreso Neto Titular
      g_str_Parame = g_str_Parame & "0, "     'Ingreso Neto Cónyuge
      
      g_str_Parame = g_str_Parame & "0, "     'Ratio Ingreso Deuda
      g_str_Parame = g_str_Parame & "0, "     'Ratio Inicial Deuda
      
      g_str_Parame = g_str_Parame & "'" & txt_RefPer.Text & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   Else
      g_str_Parame = "USP_TRA_EVACRE_ACT_VERPER ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      If cmb_TipVDm.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & CStr(cmb_TipVDm.ItemData(cmb_TipVDm.ListIndex)) & ", "
      End If
      g_str_Parame = g_str_Parame & "'" & txt_ObsVDm.Text & "', "
   
      g_str_Parame = g_str_Parame & CStr(cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_Tit_CRiFec.Text), "yyyymmdd") & ", "
      
      If cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiEnt.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl0.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl3.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_CRiCl4.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_TotDMN.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_Tit_TotDME.Value) & ", "
         
         If cmb_Tit_CodEn1.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn1(cmb_Tit_CodEn1.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn1.ItemData(cmb_Tit_ClaEn1.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn2.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn2(cmb_Tit_CodEn2.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn2.ItemData(cmb_Tit_ClaEn2.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn3.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn3(cmb_Tit_CodEn3.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn3.ItemData(cmb_Tit_ClaEn3.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn4.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn4(cmb_Tit_CodEn4.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn4.ItemData(cmb_Tit_ClaEn4.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn5.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn5(cmb_Tit_CodEn5.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn5.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn5.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn5.ItemData(cmb_Tit_ClaEn5.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         If cmb_Tit_CodEn6.Enabled Then
            g_str_Parame = g_str_Parame & "'" & l_arr_Tit_CodEn6(cmb_Tit_CodEn6.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_DeuEn6.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Tit_LimDeuEn6.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Tit_ClaEn6.ItemData(cmb_Tit_ClaEn6.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_Tit_CRiCom.Text & "', "
      
      If tab_TipCli.TabVisible(1) Then
         g_str_Parame = g_str_Parame & CStr(cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_Cyg_CRiFec.Text), "yyyymmdd") & ", "
         
         If cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiEnt.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl0.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl1.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl2.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl3.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_CRiCl4.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_TotDMN.Value) & ", "
            g_str_Parame = g_str_Parame & CStr(ipp_Cyg_TotDME.Value) & ", "
            
            If cmb_Cyg_CodEn1.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn1(cmb_Cyg_CodEn1.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn1.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn1.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn1.ItemData(cmb_Cyg_ClaEn1.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn2.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn2(cmb_Cyg_CodEn2.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn2.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn2.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn2.ItemData(cmb_Cyg_ClaEn2.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn3.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn3(cmb_Cyg_CodEn3.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn3.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn3.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn3.ItemData(cmb_Cyg_ClaEn3.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn4.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn4(cmb_Cyg_CodEn4.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn4.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn4.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn4.ItemData(cmb_Cyg_ClaEn4.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn5.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn5(cmb_Cyg_CodEn5.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn5.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn5.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn5.ItemData(cmb_Cyg_ClaEn5.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         
            If cmb_Cyg_CodEn6.Enabled Then
               g_str_Parame = g_str_Parame & "'" & l_arr_Cyg_CodEn6(cmb_Cyg_CodEn6.ListIndex + 1).Genera_Codigo & "', "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_DeuEn6.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(ipp_Cyg_LimDeuEn6.Value) & ", "
               g_str_Parame = g_str_Parame & CStr(cmb_Cyg_ClaEn6.ItemData(cmb_Cyg_ClaEn6.ListIndex)) & ", "
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            End If
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      
         g_str_Parame = g_str_Parame & "'" & txt_Cyg_CRiCom.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_RefPer.Text & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar procedimiento USP_TRA_EVACRE_INSERTA / USP_TRA_EVACRE_ACT_VERPER.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 40, 0, "", 0, 0) Then
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
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If l_int_DatCyg = 1 Then
      tab_TipCli.TabVisible(1) = False
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
      
      If Len(Trim(g_rst_Princi!EVACRE_TIPVDM & "")) > 0 Then
         If g_rst_Princi!EVACRE_TIPVDM > 0 Then
            Call gs_BuscarCombo_Item(cmb_TipVDm, g_rst_Princi!EVACRE_TIPVDM)
         End If
         txt_ObsVDm.Text = Trim(g_rst_Princi!EVACRE_OBSVDM & "")
         
         Call gs_BuscarCombo_Item(cmb_Tit_CRiFlg, g_rst_Princi!EVACRE_TIT_CRIFLG)
         ipp_Tit_CRiFec.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_TIT_CRIFEC))
         
         If cmb_Tit_CRiFlg.ItemData(cmb_Tit_CRiFlg.ListIndex) = 1 Then
            ipp_Tit_CRiEnt.Value = g_rst_Princi!EVACRE_TIT_CRIENT
            ipp_Tit_CRiCl0.Value = g_rst_Princi!EVACRE_TIT_CRICL0
            ipp_Tit_CRiCl1.Value = g_rst_Princi!EVACRE_TIT_CRICL1
            ipp_Tit_CRiCl2.Value = g_rst_Princi!EVACRE_TIT_CRICL2
            ipp_Tit_CRiCl3.Value = g_rst_Princi!EVACRE_TIT_CRICL3
            ipp_Tit_CRiCl4.Value = g_rst_Princi!EVACRE_TIT_CRICL4
            ipp_Tit_TotDMN.Value = g_rst_Princi!EVACRE_TIT_TOTDMN
            ipp_Tit_TotDME.Value = g_rst_Princi!EVACRE_TIT_TOTDME
            
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN1 & "")) > 0 Then
               cmb_Tit_CodEn1.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn1, g_rst_Princi!EVACRE_TIT_CODEN1) - 1
               ipp_Tit_DeuEn1.Value = g_rst_Princi!EVACRE_TIT_DEUEN1
               ipp_Tit_LimDeuEn1.Value = g_rst_Princi!EVACRE_TIT_LIMDE1
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn1, g_rst_Princi!EVACRE_TIT_CLAEN1)
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN2 & "")) > 0 Then
               cmb_Tit_CodEn2.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn2, g_rst_Princi!EVACRE_TIT_CODEN2) - 1
               ipp_Tit_DeuEn2.Value = g_rst_Princi!EVACRE_TIT_DEUEN2
               ipp_Tit_LimDeuEn2.Value = g_rst_Princi!EVACRE_TIT_LIMDE2
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn2, g_rst_Princi!EVACRE_TIT_CLAEN2)
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN3 & "")) > 0 Then
               cmb_Tit_CodEn3.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn3, g_rst_Princi!EVACRE_TIT_CODEN3) - 1
               ipp_Tit_DeuEn3.Value = g_rst_Princi!EVACRE_TIT_DEUEN3
               ipp_Tit_LimDeuEn3.Value = g_rst_Princi!EVACRE_TIT_LIMDE3
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn3, g_rst_Princi!EVACRE_TIT_CLAEN3)
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN4 & "")) > 0 Then
               cmb_Tit_CodEn4.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn4, g_rst_Princi!EVACRE_TIT_CODEN4) - 1
               ipp_Tit_DeuEn4.Value = g_rst_Princi!EVACRE_TIT_DEUEN4
               ipp_Tit_LimDeuEn4.Value = g_rst_Princi!EVACRE_TIT_LIMDE4
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn4, g_rst_Princi!EVACRE_TIT_CLAEN4)
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN5 & "")) > 0 Then
               cmb_Tit_CodEn5.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn5, g_rst_Princi!EVACRE_TIT_CODEN5) - 1
               ipp_Tit_DeuEn5.Value = g_rst_Princi!EVACRE_TIT_DEUEN5
               ipp_Tit_LimDeuEn5.Value = g_rst_Princi!EVACRE_TIT_LIMDE5
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn5, g_rst_Princi!EVACRE_TIT_CLAEN5)
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN6 & "")) > 0 Then
               cmb_Tit_CodEn6.ListIndex = gf_Busca_Arregl(l_arr_Tit_CodEn6, g_rst_Princi!EVACRE_TIT_CODEN6) - 1
               ipp_Tit_DeuEn6.Value = g_rst_Princi!EVACRE_TIT_DEUEN6
               ipp_Tit_LimDeuEn6.Value = g_rst_Princi!EVACRE_TIT_LIMDE6
               Call gs_BuscarCombo_Item(cmb_Tit_ClaEn6, g_rst_Princi!EVACRE_TIT_CLAEN6)
            End If
            
            txt_Tit_CRiCom.Text = Trim(g_rst_Princi!EVACRE_TIT_CRIOBS & "")
         End If
         
         If l_int_DatCyg = 2 Then
            Call gs_BuscarCombo_Item(cmb_Cyg_CRiFlg, g_rst_Princi!EVACRE_CYG_CRIFLG)
            ipp_Cyg_CRiFec.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC))
            
            If cmb_Cyg_CRiFlg.ItemData(cmb_Cyg_CRiFlg.ListIndex) = 1 Then
               ipp_Cyg_CRiEnt.Value = g_rst_Princi!EVACRE_CYG_CRIENT
               ipp_Cyg_CRiCl0.Value = g_rst_Princi!EVACRE_CYG_CRICL0
               ipp_Cyg_CRiCl1.Value = g_rst_Princi!EVACRE_CYG_CRICL1
               ipp_Cyg_CRiCl2.Value = g_rst_Princi!EVACRE_CYG_CRICL2
               ipp_Cyg_CRiCl3.Value = g_rst_Princi!EVACRE_CYG_CRICL3
               ipp_Cyg_CRiCl4.Value = g_rst_Princi!EVACRE_CYG_CRICL4
               ipp_Cyg_TotDMN.Value = g_rst_Princi!EVACRE_CYG_TOTDMN
               ipp_Cyg_TotDME.Value = g_rst_Princi!EVACRE_CYG_TOTDME
               
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN1 & "")) > 0 Then
                  cmb_Cyg_CodEn1.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn1, g_rst_Princi!EVACRE_CYG_CODEN1) - 1
                  ipp_Cyg_DeuEn1.Value = g_rst_Princi!EVACRE_CYG_DEUEN1
                  ipp_Cyg_LimDeuEn1.Value = g_rst_Princi!EVACRE_CYG_LIMDE1
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn1, g_rst_Princi!EVACRE_CYG_CLAEN1)
               End If
            
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN2 & "")) > 0 Then
                  cmb_Cyg_CodEn2.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn2, g_rst_Princi!EVACRE_CYG_CODEN2) - 1
                  ipp_Cyg_DeuEn2.Value = g_rst_Princi!EVACRE_CYG_DEUEN2
                  ipp_Cyg_LimDeuEn2.Value = g_rst_Princi!EVACRE_CYG_LIMDE2
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn2, g_rst_Princi!EVACRE_CYG_CLAEN2)
               End If
            
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN3 & "")) > 0 Then
                  cmb_Cyg_CodEn3.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn3, g_rst_Princi!EVACRE_CYG_CODEN3) - 1
                  ipp_Cyg_DeuEn3.Value = g_rst_Princi!EVACRE_CYG_DEUEN3
                  ipp_Cyg_LimDeuEn3.Value = g_rst_Princi!EVACRE_CYG_LIMDE3
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn3, g_rst_Princi!EVACRE_CYG_CLAEN3)
               End If
            
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN4 & "")) > 0 Then
                  cmb_Cyg_CodEn4.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn4, g_rst_Princi!EVACRE_CYG_CODEN4) - 1
                  ipp_Cyg_DeuEn4.Value = g_rst_Princi!EVACRE_CYG_DEUEN4
                  ipp_Cyg_LimDeuEn4.Value = g_rst_Princi!EVACRE_CYG_LIMDE4
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn4, g_rst_Princi!EVACRE_CYG_CLAEN4)
               End If
            
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN5 & "")) > 0 Then
                  cmb_Cyg_CodEn5.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn5, g_rst_Princi!EVACRE_CYG_CODEN5) - 1
                  ipp_Cyg_DeuEn5.Value = g_rst_Princi!EVACRE_CYG_DEUEN5
                  ipp_Cyg_LimDeuEn5.Value = g_rst_Princi!EVACRE_CYG_LIMDE5
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn5, g_rst_Princi!EVACRE_CYG_CLAEN5)
               End If
            
               If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN6 & "")) > 0 Then
                  cmb_Cyg_CodEn6.ListIndex = gf_Busca_Arregl(l_arr_Cyg_CodEn6, g_rst_Princi!EVACRE_CYG_CODEN6) - 1
                  ipp_Cyg_DeuEn6.Value = g_rst_Princi!EVACRE_CYG_DEUEN6
                  ipp_Cyg_LimDeuEn6.Value = g_rst_Princi!EVACRE_CYG_LIMDE6
                  Call gs_BuscarCombo_Item(cmb_Cyg_ClaEn6, g_rst_Princi!EVACRE_CYG_CLAEN6)
               End If
               
               txt_Cyg_CRiCom.Text = Trim(g_rst_Princi!EVACRE_CYG_CRIOBS & "")
            End If
         End If
         
         txt_RefPer.Text = Trim(g_rst_Princi!EVACRE_REFPER & "")
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVDm, 1, "067")

   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_CRiFlg, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_CRiFlg, 1, "214")

   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn1, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn2, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn3, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn4, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn5, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tit_ClaEn6, 1, "058")

   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn1, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn2, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn3, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn4, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn5, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Cyg_ClaEn6, 1, "058")
   
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn1, l_arr_Tit_CodEn1) 'CALL moddat_gs_Carga_LisIte(cmb_Tit_CodEn1, l_arr_Tit_CodEn1, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn2, l_arr_Tit_CodEn2) 'Call moddat_gs_Carga_LisIte(cmb_Tit_CodEn2, l_arr_Tit_CodEn2, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn3, l_arr_Tit_CodEn3) 'Call moddat_gs_Carga_LisIte(cmb_Tit_CodEn3, l_arr_Tit_CodEn3, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn4, l_arr_Tit_CodEn4) 'Call moddat_gs_Carga_LisIte(cmb_Tit_CodEn4, l_arr_Tit_CodEn4, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn5, l_arr_Tit_CodEn5) 'Call moddat_gs_Carga_LisIte(cmb_Tit_CodEn5, l_arr_Tit_CodEn5, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Tit_CodEn6, l_arr_Tit_CodEn6) 'Call moddat_gs_Carga_LisIte(cmb_Tit_CodEn6, l_arr_Tit_CodEn6, 1, "505")
   
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn1, l_arr_Cyg_CodEn1) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn1, l_arr_Cyg_CodEn1, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn2, l_arr_Cyg_CodEn2) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn2, l_arr_Cyg_CodEn2, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn3, l_arr_Cyg_CodEn3) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn3, l_arr_Cyg_CodEn3, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn4, l_arr_Cyg_CodEn4) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn4, l_arr_Cyg_CodEn4, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn5, l_arr_Cyg_CodEn5) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn5, l_arr_Cyg_CodEn5, 1, "505")
   Call moddat_gs_Carga_EntFin(cmb_Cyg_CodEn6, l_arr_Cyg_CodEn6) 'Call moddat_gs_Carga_LisIte(cmb_Cyg_CodEn6, l_arr_Cyg_CodEn6, 1, "505")
End Sub
Private Sub moddat_gs_Carga_EntFin(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
      
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM CTB_EMPSUP WHERE EMPSUP_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY EMPSUP_NOMCOR ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
      
   Do While Not g_rst_Genera.EOF

      p_Combo.AddItem Trim$(g_rst_Genera!EMPSUP_NOMCOR)
         
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(Format(g_rst_Genera!EMPSUP_CODIGO, "000000"))
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!EMPSUP_NOMCOR)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = 0
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = 0
         
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub
Private Sub fs_Limpia()
   cmb_TipVDm.ListIndex = -1
   txt_ObsVDm.Text = ""
   
   tab_TipCli.Tab = 0
   
   cmb_Tit_CRiFlg.ListIndex = -1
   ipp_Tit_CRiFec.Text = Format(date, "dd/mm/yyyy")
   
   ipp_Tit_CRiEnt.Value = 0
   ipp_Tit_CRiCl0.Value = 0
   ipp_Tit_CRiCl1.Value = 0
   ipp_Tit_CRiCl2.Value = 0
   ipp_Tit_CRiCl3.Value = 0
   ipp_Tit_CRiCl4.Value = 0
   ipp_Tit_TotDMN.Value = 0
   ipp_Tit_TotDME.Value = 0
   
   cmb_Tit_CodEn1.ListIndex = -1
   ipp_Tit_DeuEn1.Value = 0
   ipp_Tit_LimDeuEn1.Value = 0
   cmb_Tit_ClaEn1.ListIndex = -1
   
   cmb_Tit_CodEn2.ListIndex = -1
   ipp_Tit_DeuEn2.Value = 0
   ipp_Tit_LimDeuEn2.Value = 0
   cmb_Tit_ClaEn2.ListIndex = -1
   
   cmb_Tit_CodEn3.ListIndex = -1
   ipp_Tit_DeuEn3.Value = 0
   ipp_Tit_LimDeuEn3.Value = 0
   cmb_Tit_ClaEn3.ListIndex = -1
   
   cmb_Tit_CodEn4.ListIndex = -1
   ipp_Tit_DeuEn4.Value = 0
   ipp_Tit_LimDeuEn4.Value = 0
   cmb_Tit_ClaEn4.ListIndex = -1
   
   cmb_Tit_CodEn5.ListIndex = -1
   ipp_Tit_DeuEn5.Value = 0
   ipp_Tit_LimDeuEn5.Value = 0
   cmb_Tit_ClaEn5.ListIndex = -1
   
   cmb_Tit_CodEn6.ListIndex = -1
   ipp_Tit_DeuEn6.Value = 0
   ipp_Tit_LimDeuEn6.Value = 0
   cmb_Tit_ClaEn6.ListIndex = -1

   ipp_Tit_CRiEnt.Enabled = False
   ipp_Tit_CRiCl0.Enabled = False
   ipp_Tit_CRiCl1.Enabled = False
   ipp_Tit_CRiCl2.Enabled = False
   ipp_Tit_CRiCl3.Enabled = False
   ipp_Tit_CRiCl4.Enabled = False
   ipp_Tit_TotDMN.Enabled = False
   ipp_Tit_TotDME.Enabled = False
   cmb_Tit_CodEn1.Enabled = False
   ipp_Tit_DeuEn1.Enabled = False
   ipp_Tit_LimDeuEn1.Enabled = False
   cmb_Tit_ClaEn1.Enabled = False
   cmb_Tit_CodEn2.Enabled = False
   ipp_Tit_DeuEn2.Enabled = False
   ipp_Tit_LimDeuEn2.Enabled = False
   cmb_Tit_ClaEn2.Enabled = False
   cmb_Tit_CodEn3.Enabled = False
   ipp_Tit_DeuEn3.Enabled = False
   ipp_Tit_LimDeuEn3.Enabled = False
   cmb_Tit_ClaEn3.Enabled = False
   cmb_Tit_CodEn4.Enabled = False
   ipp_Tit_DeuEn4.Enabled = False
   ipp_Tit_LimDeuEn4.Enabled = False
   cmb_Tit_ClaEn4.Enabled = False
   cmb_Tit_CodEn5.Enabled = False
   ipp_Tit_DeuEn5.Enabled = False
   ipp_Tit_LimDeuEn5.Enabled = False
   cmb_Tit_ClaEn5.Enabled = False
   cmb_Tit_CodEn6.Enabled = False
   ipp_Tit_DeuEn6.Enabled = False
   ipp_Tit_LimDeuEn6.Enabled = False
   cmb_Tit_ClaEn6.Enabled = False
   
   txt_Tit_CRiCom.Text = ""
   
   cmb_Cyg_CRiFlg.ListIndex = -1
   ipp_Cyg_CRiFec.Text = Format(date, "dd/mm/yyyy")
   
   ipp_Cyg_CRiEnt.Value = 0
   ipp_Cyg_CRiCl0.Value = 0
   ipp_Cyg_CRiCl1.Value = 0
   ipp_Cyg_CRiCl2.Value = 0
   ipp_Cyg_CRiCl3.Value = 0
   ipp_Cyg_CRiCl4.Value = 0
   ipp_Cyg_TotDMN.Value = 0
   ipp_Cyg_TotDME.Value = 0

   cmb_Cyg_CodEn1.ListIndex = -1
   ipp_Cyg_DeuEn1.Value = 0
   ipp_Cyg_LimDeuEn1.Value = 0
   cmb_Cyg_ClaEn1.ListIndex = -1
   
   cmb_Cyg_CodEn2.ListIndex = -1
   ipp_Cyg_DeuEn2.Value = 0
   ipp_Cyg_LimDeuEn2.Value = 0
   cmb_Cyg_ClaEn2.ListIndex = -1
   
   cmb_Cyg_CodEn3.ListIndex = -1
   ipp_Cyg_DeuEn3.Value = 0
   ipp_Cyg_LimDeuEn3.Value = 0
   cmb_Cyg_ClaEn3.ListIndex = -1
   
   cmb_Cyg_CodEn4.ListIndex = -1
   ipp_Cyg_DeuEn4.Value = 0
   ipp_Cyg_LimDeuEn4.Value = 0
   cmb_Cyg_ClaEn4.ListIndex = -1
   
   cmb_Cyg_CodEn5.ListIndex = -1
   ipp_Cyg_DeuEn5.Value = 0
   ipp_Cyg_LimDeuEn5.Value = 0
   cmb_Cyg_ClaEn5.ListIndex = -1
   
   cmb_Cyg_CodEn6.ListIndex = -1
   ipp_Cyg_DeuEn6.Value = 0
   ipp_Cyg_LimDeuEn6.Value = 0
   cmb_Cyg_ClaEn6.ListIndex = -1

   ipp_Cyg_CRiEnt.Enabled = False
   ipp_Cyg_CRiCl0.Enabled = False
   ipp_Cyg_CRiCl1.Enabled = False
   ipp_Cyg_CRiCl2.Enabled = False
   ipp_Cyg_CRiCl3.Enabled = False
   ipp_Cyg_CRiCl4.Enabled = False
   ipp_Cyg_TotDMN.Enabled = False
   ipp_Cyg_TotDME.Enabled = False
   cmb_Cyg_CodEn1.Enabled = False
   ipp_Cyg_DeuEn1.Enabled = False
   ipp_Cyg_LimDeuEn1.Enabled = False
   cmb_Cyg_ClaEn1.Enabled = False
   cmb_Cyg_CodEn2.Enabled = False
   ipp_Cyg_DeuEn2.Enabled = False
   ipp_Cyg_LimDeuEn2.Enabled = False
   cmb_Cyg_ClaEn2.Enabled = False
   cmb_Cyg_CodEn3.Enabled = False
   ipp_Cyg_DeuEn3.Enabled = False
   ipp_Cyg_LimDeuEn3.Enabled = False
   cmb_Cyg_ClaEn3.Enabled = False
   cmb_Cyg_CodEn4.Enabled = False
   ipp_Cyg_DeuEn4.Enabled = False
   ipp_Cyg_LimDeuEn4.Enabled = False
   cmb_Cyg_ClaEn4.Enabled = False
   cmb_Cyg_CodEn5.Enabled = False
   ipp_Cyg_DeuEn5.Enabled = False
   ipp_Cyg_LimDeuEn5.Enabled = False
   cmb_Cyg_ClaEn5.Enabled = False
   cmb_Cyg_CodEn6.Enabled = False
   ipp_Cyg_DeuEn6.Enabled = False
   ipp_Cyg_LimDeuEn6.Enabled = False
   cmb_Cyg_ClaEn6.Enabled = False
   
   txt_Cyg_CRiCom.Text = ""
   txt_RefPer.Text = ""
End Sub

Private Sub ipp_Cyg_CRiCl0_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_CRiCl1)
   End If
End Sub

Private Sub ipp_Cyg_CRiCl1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_CRiCl2)
   End If
End Sub

Private Sub ipp_Cyg_CRiCl2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_CRiCl3)
   End If
End Sub

Private Sub ipp_Cyg_CRiCl3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_CRiCl4)
   End If
End Sub

Private Sub ipp_Cyg_CRiCl4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_TotDMN)
   End If
End Sub

Private Sub ipp_Cyg_CRiEnt_Change()
   Select Case ipp_Cyg_CRiEnt.Value
      Case 0
         cmb_Cyg_CodEn1.ListIndex = -1:   ipp_Cyg_DeuEn1.Value = 0:       ipp_Cyg_LimDeuEn1.Value = 0:         cmb_Cyg_ClaEn1.ListIndex = -1
         cmb_Cyg_CodEn1.Enabled = False:  ipp_Cyg_DeuEn1.Enabled = False: ipp_Cyg_LimDeuEn1.Enabled = False:   cmb_Cyg_ClaEn1.Enabled = False
         
         cmb_Cyg_CodEn2.ListIndex = -1:   ipp_Cyg_DeuEn2.Value = 0:       ipp_Cyg_LimDeuEn2.Value = 0:         cmb_Cyg_ClaEn2.ListIndex = -1
         cmb_Cyg_CodEn2.Enabled = False:  ipp_Cyg_DeuEn2.Enabled = False: ipp_Cyg_LimDeuEn2.Enabled = False:   cmb_Cyg_ClaEn2.Enabled = False
         
         cmb_Cyg_CodEn3.ListIndex = -1:   ipp_Cyg_DeuEn3.Value = 0:       ipp_Cyg_LimDeuEn3.Value = 0:         cmb_Cyg_ClaEn3.ListIndex = -1
         cmb_Cyg_CodEn3.Enabled = False:  ipp_Cyg_DeuEn3.Enabled = False: ipp_Cyg_LimDeuEn3.Enabled = False:   cmb_Cyg_ClaEn3.Enabled = False
         
         cmb_Cyg_CodEn4.ListIndex = -1:   ipp_Cyg_DeuEn4.Value = 0:       ipp_Cyg_LimDeuEn4.Value = 0:         cmb_Cyg_ClaEn4.ListIndex = -1
         cmb_Cyg_CodEn4.Enabled = False:  ipp_Cyg_DeuEn4.Enabled = False: ipp_Cyg_LimDeuEn4.Enabled = False:   cmb_Cyg_ClaEn4.Enabled = False
         
         cmb_Cyg_CodEn5.ListIndex = -1:   ipp_Cyg_DeuEn5.Value = 0:       ipp_Cyg_LimDeuEn5.Value = 0:         cmb_Cyg_ClaEn5.ListIndex = -1
         cmb_Cyg_CodEn5.Enabled = False:  ipp_Cyg_DeuEn5.Enabled = False: ipp_Cyg_LimDeuEn5.Enabled = False:   cmb_Cyg_ClaEn5.Enabled = False
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:       ipp_Cyg_LimDeuEn6.Value = 0:         cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False: ipp_Cyg_LimDeuEn6.Enabled = False:   cmb_Cyg_ClaEn6.Enabled = False
      
      Case 1
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         
         cmb_Cyg_CodEn2.ListIndex = -1:   ipp_Cyg_DeuEn2.Value = 0:        ipp_Cyg_LimDeuEn2.Value = 0:        cmb_Cyg_ClaEn2.ListIndex = -1
         cmb_Cyg_CodEn2.Enabled = False:  ipp_Cyg_DeuEn2.Enabled = False:  ipp_Cyg_LimDeuEn2.Enabled = False:  cmb_Cyg_ClaEn2.Enabled = False
         
         cmb_Cyg_CodEn3.ListIndex = -1:   ipp_Cyg_DeuEn3.Value = 0:        ipp_Cyg_LimDeuEn3.Value = 0:        cmb_Cyg_ClaEn3.ListIndex = -1
         cmb_Cyg_CodEn3.Enabled = False:  ipp_Cyg_DeuEn3.Enabled = False:  ipp_Cyg_LimDeuEn3.Enabled = False:  cmb_Cyg_ClaEn3.Enabled = False
         
         cmb_Cyg_CodEn4.ListIndex = -1:   ipp_Cyg_DeuEn4.Value = 0:        ipp_Cyg_LimDeuEn4.Value = 0:        cmb_Cyg_ClaEn4.ListIndex = -1
         cmb_Cyg_CodEn4.Enabled = False:  ipp_Cyg_DeuEn4.Enabled = False:  ipp_Cyg_LimDeuEn4.Enabled = False:  cmb_Cyg_ClaEn4.Enabled = False
         
         cmb_Cyg_CodEn5.ListIndex = -1:   ipp_Cyg_DeuEn5.Value = 0:        ipp_Cyg_LimDeuEn5.Value = 0:        cmb_Cyg_ClaEn5.ListIndex = -1
         cmb_Cyg_CodEn5.Enabled = False:  ipp_Cyg_DeuEn5.Enabled = False:  ipp_Cyg_LimDeuEn5.Enabled = False:  cmb_Cyg_ClaEn5.Enabled = False
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:        ipp_Cyg_LimDeuEn6.Value = 0:        cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False:  ipp_Cyg_LimDeuEn6.Enabled = False:  cmb_Cyg_ClaEn6.Enabled = False
      
      Case 2
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         cmb_Cyg_CodEn2.Enabled = True:   ipp_Cyg_DeuEn2.Enabled = True:   ipp_Cyg_LimDeuEn2.Enabled = True:   cmb_Cyg_ClaEn2.Enabled = True
         
         cmb_Cyg_CodEn3.ListIndex = -1:   ipp_Cyg_DeuEn3.Value = 0:        ipp_Cyg_LimDeuEn3.Value = 0:        cmb_Cyg_ClaEn3.ListIndex = -1
         cmb_Cyg_CodEn3.Enabled = False:  ipp_Cyg_DeuEn3.Enabled = False:  ipp_Cyg_LimDeuEn3.Enabled = False:  cmb_Cyg_ClaEn3.Enabled = False
         
         cmb_Cyg_CodEn4.ListIndex = -1:   ipp_Cyg_DeuEn4.Value = 0:        ipp_Cyg_LimDeuEn4.Value = 0:        cmb_Cyg_ClaEn4.ListIndex = -1
         cmb_Cyg_CodEn4.Enabled = False:  ipp_Cyg_DeuEn4.Enabled = False:  ipp_Cyg_LimDeuEn4.Enabled = False:  cmb_Cyg_ClaEn4.Enabled = False
         
         cmb_Cyg_CodEn5.ListIndex = -1:   ipp_Cyg_DeuEn5.Value = 0:        ipp_Cyg_LimDeuEn5.Value = 0:        cmb_Cyg_ClaEn5.ListIndex = -1
         cmb_Cyg_CodEn5.Enabled = False:  ipp_Cyg_DeuEn5.Enabled = False:  ipp_Cyg_LimDeuEn5.Enabled = False:  cmb_Cyg_ClaEn5.Enabled = False
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:        ipp_Cyg_LimDeuEn6.Value = 0:        cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False:  ipp_Cyg_LimDeuEn6.Enabled = False:  cmb_Cyg_ClaEn6.Enabled = False
      
      Case 3
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         cmb_Cyg_CodEn2.Enabled = True:   ipp_Cyg_DeuEn2.Enabled = True:   ipp_Cyg_LimDeuEn2.Enabled = True:   cmb_Cyg_ClaEn2.Enabled = True
         cmb_Cyg_CodEn3.Enabled = True:   ipp_Cyg_DeuEn3.Enabled = True:   ipp_Cyg_LimDeuEn3.Enabled = True:   cmb_Cyg_ClaEn3.Enabled = True
         
         cmb_Cyg_CodEn4.ListIndex = -1:   ipp_Cyg_DeuEn4.Value = 0:        ipp_Cyg_LimDeuEn4.Value = 0:        cmb_Cyg_ClaEn4.ListIndex = -1
         cmb_Cyg_CodEn4.Enabled = False:  ipp_Cyg_DeuEn4.Enabled = False:  ipp_Cyg_LimDeuEn4.Enabled = False:  cmb_Cyg_ClaEn4.Enabled = False
         
         cmb_Cyg_CodEn5.ListIndex = -1:   ipp_Cyg_DeuEn5.Value = 0:        ipp_Cyg_LimDeuEn5.Value = 0:        cmb_Cyg_ClaEn5.ListIndex = -1
         cmb_Cyg_CodEn5.Enabled = False:  ipp_Cyg_DeuEn5.Enabled = False:  ipp_Cyg_LimDeuEn5.Enabled = False:  cmb_Cyg_ClaEn5.Enabled = False
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:        ipp_Cyg_LimDeuEn6.Value = 0:        cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False:  ipp_Cyg_LimDeuEn6.Enabled = False:  cmb_Cyg_ClaEn6.Enabled = False
      
      Case 4
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         cmb_Cyg_CodEn2.Enabled = True:   ipp_Cyg_DeuEn2.Enabled = True:   ipp_Cyg_LimDeuEn2.Enabled = True:   cmb_Cyg_ClaEn2.Enabled = True
         cmb_Cyg_CodEn3.Enabled = True:   ipp_Cyg_DeuEn3.Enabled = True:   ipp_Cyg_LimDeuEn3.Enabled = True:   cmb_Cyg_ClaEn3.Enabled = True
         cmb_Cyg_CodEn4.Enabled = True:   ipp_Cyg_DeuEn4.Enabled = True:   ipp_Cyg_LimDeuEn4.Enabled = True:   cmb_Cyg_ClaEn4.Enabled = True
         
         cmb_Cyg_CodEn5.ListIndex = -1:   ipp_Cyg_DeuEn5.Value = 0:        ipp_Cyg_LimDeuEn5.Value = 0:        cmb_Cyg_ClaEn5.ListIndex = -1
         cmb_Cyg_CodEn5.Enabled = False:  ipp_Cyg_DeuEn5.Enabled = False:  ipp_Cyg_LimDeuEn5.Enabled = False:  cmb_Cyg_ClaEn5.Enabled = False
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:        ipp_Cyg_LimDeuEn6.Value = 0:        cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False:  ipp_Cyg_LimDeuEn6.Enabled = False:  cmb_Cyg_ClaEn6.Enabled = False
      
      Case 5
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         cmb_Cyg_CodEn2.Enabled = True:   ipp_Cyg_DeuEn2.Enabled = True:   ipp_Cyg_LimDeuEn2.Enabled = True:   cmb_Cyg_ClaEn2.Enabled = True
         cmb_Cyg_CodEn3.Enabled = True:   ipp_Cyg_DeuEn3.Enabled = True:   ipp_Cyg_LimDeuEn3.Enabled = True:   cmb_Cyg_ClaEn3.Enabled = True
         cmb_Cyg_CodEn4.Enabled = True:   ipp_Cyg_DeuEn4.Enabled = True:   ipp_Cyg_LimDeuEn4.Enabled = True:   cmb_Cyg_ClaEn4.Enabled = True
         cmb_Cyg_CodEn5.Enabled = True:   ipp_Cyg_DeuEn5.Enabled = True:   ipp_Cyg_LimDeuEn5.Enabled = True:   cmb_Cyg_ClaEn5.Enabled = True
         
         cmb_Cyg_CodEn6.ListIndex = -1:   ipp_Cyg_DeuEn6.Value = 0:        ipp_Cyg_LimDeuEn6.Value = 0:        cmb_Cyg_ClaEn6.ListIndex = -1
         cmb_Cyg_CodEn6.Enabled = False:  ipp_Cyg_DeuEn6.Enabled = False:  ipp_Cyg_LimDeuEn6.Enabled = True:   cmb_Cyg_ClaEn6.Enabled = False
      
      Case Is >= 6
         cmb_Cyg_CodEn1.Enabled = True:   ipp_Cyg_DeuEn1.Enabled = True:   ipp_Cyg_LimDeuEn1.Enabled = True:   cmb_Cyg_ClaEn1.Enabled = True
         cmb_Cyg_CodEn2.Enabled = True:   ipp_Cyg_DeuEn2.Enabled = True:   ipp_Cyg_LimDeuEn2.Enabled = True:   cmb_Cyg_ClaEn2.Enabled = True
         cmb_Cyg_CodEn3.Enabled = True:   ipp_Cyg_DeuEn3.Enabled = True:   ipp_Cyg_LimDeuEn3.Enabled = True:   cmb_Cyg_ClaEn3.Enabled = True
         cmb_Cyg_CodEn4.Enabled = True:   ipp_Cyg_DeuEn4.Enabled = True:   ipp_Cyg_LimDeuEn4.Enabled = True:   cmb_Cyg_ClaEn4.Enabled = True
         cmb_Cyg_CodEn5.Enabled = True:   ipp_Cyg_DeuEn5.Enabled = True:   ipp_Cyg_LimDeuEn5.Enabled = True:   cmb_Cyg_ClaEn5.Enabled = True
         cmb_Cyg_CodEn6.Enabled = True:   ipp_Cyg_DeuEn6.Enabled = True:   ipp_Cyg_LimDeuEn6.Enabled = True:   cmb_Cyg_ClaEn6.Enabled = True
      
   End Select
End Sub

Private Sub ipp_Cyg_CRiEnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_CRiCl0)
   End If
End Sub

Private Sub ipp_Cyg_CRiFec_KeyPress(KeyAscii As Integer)
   If ipp_Cyg_CRiEnt.Enabled Then
      Call gs_SetFocus(ipp_Cyg_CRiEnt)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub ipp_Cyg_TotDMN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_TotDME)
   End If
End Sub

Private Sub ipp_Cyg_TotDME_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_CodEn1)
   End If
End Sub

Private Sub ipp_Tit_CRiCl0_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CRiCl1)
   End If
End Sub

Private Sub ipp_Tit_CRiCl1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CRiCl2)
   End If
End Sub

Private Sub ipp_Tit_CRiCl2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CRiCl3)
   End If
End Sub

Private Sub ipp_Tit_CRiCl3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CRiCl4)
   End If
End Sub

Private Sub ipp_Tit_CRiCl4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_TotDMN)
   End If
End Sub

Private Sub ipp_Tit_CRiEnt_Change()
   Select Case ipp_Tit_CRiEnt.Value
      Case 0
         cmb_Tit_CodEn1.ListIndex = -1:   ipp_Tit_DeuEn1.Value = 0:        ipp_Tit_LimDeuEn1.Value = 0:        cmb_Tit_ClaEn1.ListIndex = -1
         cmb_Tit_CodEn1.Enabled = False:  ipp_Tit_DeuEn1.Enabled = False:  ipp_Tit_LimDeuEn1.Enabled = False:  cmb_Tit_ClaEn1.Enabled = False
         
         cmb_Tit_CodEn2.ListIndex = -1:   ipp_Tit_DeuEn2.Value = 0:        ipp_Tit_LimDeuEn2.Value = 0:        cmb_Tit_ClaEn2.ListIndex = -1
         cmb_Tit_CodEn2.Enabled = False:  ipp_Tit_DeuEn2.Enabled = False:  ipp_Tit_LimDeuEn2.Enabled = False:  cmb_Tit_ClaEn2.Enabled = False
         
         cmb_Tit_CodEn3.ListIndex = -1:   ipp_Tit_DeuEn3.Value = 0:        ipp_Tit_LimDeuEn3.Value = 0:        cmb_Tit_ClaEn3.ListIndex = -1
         cmb_Tit_CodEn3.Enabled = False:  ipp_Tit_DeuEn3.Enabled = False:  ipp_Tit_LimDeuEn3.Enabled = False:  cmb_Tit_ClaEn3.Enabled = False
         
         cmb_Tit_CodEn4.ListIndex = -1:   ipp_Tit_DeuEn4.Value = 0:        ipp_Tit_LimDeuEn4.Value = 0:        cmb_Tit_ClaEn4.ListIndex = -1
         cmb_Tit_CodEn4.Enabled = False:  ipp_Tit_DeuEn4.Enabled = False:  ipp_Tit_LimDeuEn4.Enabled = False:  cmb_Tit_ClaEn4.Enabled = False
         
         cmb_Tit_CodEn5.ListIndex = -1:   ipp_Tit_DeuEn5.Value = 0:        ipp_Tit_LimDeuEn5.Value = 0:        cmb_Tit_ClaEn5.ListIndex = -1
         cmb_Tit_CodEn5.Enabled = False:  ipp_Tit_DeuEn5.Enabled = False:  ipp_Tit_LimDeuEn5.Enabled = False:  cmb_Tit_ClaEn5.Enabled = False
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case 1
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         
         cmb_Tit_CodEn2.ListIndex = -1:   ipp_Tit_DeuEn2.Value = 0:        ipp_Tit_LimDeuEn2.Value = 0:        cmb_Tit_ClaEn2.ListIndex = -1
         cmb_Tit_CodEn2.Enabled = False:  ipp_Tit_DeuEn2.Enabled = False:  ipp_Tit_LimDeuEn2.Enabled = False:  cmb_Tit_ClaEn2.Enabled = False
         
         cmb_Tit_CodEn3.ListIndex = -1:   ipp_Tit_DeuEn3.Value = 0:        ipp_Tit_LimDeuEn3.Value = 0:        cmb_Tit_ClaEn3.ListIndex = -1
         cmb_Tit_CodEn3.Enabled = False:  ipp_Tit_DeuEn3.Enabled = False:  ipp_Tit_LimDeuEn3.Enabled = False:  cmb_Tit_ClaEn3.Enabled = False
         
         cmb_Tit_CodEn4.ListIndex = -1:   ipp_Tit_DeuEn4.Value = 0:        ipp_Tit_LimDeuEn4.Value = 0:        cmb_Tit_ClaEn4.ListIndex = -1
         cmb_Tit_CodEn4.Enabled = False:  ipp_Tit_DeuEn4.Enabled = False:  ipp_Tit_LimDeuEn4.Enabled = False:  cmb_Tit_ClaEn4.Enabled = False
         
         cmb_Tit_CodEn5.ListIndex = -1:   ipp_Tit_DeuEn5.Value = 0:        ipp_Tit_LimDeuEn5.Value = 0:        cmb_Tit_ClaEn5.ListIndex = -1
         cmb_Tit_CodEn5.Enabled = False:  ipp_Tit_DeuEn5.Enabled = False:  ipp_Tit_LimDeuEn5.Enabled = False:  cmb_Tit_ClaEn5.Enabled = False
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case 2
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         cmb_Tit_CodEn2.Enabled = True:   ipp_Tit_DeuEn2.Enabled = True:   ipp_Tit_LimDeuEn2.Enabled = True:   cmb_Tit_ClaEn2.Enabled = True
         
         cmb_Tit_CodEn3.ListIndex = -1:   ipp_Tit_DeuEn3.Value = 0:        ipp_Tit_LimDeuEn3.Value = 0:        cmb_Tit_ClaEn3.ListIndex = -1
         cmb_Tit_CodEn3.Enabled = False:  ipp_Tit_DeuEn3.Enabled = False:  ipp_Tit_LimDeuEn3.Enabled = False:  cmb_Tit_ClaEn3.Enabled = False
         
         cmb_Tit_CodEn4.ListIndex = -1:   ipp_Tit_DeuEn4.Value = 0:        ipp_Tit_LimDeuEn4.Value = 0:        cmb_Tit_ClaEn4.ListIndex = -1
         cmb_Tit_CodEn4.Enabled = False:  ipp_Tit_DeuEn4.Enabled = False:  ipp_Tit_LimDeuEn4.Enabled = False:  cmb_Tit_ClaEn4.Enabled = False
         
         cmb_Tit_CodEn5.ListIndex = -1:   ipp_Tit_DeuEn5.Value = 0:        ipp_Tit_LimDeuEn5.Value = 0:        cmb_Tit_ClaEn5.ListIndex = -1
         cmb_Tit_CodEn5.Enabled = False:  ipp_Tit_DeuEn5.Enabled = False:  ipp_Tit_LimDeuEn5.Enabled = False:  cmb_Tit_ClaEn5.Enabled = False
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case 3
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         cmb_Tit_CodEn2.Enabled = True:   ipp_Tit_DeuEn2.Enabled = True:   ipp_Tit_LimDeuEn2.Enabled = True:   cmb_Tit_ClaEn2.Enabled = True
         cmb_Tit_CodEn3.Enabled = True:   ipp_Tit_DeuEn3.Enabled = True:   ipp_Tit_LimDeuEn3.Enabled = True:   cmb_Tit_ClaEn3.Enabled = True
         
         cmb_Tit_CodEn4.ListIndex = -1:   ipp_Tit_DeuEn4.Value = 0:        ipp_Tit_LimDeuEn4.Value = 0:        cmb_Tit_ClaEn4.ListIndex = -1
         cmb_Tit_CodEn4.Enabled = False:  ipp_Tit_DeuEn4.Enabled = False:  ipp_Tit_LimDeuEn4.Enabled = False:  cmb_Tit_ClaEn4.Enabled = False
         
         cmb_Tit_CodEn5.ListIndex = -1:   ipp_Tit_DeuEn5.Value = 0:        ipp_Tit_LimDeuEn5.Value = 0:        cmb_Tit_ClaEn5.ListIndex = -1
         cmb_Tit_CodEn5.Enabled = False:  ipp_Tit_DeuEn5.Enabled = False:  ipp_Tit_LimDeuEn5.Enabled = False:  cmb_Tit_ClaEn5.Enabled = False
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case 4
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         cmb_Tit_CodEn2.Enabled = True:   ipp_Tit_DeuEn2.Enabled = True:   ipp_Tit_LimDeuEn2.Enabled = True:   cmb_Tit_ClaEn2.Enabled = True
         cmb_Tit_CodEn3.Enabled = True:   ipp_Tit_DeuEn3.Enabled = True:   ipp_Tit_LimDeuEn3.Enabled = True:   cmb_Tit_ClaEn3.Enabled = True
         cmb_Tit_CodEn4.Enabled = True:   ipp_Tit_DeuEn4.Enabled = True:   ipp_Tit_LimDeuEn4.Enabled = True:   cmb_Tit_ClaEn4.Enabled = True
         
         cmb_Tit_CodEn5.ListIndex = -1:   ipp_Tit_DeuEn5.Value = 0:        ipp_Tit_LimDeuEn5.Value = 0:        cmb_Tit_ClaEn5.ListIndex = -1
         cmb_Tit_CodEn5.Enabled = False:  ipp_Tit_DeuEn5.Enabled = False:  ipp_Tit_LimDeuEn5.Enabled = False:  cmb_Tit_ClaEn5.Enabled = False
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case 5
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         cmb_Tit_CodEn2.Enabled = True:   ipp_Tit_DeuEn2.Enabled = True:   ipp_Tit_LimDeuEn2.Enabled = True:   cmb_Tit_ClaEn2.Enabled = True
         cmb_Tit_CodEn3.Enabled = True:   ipp_Tit_DeuEn3.Enabled = True:   ipp_Tit_LimDeuEn3.Enabled = True:   cmb_Tit_ClaEn3.Enabled = True
         cmb_Tit_CodEn4.Enabled = True:   ipp_Tit_DeuEn4.Enabled = True:   ipp_Tit_LimDeuEn4.Enabled = True:   cmb_Tit_ClaEn4.Enabled = True
         cmb_Tit_CodEn5.Enabled = True:   ipp_Tit_DeuEn5.Enabled = True:   ipp_Tit_LimDeuEn5.Enabled = True:   cmb_Tit_ClaEn5.Enabled = True
         
         cmb_Tit_CodEn6.ListIndex = -1:   ipp_Tit_DeuEn6.Value = 0:        ipp_Tit_LimDeuEn6.Value = 0:        cmb_Tit_ClaEn6.ListIndex = -1
         cmb_Tit_CodEn6.Enabled = False:  ipp_Tit_DeuEn6.Enabled = False:  ipp_Tit_LimDeuEn6.Enabled = False:  cmb_Tit_ClaEn6.Enabled = False
      
      Case Is >= 6
         cmb_Tit_CodEn1.Enabled = True:   ipp_Tit_DeuEn1.Enabled = True:   ipp_Tit_LimDeuEn1.Enabled = True:   cmb_Tit_ClaEn1.Enabled = True
         cmb_Tit_CodEn2.Enabled = True:   ipp_Tit_DeuEn2.Enabled = True:   ipp_Tit_LimDeuEn2.Enabled = True:   cmb_Tit_ClaEn2.Enabled = True
         cmb_Tit_CodEn3.Enabled = True:   ipp_Tit_DeuEn3.Enabled = True:   ipp_Tit_LimDeuEn3.Enabled = True:   cmb_Tit_ClaEn3.Enabled = True
         cmb_Tit_CodEn4.Enabled = True:   ipp_Tit_DeuEn4.Enabled = True:   ipp_Tit_LimDeuEn4.Enabled = True:   cmb_Tit_ClaEn4.Enabled = True
         cmb_Tit_CodEn5.Enabled = True:   ipp_Tit_DeuEn5.Enabled = True:   ipp_Tit_LimDeuEn5.Enabled = True:   cmb_Tit_ClaEn5.Enabled = True
         cmb_Tit_CodEn6.Enabled = True:   ipp_Tit_DeuEn6.Enabled = True:   ipp_Tit_LimDeuEn6.Enabled = True:   cmb_Tit_ClaEn6.Enabled = True
      
   End Select
End Sub

Private Sub ipp_Tit_DeuEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn1)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn1)
   End If
End Sub
Private Sub ipp_Tit_DeuEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn2)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn2)
   End If
End Sub
Private Sub ipp_Tit_DeuEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn3)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn3)
   End If
End Sub

Private Sub ipp_Tit_DeuEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn4)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn4)
   End If
End Sub
Private Sub ipp_Tit_DeuEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn5)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn5)
   End If
End Sub
Private Sub ipp_Tit_DeuEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_LimDeuEn6)
   End If
End Sub
Private Sub ipp_Tit_LimDeuEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_ClaEn6)
   End If
End Sub
Private Sub ipp_Tit_TotDMN_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_TotDME)
   End If
End Sub

Private Sub ipp_Tit_TotDME_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Tit_CodEn1)
   End If
End Sub

Private Sub ipp_Tit_CRiEnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CRiCl0)
   End If
End Sub

Private Sub ipp_Tit_CRiFec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_Tit_CRiEnt.Enabled Then
         Call gs_SetFocus(ipp_Tit_CRiEnt)
      Else
         Call gs_SetFocus(txt_Tit_CRiCom)
      End If
   End If
End Sub


Private Sub txt_Cyg_CRiCom_GotFocus()
   Call gs_SelecTodo(txt_Cyg_CRiCom)
End Sub

Private Sub txt_Cyg_CRiCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RefPer)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_ObsVDm_GotFocus()
   Call gs_SelecTodo(txt_ObsVDm)
End Sub

Private Sub txt_ObsVDm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_TipCli.Tab = 0
      Call gs_SetFocus(cmb_Tit_CRiFlg)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Tit_CRiCom_GotFocus()
   Call gs_SelecTodo(txt_Tit_CRiCom)
End Sub

Private Sub txt_Tit_CRiCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If tab_TipCli.TabVisible(1) Then
         tab_TipCli.Tab = 1
         Call gs_SetFocus(cmb_Cyg_CRiFlg)
      Else
         Call gs_SetFocus(txt_RefPer)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_RefPer_GotFocus()
   Call gs_SelecTodo(txt_RefPer)
End Sub

Private Sub txt_RefPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub ipp_Cyg_DeuEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn1)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn1)
   End If
End Sub
Private Sub ipp_Cyg_DeuEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn2)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn2)
   End If
End Sub
Private Sub ipp_Cyg_DeuEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn3)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn3)
   End If
End Sub
Private Sub ipp_Cyg_DeuEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn4)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn4)
   End If
End Sub

Private Sub ipp_Cyg_DeuEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn5)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn5)
   End If
End Sub
Private Sub ipp_Cyg_DeuEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Cyg_LimDeuEn6)
   End If
End Sub
Private Sub ipp_Cyg_LimDeuEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Cyg_ClaEn6)
   End If
End Sub
Private Sub cmb_Cyg_ClaEn1_Click()
   If cmb_Cyg_CodEn2.Enabled Then
      Call gs_SetFocus(cmb_Cyg_CodEn2)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub cmb_Cyg_ClaEn1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn1_Click
   End If
End Sub

Private Sub cmb_Cyg_ClaEn2_Click()
   If cmb_Cyg_CodEn3.Enabled Then
      Call gs_SetFocus(cmb_Cyg_CodEn3)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub cmb_Cyg_ClaEn2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn2_Click
   End If
End Sub

Private Sub cmb_Cyg_ClaEn3_Click()
   If cmb_Cyg_CodEn4.Enabled Then
      Call gs_SetFocus(cmb_Cyg_CodEn4)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub cmb_Cyg_ClaEn3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn3_Click
   End If
End Sub

Private Sub cmb_Cyg_ClaEn4_Click()
   If cmb_Cyg_CodEn5.Enabled Then
      Call gs_SetFocus(cmb_Cyg_CodEn5)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub cmb_Cyg_ClaEn4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn4_Click
   End If
End Sub

Private Sub cmb_Cyg_ClaEn5_Click()
   If cmb_Cyg_CodEn6.Enabled Then
      Call gs_SetFocus(cmb_Cyg_CodEn6)
   Else
      Call gs_SetFocus(txt_Cyg_CRiCom)
   End If
End Sub

Private Sub cmb_Cyg_ClaEn5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn5_Click
   End If
End Sub

Private Sub cmb_Cyg_ClaEn6_Click()
   Call gs_SetFocus(txt_Cyg_CRiCom)
End Sub

Private Sub cmb_Cyg_ClaEn6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Cyg_ClaEn6_Click
   End If
End Sub
