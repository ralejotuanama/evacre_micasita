VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_EvaCre_70 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "EvaCre_frm_536.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10650
      Left            =   -30
      TabIndex        =   24
      Top             =   -30
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
      _ExtentY        =   18785
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
      Begin Threed.SSPanel SSPanel19 
         Height          =   6915
         Left            =   60
         TabIndex        =   37
         Top             =   3630
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   12197
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
         Begin TabDlg.SSTab tab_Ratios 
            Height          =   6735
            Left            =   60
            TabIndex        =   38
            Top             =   90
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   11880
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Información"
            TabPicture(0)   =   "EvaCre_frm_536.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel14"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Ratios Financieros"
            TabPicture(1)   =   "EvaCre_frm_536.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel18"
            Tab(1).Control(1)=   "SSPanel7"
            Tab(1).Control(2)=   "SSPanel17"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Comentarios Indicadores"
            TabPicture(2)   =   "EvaCre_frm_536.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel16"
            Tab(2).Control(1)=   "SSPanel15"
            Tab(2).Control(2)=   "SSPanel13"
            Tab(2).Control(3)=   "SSPanel10"
            Tab(2).Control(4)=   "SSPanel9"
            Tab(2).ControlCount=   5
            Begin Threed.SSPanel SSPanel8 
               Height          =   3015
               Left            =   90
               TabIndex        =   39
               Top             =   390
               Width           =   11205
               _Version        =   65536
               _ExtentX        =   19764
               _ExtentY        =   5318
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
               Begin EditLib.fpDoubleSingle ipp_Tit_ActCte 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   0
                  Top             =   420
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
               Begin EditLib.fpDoubleSingle ipp_Tit_PasCte 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   5
                  Top             =   1290
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
               Begin EditLib.fpDoubleSingle ipp_Tit_ActTot 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   4
                  Top             =   1290
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
               Begin EditLib.fpDoubleSingle ipp_Tit_PasTot 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   7
                  Top             =   1725
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
               Begin EditLib.fpDoubleSingle ipp_Tit_Patrim 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   8
                  Top             =   2160
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
               Begin EditLib.fpDoubleSingle ipp_Tit_MatPri 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   3
                  Top             =   855
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CtaCob 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   1
                  Top             =   420
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
               Begin EditLib.fpDoubleSingle ipp_Tit_Invent 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   2
                  Top             =   855
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CtePrv 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   6
                  Top             =   1725
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
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   22
                  Left            =   2820
                  TabIndex        =   58
                  Top             =   2160
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Patrimonio:"
                  Height          =   255
                  Index           =   19
                  Left            =   60
                  TabIndex        =   57
                  Top             =   2190
                  Width           =   2025
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   13
                  Left            =   2820
                  TabIndex        =   56
                  Top             =   1290
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Activo Total:"
                  Height          =   195
                  Index           =   11
                  Left            =   60
                  TabIndex        =   55
                  Top             =   1350
                  Width           =   1815
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   10
                  Left            =   8580
                  TabIndex        =   54
                  Top             =   1350
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Pasivo Corriente:"
                  Height          =   195
                  Index           =   0
                  Left            =   6000
                  TabIndex        =   53
                  Top             =   1350
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   9
                  Left            =   2820
                  TabIndex        =   52
                  Top             =   480
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Activo Corriente:"
                  Height          =   195
                  Index           =   5
                  Left            =   60
                  TabIndex        =   51
                  Top             =   480
                  Width           =   2295
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   3
                  Left            =   8580
                  TabIndex        =   50
                  Top             =   1785
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Pasivo Total:"
                  Height          =   195
                  Index           =   4
                  Left            =   6000
                  TabIndex        =   49
                  Top             =   1785
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   27
                  Left            =   8580
                  TabIndex        =   48
                  Top             =   915
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Materia Prima (insumos):"
                  Height          =   195
                  Index           =   28
                  Left            =   6000
                  TabIndex        =   47
                  Top             =   915
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   37
                  Left            =   8580
                  TabIndex        =   46
                  Top             =   480
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Cuentas y Documentos por Cobrar:"
                  Height          =   285
                  Index           =   36
                  Left            =   6000
                  TabIndex        =   45
                  Top             =   435
                  Width           =   2475
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   23
                  Left            =   2820
                  TabIndex        =   44
                  Top             =   915
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Inventario:"
                  Height          =   195
                  Index           =   24
                  Left            =   60
                  TabIndex        =   43
                  Top             =   915
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   25
                  Left            =   2820
                  TabIndex        =   42
                  Top             =   1725
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Pasivo Corriente (Proveedor):"
                  Height          =   195
                  Index           =   26
                  Left            =   60
                  TabIndex        =   41
                  Top             =   1785
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Balance General"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   40
                  Left            =   60
                  TabIndex        =   40
                  Top             =   60
                  Width           =   2385
               End
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   3165
               Left            =   90
               TabIndex        =   59
               Top             =   3465
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
               _ExtentY        =   5583
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CobCre 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   10
                  Top             =   600
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
               Begin EditLib.fpDoubleSingle ipp_Tit_EgrCom 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   11
                  Top             =   1050
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
               Begin EditLib.fpDoubleSingle ipp_Tit_CuoPre 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   15
                  Top             =   1935
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
               Begin EditLib.fpDoubleSingle ipp_Tit_ExcMen 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   16
                  Top             =   1935
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
               Begin EditLib.fpDoubleSingle ipp_Tit_IngNet 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   13
                  Top             =   1500
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
               Begin EditLib.fpDoubleSingle ipp_Tit_Ingres 
                  Height          =   315
                  Left            =   3330
                  TabIndex        =   9
                  Top             =   600
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
               Begin EditLib.fpDoubleSingle ipp_Tit_IngBto 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   12
                  Top             =   1050
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
               Begin EditLib.fpDoubleSingle ipp_Tit_NtoNeg 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   14
                  Top             =   1500
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
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   34
                  Left            =   2820
                  TabIndex        =   76
                  Top             =   1110
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Egresos por Compras:"
                  Height          =   195
                  Index           =   35
                  Left            =   60
                  TabIndex        =   75
                  Top             =   1110
                  Width           =   1965
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Cobros (por ventas a crédito):"
                  Height          =   195
                  Index           =   38
                  Left            =   6000
                  TabIndex        =   74
                  Top             =   660
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   39
                  Left            =   8580
                  TabIndex        =   73
                  Top             =   660
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Simulación de Cuota del Préstamo:"
                  Height          =   195
                  Index           =   46
                  Left            =   60
                  TabIndex        =   72
                  Top             =   1995
                  Width           =   2475
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   47
                  Left            =   2820
                  TabIndex        =   71
                  Top             =   1995
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Excedente Mensual:"
                  Height          =   195
                  Index           =   48
                  Left            =   6000
                  TabIndex        =   70
                  Top             =   1995
                  Width           =   2355
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   49
                  Left            =   8580
                  TabIndex        =   69
                  Top             =   1995
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Ingreso Neto:"
                  Height          =   255
                  Index           =   68
                  Left            =   60
                  TabIndex        =   68
                  Top             =   1530
                  Width           =   2025
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   69
                  Left            =   2820
                  TabIndex        =   67
                  Top             =   1560
                  Width           =   465
               End
               Begin VB.Label Label3 
                  Caption         =   "Ingresos:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   66
                  Top             =   660
                  Width           =   2115
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   8
                  Left            =   2820
                  TabIndex        =   65
                  Top             =   660
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   2
                  Left            =   8580
                  TabIndex        =   64
                  Top             =   1110
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Ingreso Bruto (Negocio) :"
                  Height          =   195
                  Index           =   1
                  Left            =   6000
                  TabIndex        =   63
                  Top             =   1110
                  Width           =   1965
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Ingreso Neto(Negocio) :"
                  Height          =   195
                  Index           =   6
                  Left            =   6000
                  TabIndex        =   62
                  Top             =   1560
                  Width           =   1965
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   7
                  Left            =   8580
                  TabIndex        =   61
                  Top             =   1560
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Flujo de Caja Mensual y Familiar"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   41
                  Left            =   60
                  TabIndex        =   60
                  Top             =   60
                  Width           =   3345
               End
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   1965
               Left            =   -74910
               TabIndex        =   77
               Top             =   390
               Width           =   11205
               _Version        =   65536
               _ExtentX        =   19764
               _ExtentY        =   3466
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
               Begin Threed.SSPanel pnl_Rat_RazCor 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   78
                  Top             =   480
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_CapTrb 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   79
                  Top             =   840
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_SolTot 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   80
                  Top             =   480
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
                  Caption         =   "Capital de Trabajo"
                  Height          =   195
                  Index           =   50
                  Left            =   180
                  TabIndex        =   86
                  Top             =   930
                  Width           =   1575
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Razón Corriente"
                  Height          =   195
                  Index           =   51
                  Left            =   180
                  TabIndex        =   85
                  Top             =   600
                  Width           =   1665
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "(a) Ratio de Liquidez"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   52
                  Left            =   180
                  TabIndex        =   84
                  Top             =   120
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Alignment       =   1  'Right Justify
                  Caption         =   "S/."
                  Height          =   195
                  Index           =   54
                  Left            =   2940
                  TabIndex        =   83
                  Top             =   900
                  Width           =   465
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "(b) Ratios de solvencia (Endeudamiento)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   55
                  Left            =   6000
                  TabIndex        =   82
                  Top             =   120
                  Width           =   3825
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Solvecia total"
                  Height          =   195
                  Index           =   56
                  Left            =   6000
                  TabIndex        =   81
                  Top             =   600
                  Width           =   1665
               End
            End
            Begin Threed.SSPanel SSPanel7 
               Height          =   2355
               Left            =   -74910
               TabIndex        =   87
               Top             =   2400
               Width           =   11205
               _Version        =   65536
               _ExtentX        =   19764
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
               Begin Threed.SSPanel pnl_Rat_RenVta 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   88
                  Top             =   540
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_RenPat 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   89
                  Top             =   870
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_RenAct 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   90
                  Top             =   1200
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_RotCob 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   91
                  Top             =   480
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_PerCob 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   92
                  Top             =   810
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_RotMer 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   93
                  Top             =   1140
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_RotPag 
                  Height          =   315
                  Left            =   9090
                  TabIndex        =   94
                  Top             =   1470
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
                  Caption         =   "Rentabilidad de Ventas"
                  Height          =   195
                  Index           =   20
                  Left            =   180
                  TabIndex        =   103
                  Top             =   600
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Rentabilidad del Patrimonio - ROE"
                  Height          =   195
                  Index           =   14
                  Left            =   180
                  TabIndex        =   102
                  Top             =   930
                  Width           =   2415
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Rentabilidad del Activo - ROA"
                  Height          =   195
                  Index           =   16
                  Left            =   180
                  TabIndex        =   101
                  Top             =   1260
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "(c ) Ratios de rentabilidad"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   58
                  Left            =   180
                  TabIndex        =   100
                  Top             =   120
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "(d) Ciclo del negocio"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   59
                  Left            =   6000
                  TabIndex        =   99
                  Top             =   120
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Rotación de Mercaderias "
                  Height          =   195
                  Index           =   60
                  Left            =   6000
                  TabIndex        =   98
                  Top             =   1200
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Periodo promedio de Cobro"
                  Height          =   195
                  Index           =   62
                  Left            =   6000
                  TabIndex        =   97
                  Top             =   870
                  Width           =   2415
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Rotación de Ctas. X Cobrar"
                  Height          =   195
                  Index           =   64
                  Left            =   6000
                  TabIndex        =   96
                  Top             =   540
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Rotación de Cuentas por Pagar"
                  Height          =   195
                  Index           =   66
                  Left            =   6000
                  TabIndex        =   95
                  Top             =   1530
                  Width           =   2385
               End
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   1845
               Left            =   -74910
               TabIndex        =   104
               Top             =   4800
               Width           =   11205
               _Version        =   65536
               _ExtentX        =   19764
               _ExtentY        =   3254
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
               Begin Threed.SSPanel pnl_Rat_CuoIng 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_CuoExc 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   106
                  Top             =   570
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
               Begin Threed.SSPanel pnl_Rat_ApaFin 
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   107
                  Top             =   900
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "0.00  "
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
                  Caption         =   "Ratio cuota /Ingreso  Neto:"
                  Height          =   195
                  Index           =   18
                  Left            =   180
                  TabIndex        =   110
                  Top             =   300
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Ratio cuota/excedente Mensual:"
                  Height          =   195
                  Index           =   30
                  Left            =   180
                  TabIndex        =   109
                  Top             =   630
                  Width           =   2385
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "(e) Apalancamiento Financiero"
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
                  Index           =   32
                  Left            =   180
                  TabIndex        =   108
                  Top             =   960
                  Width           =   2985
               End
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   1215
               Left            =   -74910
               TabIndex        =   111
               Top             =   390
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   2143
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1065
                  Index           =   0
                  Left            =   2670
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   17
                  Top             =   60
                  Width           =   8295
               End
               Begin VB.Label Label6 
                  Caption         =   "a) Liquidez (Capital de Trabajo)"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   112
                  Top             =   435
                  Width           =   2205
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   1215
               Left            =   -74910
               TabIndex        =   113
               Top             =   1650
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   2143
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1065
                  Index           =   1
                  Left            =   2670
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   18
                  Top             =   60
                  Width           =   8295
               End
               Begin VB.Label Label8 
                  Caption         =   "b) Solvencia (Endeudamiento)"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   114
                  Top             =   450
                  Width           =   2175
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   1215
               Left            =   -74910
               TabIndex        =   115
               Top             =   2910
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   2143
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1065
                  Index           =   2
                  Left            =   2670
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   19
                  Top             =   60
                  Width           =   8295
               End
               Begin VB.Label Label9 
                  Caption         =   "c) Rentabilidad"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   116
                  Top             =   450
                  Width           =   2085
               End
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   1215
               Left            =   -74910
               TabIndex        =   117
               Top             =   4170
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   2143
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1065
                  Index           =   3
                  Left            =   2670
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   20
                  Top             =   60
                  Width           =   8295
               End
               Begin VB.Label Label10 
                  Caption         =   "d) Ciclo del Negocio"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   118
                  Top             =   450
                  Width           =   1725
               End
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   1215
               Left            =   -74910
               TabIndex        =   119
               Top             =   5430
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   2143
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
               Begin VB.TextBox txt_Coment 
                  Height          =   1065
                  Index           =   4
                  Left            =   2670
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   21
                  Top             =   60
                  Width           =   8295
               End
               Begin VB.Label Label12 
                  Caption         =   "e) Endeudamiento de Apalancamiento Financiero"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   120
                  Top             =   360
                  Width           =   2025
               End
            End
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   60
         TabIndex        =   25
         Top             =   1470
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
            TabIndex        =   26
            Top             =   390
            Width           =   10065
            _Version        =   65536
            _ExtentX        =   17754
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
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   60
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
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
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   90
            TabIndex        =   31
            Top             =   450
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Solicitud:"
            Height          =   195
            Left            =   7740
            TabIndex        =   29
            Top             =   120
            Width           =   1140
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   32
         Top             =   780
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
            Picture         =   "EvaCre_frm_536.frx":0060
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10950
            Picture         =   "EvaCre_frm_536.frx":04A2
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   33
         Top             =   60
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
            Width           =   6015
            _Version        =   65536
            _ExtentX        =   10610
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Ratios Financieros MicroEmpresario"
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
            Picture         =   "EvaCre_frm_536.frx":08E4
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   1305
         Left            =   60
         TabIndex        =   36
         Top             =   2280
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   2302
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
            Height          =   1245
            Left            =   60
            TabIndex        =   121
            Top             =   30
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2196
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Titular"
            TabPicture(0)   =   "EvaCre_frm_536.frx":0BEE
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_536.frx":0C0A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Inmueble"
            TabPicture(2)   =   "EvaCre_frm_536.frx":0C26
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Datos del Crédito"
            TabPicture(3)   =   "EvaCre_frm_536.frx":0C42
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   795
               Index           =   0
               Left            =   90
               TabIndex        =   122
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   795
               Index           =   1
               Left            =   -74910
               TabIndex        =   123
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   855
               Left            =   -74940
               TabIndex        =   124
               Top             =   3120
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1508
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin VB.TextBox txt_CarMat 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   126
                  Top             =   360
                  Width           =   3945
               End
               Begin VB.TextBox txt_CarFac 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   125
                  Top             =   360
                  Width           =   4335
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Características"
                  Height          =   195
                  Index           =   43
                  Left            =   60
                  TabIndex        =   129
                  Top             =   30
                  Width           =   1065
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Material"
                  Height          =   195
                  Index           =   44
                  Left            =   60
                  TabIndex        =   128
                  Top             =   420
                  Width           =   555
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Fachada"
                  Height          =   195
                  Index           =   45
                  Left            =   6000
                  TabIndex        =   127
                  Top             =   420
                  Width           =   630
               End
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   495
               Left            =   -74940
               TabIndex        =   130
               Top             =   4020
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   873
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   131
                  Top             =   160
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Luz"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   1
                  Left            =   4050
                  TabIndex        =   132
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Agua"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   2
                  Left            =   6660
                  TabIndex        =   133
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Teléfono"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck sck_SerVic 
                  Height          =   195
                  Index           =   3
                  Left            =   9330
                  TabIndex        =   134
                  Top             =   165
                  Width           =   915
                  _Version        =   65536
                  _ExtentX        =   1614
                  _ExtentY        =   344
                  _StockProps     =   78
                  Caption         =   "Internet"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Servicios"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   135
                  Top             =   160
                  Width           =   645
               End
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   870
               Left            =   -74940
               TabIndex        =   136
               Top             =   4560
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1535
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin VB.TextBox txt_ObsTit 
                  Height          =   715
                  Left            =   1800
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   137
                  Text            =   "EvaCre_frm_536.frx":0C5E
                  Top             =   60
                  Width           =   9435
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Observaciones"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   138
                  Top             =   320
                  Width           =   1065
               End
            End
            Begin Threed.SSPanel SSPanel20 
               Height          =   585
               Left            =   -74940
               TabIndex        =   139
               Top             =   6405
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1032
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin VB.CommandButton cmd_VerArc 
                  Caption         =   "Ver"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   10770
                  TabIndex        =   141
                  ToolTipText     =   "Adjuntar Croquis del Negocio"
                  Top             =   120
                  Width           =   435
               End
               Begin VB.CommandButton cmd_CroTit 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   10350
                  TabIndex        =   140
                  ToolTipText     =   "Adjuntar Croquis del Negocio"
                  Top             =   120
                  Width           =   405
               End
               Begin Threed.SSPanel pnl_ArcCqr 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   142
                  Top             =   150
                  Width           =   8505
                  _Version        =   65536
                  _ExtentX        =   15002
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
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Croquis del Negocio"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   143
                  Top             =   210
                  Width           =   1425
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   555
               Left            =   -74940
               TabIndex        =   144
               Top             =   2520
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   979
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
               Begin VB.TextBox txt_OtrEst 
                  Height          =   315
                  Left            =   6810
                  TabIndex        =   146
                  Top             =   120
                  Width           =   4365
               End
               Begin VB.ComboBox cmb_TipEst 
                  Height          =   315
                  Left            =   1800
                  Style           =   2  'Dropdown List
                  TabIndex        =   145
                  Top             =   120
                  Width           =   3945
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Otros"
                  Height          =   195
                  Index           =   70
                  Left            =   6000
                  TabIndex        =   148
                  Top             =   180
                  Width           =   375
               End
               Begin VB.Label lbl_Etique 
                  AutoSize        =   -1  'True
                  Caption         =   "Establecimiento"
                  Height          =   195
                  Index           =   71
                  Left            =   60
                  TabIndex        =   147
                  Top             =   180
                  Width           =   1110
               End
            End
            Begin Threed.SSPanel SSPanel22 
               Height          =   885
               Left            =   -74940
               TabIndex        =   149
               Top             =   5475
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   1561
               _StockProps     =   15
               BackColor       =   14215660
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
               Begin VB.TextBox txt_ResTit 
                  Height          =   735
                  Left            =   1800
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   150
                  Text            =   "EvaCre_frm_536.frx":0C65
                  Top             =   60
                  Width           =   9435
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Resumen"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   151
                  Top             =   330
                  Width           =   675
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   795
               Index           =   2
               Left            =   -74910
               TabIndex        =   152
               Top             =   360
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   1402
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   795
               Index           =   3
               Left            =   -74910
               TabIndex        =   153
               Top             =   360
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   1402
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
      End
   End
End
Attribute VB_Name = "frm_EvaCre_70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
Dim r_int_Contad     As Integer

'   If ipp_Tit_ActCte.Text = "" Then
'      MsgBox "Debe ingresar Activo Corriente.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Tit_ActCte)
'      Exit Sub
'   End If
'
'   If ipp_Tit_PasCte.Text = "" Then
'      MsgBox "Debe ingresar Pasivo Corriente.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_Tit_PasCte)
'      Exit Sub
'   End If
   Call gs_Calcular_Ratios
   For r_int_Contad = 0 To 4
      If txt_Coment(r_int_Contad).Text = "" Then
         MsgBox "Debe ingresar Comentario de los indicadores.", vbExclamation, modgen_g_str_NomPlt
         tab_Ratios.Tab = 2
         Call gs_SetFocus(txt_Coment(r_int_Contad))
         Exit Sub
      End If
   Next r_int_Contad
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
    Do While moddat_g_int_FlgGOK = False
        Screen.MousePointer = 11
        
        g_str_Parame = "USP_MIC_RATFIN ("
        g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
        g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
        g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_ActCte.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_PasCte.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_ActTot.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_PasTot.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_IngBto.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_NtoNeg.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_Ingres.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_Patrim.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_CtaCob.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_Invent.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_CobCre.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_CtePrv.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_MatPri.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_EgrCom.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_CuoPre.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_ExcMen.Value) & ", "
        g_str_Parame = g_str_Parame & CDbl(ipp_Tit_IngNet.Value) & ", "
        
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_RazCor.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_CapTrb.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_SolTot.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_RenVta.Caption)) & ", "
        g_str_Parame = g_str_Parame & Trim(Replace(pnl_Rat_RenAct.Caption, "%", "")) & ", "
        g_str_Parame = g_str_Parame & Trim(Replace(pnl_Rat_RenPat.Caption, "%", "")) & ", "
        
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_RotCob.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_PerCob.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_RotMer.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_RotPag.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_CuoIng.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_CuoExc.Caption)) & ", "
        g_str_Parame = g_str_Parame & CDbl(CStr(pnl_Rat_ApaFin.Caption)) & ", "
         
        For r_int_Contad = 0 To 4
            If Trim(txt_Coment(r_int_Contad).Text) = "" Then
                g_str_Parame = g_str_Parame & "'', "
            Else
                g_str_Parame = g_str_Parame & "'" & Trim(txt_Coment(r_int_Contad).Text) & "', "
            End If
        Next
      
        g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
        g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
        g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
        g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
        g_str_Parame = g_str_Parame & moddat_g_int_FlgAct_1 & ") "
        
        If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
           moddat_g_int_CntErr = moddat_g_int_CntErr + 1
        Else
           moddat_g_int_FlgGOK = True
        End If
        
        If moddat_g_int_CntErr = 6 Then
           If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
              Screen.MousePointer = 0
              Exit Sub
           Else
              moddat_g_int_CntErr = 0
           End If
        End If
      
        Screen.MousePointer = 0
   Loop
   
   Call frm_EvaCre_68.fs_DatRat(moddat_g_str_NumSol, frm_EvaCre_68.grd_Listad(4))
   Call frm_EvaCre_68.fs_DatEEFF(moddat_g_str_NumSol)
   
   MsgBox "Se grabó satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
   
   Unload Me
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub
Private Sub fs_Limpia()
Dim r_int_Contad  As Integer

   ipp_Tit_ActCte.Value = 0
   ipp_Tit_PasCte.Value = 0
   ipp_Tit_ActTot.Value = 0
   ipp_Tit_PasTot.Value = 0
   ipp_Tit_IngBto.Value = 0
   ipp_Tit_NtoNeg.Value = 0
   ipp_Tit_Ingres.Value = 0
   ipp_Tit_Patrim.Value = 0
   ipp_Tit_CtaCob.Value = 0
   ipp_Tit_Invent.Value = 0
   ipp_Tit_CobCre.Value = 0
   ipp_Tit_CtePrv.Value = 0
'   ipp_Tit_RotCob.Value = 0
   ipp_Tit_MatPri.Value = 0
   ipp_Tit_EgrCom.Value = 0
   ipp_Tit_IngNet.Value = 0
   pnl_Rat_RazCor.Caption = "0.00  "
   pnl_Rat_CapTrb.Caption = "0.00  "
   pnl_Rat_SolTot.Caption = "0.00  "
   pnl_Rat_RenVta.Caption = "0.00  "
   pnl_Rat_RenAct.Caption = "0.00  "
   pnl_Rat_RenPat.Caption = "0.00  "
   pnl_Rat_RotCob.Caption = "0.00  "
   pnl_Rat_PerCob.Caption = "0.00  "
   pnl_Rat_RotMer.Caption = "0.00  "
   pnl_Rat_RotPag.Caption = "0.00  "
   pnl_Rat_CuoExc.Caption = "0.00  "
   pnl_Rat_ApaFin.Caption = "0.00  "
   pnl_Rat_CuoIng.Caption = "0.00  "
   
   For r_int_Contad = 0 To 4
      txt_Coment(r_int_Contad).Text = ""
   Next
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   'Inicializando Grid de Cliente y de Cónyuge
    For r_int_Contad = 0 To 3
       grd_Listad(r_int_Contad).ColWidth(0) = 2900:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
    Next r_int_Contad
End Sub
Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

    Me.Caption = modgen_g_str_NomPlt
    
    pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
    pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
    pnl_FecSol.Caption = moddat_g_str_FecIng
    
    Call fs_Limpia
    Call fs_Inicia
    Call fs_Buscar
    Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
    Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
    Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Información del Inmueble
    Call modmip_gs_DatCre(grd_Listad(3), r_arr_Mtz)
   
    'Call fs_Activa(False)
    Call gs_CentraForm(Me)
    
    Screen.MousePointer = 0
End Sub
Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT RATFIN_ACTCTE, RATFIN_CTACOB, RATFIN_INVENT, RATFIN_MATPRI, RATFIN_ACTTOT, RATFIN_PASTOT, RATFIN_PASCTE, RATFIN_CTEPRV, RATFIN_PATRIM, "
   g_str_Parame = g_str_Parame & "         RATFIN_INGRES, RATFIN_COBCRE, RATFIN_EGRCOM, RATFIN_INGBTO, RATFIN_INGNET, RATFIN_NETNEG, RATFIN_CUOPRE, RATFIN_EXCMEN, RATFIN_RAZCTE, "
   g_str_Parame = g_str_Parame & "         RATFIN_CAPTRA, RATFIN_SOLTOT, RATFIN_RENVTA, RATFIN_RENACT, RATFIN_RENPAT, RATFIN_ROTCOB, RATFIN_PERCOB, RATFIN_ROTMER, RATFIN_ROTPAG, "
   g_str_Parame = g_str_Parame & "         RATFIN_CUOING, RATFIN_CUOEXC, RATFIN_APAFIN, RATFIN_COMLIQ, RATFIN_COMSOL, RATFIN_COMREN, RATFIN_COMCIC, RATFIN_COMEND "
   g_str_Parame = g_str_Parame & "    FROM MIC_RATFIN WHERE RATFIN_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       moddat_g_int_FlgAct_1 = 1
       Exit Sub
   End If
   
   moddat_g_int_FlgAct_1 = 2
   g_rst_Princi.MoveFirst
   
   ipp_Tit_ActCte.Text = Format(g_rst_Princi!RATFIN_ACTCTE, "###,###,###,##0.00")
   ipp_Tit_CtaCob.Text = Format(g_rst_Princi!RATFIN_CTACOB, "###,###,###,##0.00")
   ipp_Tit_Invent.Text = Format(g_rst_Princi!RATFIN_INVENT, "###,###,###,##0.00")
   ipp_Tit_MatPri.Text = Format(g_rst_Princi!RATFIN_MATPRI, "###,###,###,##0.00")
   ipp_Tit_ActTot.Text = Format(g_rst_Princi!RATFIN_ACTTOT, "###,###,###,##0.00")
   ipp_Tit_PasCte.Text = Format(g_rst_Princi!RATFIN_PASCTE, "###,###,###,##0.00")
   ipp_Tit_CtePrv.Text = Format(g_rst_Princi!RATFIN_CTEPRV, "###,###,###,##0.00")
   ipp_Tit_PasTot.Text = Format(g_rst_Princi!RATFIN_PASTOT, "###,###,###,##0.00")
   ipp_Tit_Patrim.Text = Format(g_rst_Princi!RATFIN_PATRIM, "###,###,###,##0.00")
   
   ipp_Tit_Ingres.Text = Format(g_rst_Princi!RATFIN_INGRES, "###,###,###,##0.00")
   ipp_Tit_CobCre.Text = Format(g_rst_Princi!RATFIN_COBCRE, "###,###,###,##0.00")
   ipp_Tit_EgrCom.Text = Format(g_rst_Princi!RATFIN_EGRCOM, "###,###,###,##0.00")
   ipp_Tit_IngBto.Text = Format(g_rst_Princi!RATFIN_INGBTO, "###,###,###,##0.00")
   ipp_Tit_IngNet.Text = Format(g_rst_Princi!RATFIN_INGNET, "###,###,###,##0.00")
   ipp_Tit_NtoNeg.Text = Format(g_rst_Princi!RATFIN_NETNEG, "###,###,###,##0.00")
   ipp_Tit_CuoPre.Text = Format(g_rst_Princi!RATFIN_CUOPRE, "###,###,###,##0.00")
   ipp_Tit_ExcMen.Text = Format(g_rst_Princi!RATFIN_EXCMEN, "###,###,###,##0.00")
   
   pnl_Rat_RazCor.Caption = Format(g_rst_Princi!RATFIN_RAZCTE, "0.000000000") & "  "
   pnl_Rat_SolTot.Caption = Format(g_rst_Princi!RATFIN_SOLTOT, "0.000000000") & "  "
   pnl_Rat_CapTrb.Caption = Format(g_rst_Princi!RATFIN_CAPTRA, "###,###,###,##0.00") & "  "
   pnl_Rat_RenVta.Caption = Format(g_rst_Princi!RATFIN_RENVTA, "0.000000000") & "  "
   pnl_Rat_RotCob.Caption = Format(g_rst_Princi!RATFIN_ROTCOB, "0.000000") & "  "
   pnl_Rat_RenPat.Caption = Format(g_rst_Princi!RATFIN_RENPAT, "0.00") & " % "
   pnl_Rat_PerCob.Caption = Format(g_rst_Princi!RATFIN_PERCOB, "0.000000") & "  "
   pnl_Rat_RenAct.Caption = Format(g_rst_Princi!RATFIN_RENACT, "0.00") & " % "
   pnl_Rat_RotMer.Caption = Format(g_rst_Princi!RATFIN_ROTMER, "0.000000") & "  "
   pnl_Rat_RotPag.Caption = Format(g_rst_Princi!RATFIN_ROTPAG, "0.000000") & "  "
   pnl_Rat_CuoIng.Caption = Format(g_rst_Princi!RATFIN_CUOING, "0.000000000") & "  "
   pnl_Rat_CuoExc.Caption = Format(g_rst_Princi!RATFIN_CUOEXC, "0.000000000") & "  "
   pnl_Rat_ApaFin.Caption = Format(g_rst_Princi!RATFIN_APAFIN, "0.000000000") & "  "
   
   If Not IsNull(g_rst_Princi!RATFIN_COMLIQ) Then
      txt_Coment(0).Text = g_rst_Princi!RATFIN_COMLIQ
   End If
   If Not IsNull(g_rst_Princi!RATFIN_COMSOL) Then
      txt_Coment(1).Text = g_rst_Princi!RATFIN_COMSOL
   End If
   If Not IsNull(g_rst_Princi!RATFIN_COMREN) Then
      txt_Coment(2).Text = g_rst_Princi!RATFIN_COMREN
   End If
   If Not IsNull(g_rst_Princi!RATFIN_COMCIC) Then
      txt_Coment(3).Text = g_rst_Princi!RATFIN_COMCIC
   End If
   If Not IsNull(g_rst_Princi!RATFIN_COMEND) Then
      txt_Coment(4).Text = g_rst_Princi!RATFIN_COMEND
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub gs_Calcular_Ratios()
    'Ratio de Liquidez
    If ipp_Tit_PasTot.Value > 0 Then
        pnl_Rat_RazCor.Caption = Format(CDbl(ipp_Tit_ActCte.Value) / CDbl(ipp_Tit_PasTot.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_RazCor.Caption = "0.00  "
    End If
    
    pnl_Rat_CapTrb.Caption = Format(CDbl(ipp_Tit_ActCte.Value) - CDbl(ipp_Tit_PasTot.Value), "###,###,###,##0.00") & "  "
    
    'Ratio de Solvencia
    If ipp_Tit_PasTot.Value > 0 Then
        pnl_Rat_SolTot.Caption = Format(CDbl(ipp_Tit_ActTot.Value) / CDbl(ipp_Tit_PasTot.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_SolTot.Caption = "0.00  "
    End If
    
    'Ratios de rentabilidad
    If ipp_Tit_Ingres.Value > 0 Then
        pnl_Rat_RenVta.Caption = Format(CDbl(ipp_Tit_IngBto.Value) / CDbl(ipp_Tit_Ingres.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_RenVta.Caption = "0.00  "
    End If
    
    If ipp_Tit_Patrim.Value > 0 Then
        pnl_Rat_RenPat.Caption = Format((CDbl(ipp_Tit_NtoNeg.Value) / (CDbl(ipp_Tit_Patrim.Value)) * 100), "0") & " % "
    Else
        pnl_Rat_RenPat.Caption = "0.00  "
    End If
    
    If ipp_Tit_ActTot.Value > 0 Then
        pnl_Rat_RenAct.Caption = Format((CDbl(ipp_Tit_NtoNeg.Value) / (CDbl(ipp_Tit_ActTot.Value)) * 100), "0") & " % "
    Else
        pnl_Rat_RenAct.Caption = "0.00  "
    End If
    
    'Ciclo del negocio
    If ipp_Tit_CobCre.Value > 0 Then
        pnl_Rat_RotCob.Caption = Format(CDbl(ipp_Tit_CtaCob.Value) / CDbl(ipp_Tit_CobCre.Value) * 30, "###,###,###,##0.000000") & "  "
    Else
        pnl_Rat_RotCob.Caption = "0.00  "
    End If
    
    If CDbl(pnl_Rat_RotCob.Caption) > 0 Then
        pnl_Rat_PerCob.Caption = Format(CDbl(30) / CDbl(pnl_Rat_RotCob.Caption), "###,###,###,##0.000000") & "  "
    Else
        pnl_Rat_PerCob.Caption = "0.00  "
    End If
    
    If ipp_Tit_Invent.Value > 0 Then
        pnl_Rat_RotMer.Caption = Format(CDbl(ipp_Tit_EgrCom.Value) / CDbl(ipp_Tit_Invent.Value), "###,###,###,##0.000000") & "  "
    Else
        pnl_Rat_RotMer.Caption = "0.00  "
    End If
    
    If ipp_Tit_MatPri.Value > 0 Then
        pnl_Rat_RotPag.Caption = Format(CDbl(ipp_Tit_CtePrv.Value) / CDbl(ipp_Tit_MatPri.Value), "###,###,###,##0.000000") & "  "
    Else
        pnl_Rat_RotPag.Caption = "0.00  "
    End If
    
    If ipp_Tit_IngNet.Value > 0 Then
        pnl_Rat_CuoIng.Caption = Format(CDbl(ipp_Tit_CuoPre.Value) / CDbl(ipp_Tit_IngNet.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_CuoIng.Caption = "0.00  "
    End If
    
    If ipp_Tit_ExcMen.Value > 0 Then
        pnl_Rat_CuoExc.Caption = Format(CDbl(ipp_Tit_CuoPre.Value) / CDbl(ipp_Tit_ExcMen.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_CuoExc.Caption = "0.00  "
    End If
    
    If ipp_Tit_Patrim.Value > 0 Then
        pnl_Rat_ApaFin.Caption = Format(CDbl(ipp_Tit_PasTot.Value) / CDbl(ipp_Tit_Patrim.Value), "###,###,###,##0.000000000") & "  "
    Else
        pnl_Rat_ApaFin.Caption = "0.00  "
    End If
End Sub

Private Sub ipp_Tit_ActCte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CtaCob)
   End If
End Sub

Private Sub ipp_Tit_ActCte_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_ActTot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_PasCte)
    End If
End Sub

Private Sub ipp_Tit_ActTot_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_CobCre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_EgrCom)
    End If
End Sub

Private Sub ipp_Tit_CobCre_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_CtaCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_Invent)
    End If
End Sub


Private Sub ipp_Tit_CtaCob_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_CtePrv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_PasTot)
    End If
End Sub

Private Sub ipp_Tit_CtePrv_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_CuoPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_ExcMen)
    End If
End Sub

Private Sub ipp_Tit_CuoPre_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_EgrCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_IngBto)
    End If
End Sub

Private Sub ipp_Tit_EgrCom_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_ExcMen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Grabar)
    End If
End Sub

Private Sub ipp_Tit_ExcMen_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_IngBto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_IngNet)
    End If
End Sub

Private Sub ipp_Tit_IngBto_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_IngNet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_NtoNeg)
    End If
End Sub

Private Sub ipp_Tit_NtoNeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_CuoPre)
    End If
End Sub

Private Sub ipp_Tit_NtoNeg_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_Ingres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_Tit_CobCre)
    End If
End Sub

Private Sub ipp_Tit_Ingres_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_Invent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_MatPri)
   End If
End Sub

Private Sub ipp_Tit_Invent_LostFocus()
   Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_MatPri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_ActTot)
   End If
End Sub

Private Sub ipp_Tit_MatPri_LostFocus()
   Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_PasCte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_CtePrv)
   End If
End Sub

Private Sub ipp_Tit_PasCte_LostFocus()
    Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_PasTot_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_Patrim)
   End If
End Sub

Private Sub ipp_Tit_PasTot_LostFocus()
   Call gs_Calcular_Ratios
End Sub

Private Sub ipp_Tit_Patrim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tit_Ingres)
   End If
End Sub

Private Sub ipp_Tit_Patrim_LostFocus()
   Call gs_Calcular_Ratios
End Sub

Private Sub txt_Coment_GotFocus(Index As Integer)
   Call gs_SelecTodo(txt_Coment(Index))
End Sub

'Private Sub ipp_Tit_RotCob_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Call gs_SetFocus(ipp_Tit_MatPri)
'    End If
'End Sub
'
'Private Sub ipp_Tit_RotCob_LostFocus()
'    Call gs_Calcular_Ratios
'End Sub

Private Sub txt_Coment_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      Select Case Index
         Case 0:  Call gs_SetFocus(txt_Coment(1))
         Case 1:  Call gs_SetFocus(txt_Coment(2))
         Case 2:  Call gs_SetFocus(txt_Coment(3))
         Case 3:  Call gs_SetFocus(txt_Coment(4))
         Case 4:  Call gs_SetFocus(cmd_Grabar)
      End Select
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,@#$%&;:()/º")
   End If
End Sub
