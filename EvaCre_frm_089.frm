VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_FicEva_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   6435
   ClientLeft      =   4035
   ClientTop       =   3555
   ClientWidth     =   12405
   Icon            =   "EvaCre_frm_089.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6435
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12405
      _Version        =   65536
      _ExtentX        =   21881
      _ExtentY        =   11351
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
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   780
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Left            =   11700
            Picture         =   "EvaCre_frm_089.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_089.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_089.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1230
            Picture         =   "EvaCre_frm_089.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10710
            Top             =   90
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   4125
         Left            =   60
         TabIndex        =   9
         Top             =   2280
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   7276
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
            Height          =   3705
            Left            =   30
            TabIndex        =   6
            Top             =   360
            Width           =   12225
            _ExtentX        =   21564
            _ExtentY        =   6535
            _Version        =   393216
            Rows            =   30
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   3330
            TabIndex        =   10
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   1950
            TabIndex        =   11
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   4560
            TabIndex        =   12
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   8010
            TabIndex        =   13
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   10410
            TabIndex        =   15
            Top             =   60
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consej. Hipotecario"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_IngIns 
            Height          =   285
            Left            =   9210
            TabIndex        =   16
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Resultado"
            ForeColor       =   16777215
            BackColor       =   16384
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Height          =   585
            Left            =   690
            TabIndex        =   18
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Emisión de Fichas de Evaluación Crediticia de Créditos Hipotecarios"
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
            Picture         =   "EvaCre_frm_089.frx":0EA4
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   60
         TabIndex        =   19
         Top             =   1470
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   60
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            Text            =   "01/01/2008"
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1440
            TabIndex        =   1
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            Text            =   "01/01/2008"
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
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_FicEva_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "La Fecha de Inicio no puede ser mayor a la Fecha de Fin.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
Dim r_rst_MaeCli        As ADODB.Recordset
Dim r_rst_MaeSol        As ADODB.Recordset
Dim r_rst_SegCon        As ADODB.Recordset
Dim r_rst_SegExc        As ADODB.Recordset
Dim r_rst_Coment        As ADODB.Recordset
Dim r_str_Tit_CodPri    As String
Dim r_str_Tit_CodSec    As String
Dim r_int_Tit_TDoPri    As Integer
Dim r_str_Tit_NDoPri    As String
Dim r_str_Tit_EmpPri    As String
Dim r_int_Tit_TDoSec    As Integer
Dim r_str_Tit_NDoSec    As String
Dim r_str_Tit_EmpSec    As String
Dim r_str_Cyg_CodPri    As String
Dim r_str_Cyg_CodSec    As String
Dim r_int_Cyg_TDoPri    As Integer
Dim r_str_Cyg_NDoPri    As String
Dim r_str_Cyg_EmpPri    As String
Dim r_int_Cyg_TDoSec    As Integer
Dim r_str_Cyg_NDoSec    As String
Dim r_str_Cyg_EmpSec    As String
Dim r_int_TipEva        As Integer
Dim r_int_FlgCon        As Integer
Dim r_str_AprCon        As String
Dim r_int_FlgExc        As Integer
Dim r_str_AprExc        As String
Dim r_str_Comentario    As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro de imprimir la Ficha de Evaluación Crediticia?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM RPT_FICCRE WHERE FICCRE_NOMTER = '" & modgen_g_str_NombPC & "' AND FICCRE_NUMSOL = '" & gf_Formato_NumSol(moddat_g_str_NumSol) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_FICCRE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'Para obtener Actividades Económicas del Cliente Titular
   r_str_Tit_CodPri = ""
   r_str_Tit_CodSec = ""
   r_int_Tit_TDoPri = 0
   r_str_Tit_NDoPri = ""
   r_str_Tit_EmpPri = ""
   r_int_Tit_TDoSec = 0
   r_str_Tit_NDoSec = ""
   r_str_Tit_EmpSec = ""
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!ActEco_OrdAct = 1 Then
            r_str_Tit_CodPri = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
         Else
            r_str_Tit_CodSec = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
         End If
            
         Select Case g_rst_Princi!ACTECO_CODACT
            Case 11
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  r_int_Tit_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                  r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
               Else
                  r_int_Tit_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                  r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
               End If
               
            Case 21
               If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Tit_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                  Else
                     r_int_Tit_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                  End If
               End If
               
            Case 41
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  r_int_Tit_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                  r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
               Else
                  r_int_Tit_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                  r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
               End If
         End Select
            
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscando Calificación de Empresas (Actividad Principal - Titular)
   r_str_Tit_EmpPri = ""
   If r_int_Tit_TDoPri > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Tit_TDoPri) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Tit_NDoPri & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         r_str_Tit_EmpPri = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   'Buscando Calificación de Empresas (Actividad Secundaria - Titular)
   r_str_Tit_EmpSec = ""
   If r_int_Tit_TDoSec > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Tit_TDoSec) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Tit_NDoSec & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         r_str_Tit_EmpSec = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   'Cónyuge
   r_str_Cyg_CodPri = ""
   r_str_Cyg_CodSec = ""
   r_int_Cyg_TDoPri = 0
   r_str_Cyg_NDoPri = ""
   r_str_Cyg_EmpPri = ""
   r_int_Cyg_TDoSec = 0
   r_str_Cyg_NDoSec = ""
   r_str_Cyg_EmpSec = ""
   
   grd_Listad.Col = 12
   moddat_g_int_CygTDo = Left(grd_Listad.Text, 1)
   moddat_g_str_CygNDo = Mid(grd_Listad.Text, 3)
   
   If moddat_g_int_CygTDo > 0 Then
      'Para obtener Actividades Económicas del Cliente Cónyuge
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' "
      g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!ActEco_OrdAct = 1 Then
               r_str_Cyg_CodPri = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
            Else
               r_str_Cyg_CodSec = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
            End If
               
            Select Case g_rst_Princi!ACTECO_CODACT
               Case 11
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                     r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  Else
                     r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                     r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  End If
                  
               Case 21
                  If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                     If g_rst_Princi!ActEco_OrdAct = 1 Then
                        r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     Else
                        r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     End If
                  End If
                  
               Case 41
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                     r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  Else
                     r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                     r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  End If
            End Select
               
            g_rst_Princi.MoveNext
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      'Buscando Calificación de Empresas (Actividad Principal - Cónyuge)
      r_str_Cyg_EmpPri = ""
      If r_int_Cyg_TDoPri > 0 Then
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Cyg_TDoPri) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Cyg_NDoPri & "' "
       
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_str_Cyg_EmpPri = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End If
   
      'Buscando Calificación de Empresas (Actividad Secundaria - Cónyuge)
      r_str_Cyg_EmpSec = ""
      If r_int_Cyg_TDoSec > 0 Then
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Cyg_TDoSec) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Cyg_NDoSec & "' "
       
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_str_Cyg_EmpSec = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End If
   End If

   'Abriendo Cursor sobre Maestro de Clientes
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MaeCli, 3) Then
       Exit Sub
   End If
   
   r_rst_MaeCli.MoveFirst
   
   'Abriendo Cursor sobre Maestro de Solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MaeSol, 3) Then
       Exit Sub
   End If
   
   r_rst_MaeSol.MoveFirst
   r_int_TipEva = r_rst_MaeSol!SOLMAE_TIPEVA
   
   'Abriendo Cursos sobre Aprobación Condicionada
   r_int_FlgCon = 0
   r_str_AprCon = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGCON_CODINS = 21"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SegCon, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_SegCon.BOF And r_rst_SegCon.EOF) Then
      r_rst_SegCon.MoveFirst
      r_int_FlgCon = 1
      r_str_AprCon = Trim(r_rst_SegCon!SEGCON_OBSCON & "")
   End If
   
   'Abriendo Cursos sobre Aprobación Por Excepción
   r_int_FlgExc = 0
   r_str_AprExc = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGEXC WHERE "
   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGEXC_CODINS = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGEXC_NUMEXC DESC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SegExc, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_SegExc.BOF And r_rst_SegExc.EOF) Then
      r_rst_SegExc.MoveFirst
      r_int_FlgExc = 1
      r_str_AprExc = Trim(r_rst_SegExc!SEGEXC_DESCRI & "")
   End If
   
   'Obtener el comentario de datos de una solicitud
   g_str_Parame = ""
   g_str_Parame = "SELECT SEGDET_OBSERV FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 AND SEGDET_CODOCU = 17 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Coment, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Coment.BOF And r_rst_Coment.EOF) Then
      r_rst_Coment.MoveFirst
      r_str_Comentario = Trim(r_rst_Coment!SEGDET_OBSERV & "")
   End If
   
   'Abriendo Cursos sobre Evaluación Crediticia
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If g_rst_Princi!EVACRE_TIPVDM = 0 And g_rst_Princi!SEGFECCRE < 20100630 Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      r_rst_MaeCli.Close
      Set r_rst_MaeCli = Nothing
      
      r_rst_MaeSol.Close
      Set r_rst_MaeSol = Nothing
   
      r_rst_SegCon.Close
      Set r_rst_SegCon = Nothing
      
      r_rst_SegExc.Close
      Set r_rst_SegExc = Nothing
   
      MsgBox "No se puede imprimir la Ficha de Evaluación Credicitia, porque este cliente no ha sido generado bajo el nuevo formato.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "INSERT INTO RPT_FICCRE ("
   g_str_Parame = g_str_Parame & "FICCRE_NOMTER, "
   g_str_Parame = g_str_Parame & "FICCRE_NUMSOL, "
   g_str_Parame = g_str_Parame & "FICCRE_PRODUC, "
   g_str_Parame = g_str_Parame & "FICCRE_SUBPRD, "
   g_str_Parame = g_str_Parame & "FICCRE_TIPEVA, "
   g_str_Parame = g_str_Parame & "FICCRE_CLINOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CLIDOC, "
   g_str_Parame = g_str_Parame & "FICCRE_CLIEST, "
   g_str_Parame = g_str_Parame & "FICCRE_CYGNOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CYGDOC, "
   g_str_Parame = g_str_Parame & "FICCRE_MONCRE, "
   g_str_Parame = g_str_Parame & "FICCRE_MONSIM, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLCVT, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLINI, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLPRE, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLPLA, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLGRA, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLTSG, "
   g_str_Parame = g_str_Parame & "FICCRE_TASINT, "
   g_str_Parame = g_str_Parame & "FICCRE_VERDOM, "
   g_str_Parame = g_str_Parame & "FICCRE_REFPER, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGLIQ, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_TOTDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INIDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_CUOSOL, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_CUOMPR, "
   g_str_Parame = g_str_Parame & "FICCRE_APRPRE, "
   g_str_Parame = g_str_Parame & "FICCRE_APRPLA, "
   g_str_Parame = g_str_Parame & "FICCRE_APRGRA, "
   g_str_Parame = g_str_Parame & "FICCRE_APRTSG, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEFLG, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEFEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEENT, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL0, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL1, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL2, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL3, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL4, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_TDEUMN, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_TDEUME, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECOM, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_ACTPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_EMPPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_VERPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_ACTSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_EMPSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_VERSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGADI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEFLG, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEFEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEENT, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL0, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL1, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL2, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL3, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL4, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_TDEUMN, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_TDEUME, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_ACTPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_EMPPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_VERPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_ACTSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_EMPSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_VERSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGADI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOENT, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOMON, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOMTO, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOTPO, "
   g_str_Parame = g_str_Parame & "FICCRE_FLGCON, "
   g_str_Parame = g_str_Parame & "FICCRE_APRCON, "
   g_str_Parame = g_str_Parame & "FICCRE_FLGEXC, "
   g_str_Parame = g_str_Parame & "FICCRE_APREXC, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLCUOEXT, "
   g_str_Parame = g_str_Parame & "FICCRE_APRCUOEXT,"
   g_str_Parame = g_str_Parame & "FICCRE_COMENT,"
   g_str_Parame = g_str_Parame & "FICCRE_GASCIE,"
   g_str_Parame = g_str_Parame & "FICCRE_FMVBBP,"
   g_str_Parame = g_str_Parame & "FICCRE_BMSMTO)"
   
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(moddat_g_str_NumSol) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_MaeSol!SOLMAE_CODPRD) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_SubPrd(r_rst_MaeSol!SOLMAE_CODPRD, r_rst_MaeSol!SOLMAE_CODSUB) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_MaeSol!SOLMAE_TIPEVA)) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NomCli & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & "-" & Trim(moddat_g_str_NumDoc) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("205", r_rst_MaeCli!DATGEN_ESTCIV) & "', "
   
   If moddat_g_int_CygTDo > 0 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_CygTDo) & "-" & Trim(moddat_g_str_CygNDo) & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", r_rst_MaeSol!SOLMAE_TIPMON) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", r_rst_MaeSol!SOLMAE_TIPMON) & "', "
   
   If r_rst_MaeSol!SOLMAE_TIPMON = 2 Then
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_COMVTA_DOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_APOPRO_DOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOPRE_DOL) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_COMVTA_SOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_APOPRO_SOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOPRE_SOL) & ", "
   End If
   
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_PLAANO) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_PERGRA) & ", "
   
   If r_rst_MaeSol!SOLMAE_TASESP = 0 Or r_rst_MaeSol!SOLMAE_TASESP = 1 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(r_rst_MaeSol!SOLMAE_ESGDES, r_rst_MaeSol!SOLMAE_TIPSEG) & "', "
   Else
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(r_rst_MaeSol!SOLMAE_ESGDES, r_rst_MaeSol!SOLMAE_TIPSEG) & " - " & Trim(Replace(moddat_gf_Consulta_ParDes("522", CStr(r_rst_MaeSol!SOLMAE_TASESP)), "ADIC", "")) & " Adicional" & "', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_TASINT) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("067", CStr(g_rst_Princi!EVACRE_TIPVDM)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_OBSVDM) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_REFPER) & "', "
   
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGTOT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OBLMEN) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_MTODEU) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGNET) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_RINGDE) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_RINIDE) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CUOSOL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CUOMPR) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_MTOPRE_CAL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(r_rst_MaeSol!SOLMAE_ESGDES, g_rst_Princi!EVACRE_TIPSEG_CAL) & "', "
   
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_TIT_CRIFLG)) & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_TIT_CRIFEC)) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRIENT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL0) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL1) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL2) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL3) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL4) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_TOTDMN) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_TOTDME) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_TIT_CRIOBS) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_CodPri) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_EmpPri) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP1) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_TIT_TIPVE1)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_TIT_LABVE1) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_CodSec) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_EmpSec) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP2) & ", "
   
   If Len(Trim(g_rst_Princi!EVACRE_TIT_LABVE2)) > 0 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_TIT_TIPVE2)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_TIT_LABVE2) & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGAD1) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OMETIT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INTTIT) & ", "
   
   'If (g_rst_Princi!EVACRE_CYG_CRIFLG > 0) And (g_rst_Princi!EVACRE_TDECYG > 0) Then
   If (g_rst_Princi!EVACRE_CYG_CRIFLG > 0) And (g_rst_Princi!EVACRE_INTCYG > 0) Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG)) & "', "
      g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC)) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRIENT) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL0) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL1) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL2) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL3) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL4) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_TOTDMN) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_TOTDME) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_CYG_CRIOBS) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_CodPri) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_EmpPri) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP3) & ", "
      
      If Len(Trim(g_rst_Princi!EVACRE_CYG_LABVE1)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_CYG_TIPVE1)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_CYG_LABVE1) & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_CodSec) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_EmpSec) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP4) & ", "
      
      If Len(Trim(g_rst_Princi!EVACRE_CYG_LABVE2)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_CYG_TIPVE2)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_CYG_LABVE2) & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGAD2) & ", "
      'g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OMECYG) & ", "
      'g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INTCYG) & ", "
      
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGAD2) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OMECYG) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INTCYG) & ", "
      
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   If r_rst_MaeSol!SOLMAE_TIPEVA = 2 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("505", r_rst_MaeSol!SOLMAE_INSFIN) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", r_rst_MaeSol!SOLMAE_MONAHO) & "', "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOAHO) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MESAHO) & ", "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   g_str_Parame = g_str_Parame & CStr(r_int_FlgCon) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_AprCon & "', "
   g_str_Parame = g_str_Parame & CStr(r_int_FlgExc) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_AprExc & "', "
   
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT) & ", "
   If IsNull(r_rst_MaeSol!SOLMAE_CUOEXT_CAL) Then
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT_CAL) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & CStr(r_str_Comentario) & "', "
   g_str_Parame = g_str_Parame & r_rst_MaeSol!SOLMAE_MTOGCI & ", "
   g_str_Parame = g_str_Parame & r_rst_MaeSol!SOLMAE_FMVBBP + r_rst_MaeSol!SOLMAE_PBPMTO & ", "
   g_str_Parame = g_str_Parame & r_rst_MaeSol!SOLMAE_BMSMTO & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "Error al insertar registro en tabla RPT_FICCRE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_rst_MaeCli.Close
   Set r_rst_MaeCli = Nothing
   
   r_rst_MaeSol.Close
   Set r_rst_MaeSol = Nothing
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "RPT_FICCRE"
   crp_Imprim.SelectionFormula = "{RPT_FICCRE.FICCRE_NUMSOL} = '" & gf_Formato_NumSol(moddat_g_str_NumSol) & "' AND {RPT_FICCRE.FICCRE_NOMTER} = '" & modgen_g_str_NombPC & "'"
   
   If r_int_TipEva = 1 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_01.RPT"
   ElseIf r_int_TipEva = 2 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_02.RPT"
   Else
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_03.RPT"
   End If
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1895
   grd_Listad.ColWidth(1) = 1375
   grd_Listad.ColWidth(2) = 1235
   grd_Listad.ColWidth(3) = 3455
   grd_Listad.ColWidth(4) = 1195
   grd_Listad.ColWidth(5) = 1195
   grd_Listad.ColWidth(6) = 1650
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
'   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   ipp_FecIni.Text = Format(date - CDate(15), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   ipp_FecIni.Enabled = p_Habilita
   ipp_FecFin.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   cmd_Imprim.Enabled = Not p_Habilita
   grd_Listad.Enabled = Not p_Habilita
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Imprim_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A, TRA_SEGUIM B "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = SEGUIM_NUMSOL "
   g_str_Parame = g_str_Parame & "   AND SEGUIM_CODINS = 21 "
   g_str_Parame = g_str_Parame & "   AND SEGUIM_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND SEGUIM_FECFIN >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND SEGUIM_FECFIN <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_NUMERO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         grd_Listad.Col = 3
         grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         
         grd_Listad.Col = 4
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = 21 "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            grd_Listad.Col = 5
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Genera!SEGUIM_FECFIN))
            
            grd_Listad.Col = 11
            grd_Listad.Text = g_rst_Genera!SEGUIM_FECINI
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         grd_Listad.Col = 6
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
         
         grd_Listad.Col = 9
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         
         grd_Listad.Col = 12
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_CYGTDO) & "-" & Trim(g_rst_Princi!SOLMAE_CYGNDO)
         'grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_CYGTDO), Trim(IIf(IsNull(g_rst_Princi!SOLMAE_CYGNDO), "", g_rst_Princi!SOLMAE_CYGNDO)))
         
         g_rst_Princi.MoveNext
      Loop
      
      'Ordenando por Nombre de Cliente
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 11, "N-")
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Len(Trim(pnl_Tit_FecSol.Tag)) = 0 Or pnl_Tit_FecSol.Tag = "D" Then
      pnl_Tit_FecSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_IngIns_Click()
   If Len(Trim(pnl_Tit_IngIns.Tag)) = 0 Or pnl_Tit_IngIns.Tag = "D" Then
      pnl_Tit_IngIns.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_IngIns.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub
