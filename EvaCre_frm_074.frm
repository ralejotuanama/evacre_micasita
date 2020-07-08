VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_EvaCre_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   2280
   ClientLeft      =   6120
   ClientTop       =   3555
   ClientWidth     =   7740
   Icon            =   "EvaCre_frm_074.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
         Begin Threed.SSPanel pnl_Tit01 
            Height          =   285
            Left            =   660
            TabIndex        =   7
            Top             =   30
            Width           =   6915
            _Version        =   65536
            _ExtentX        =   12197
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Evaluación Crediticia"
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
         Begin Threed.SSPanel pnl_Tit02 
            Height          =   285
            Left            =   660
            TabIndex        =   8
            Top             =   300
            Width           =   6975
            _Version        =   65536
            _ExtentX        =   12303
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
            Picture         =   "EvaCre_frm_074.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_074.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7050
            Picture         =   "EvaCre_frm_074.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_074.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6570
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
         Height          =   795
         Left            =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6585
         End
         Begin VB.CheckBox chk_TipVal 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1020
            TabIndex        =   1
            Top             =   420
            Width           =   6375
         End
         Begin VB.Label pnl_Tit03 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaCre_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   If moddat_g_int_TipPan = 1 Then
      'Reporte de solicitudes en evaluacion crediticia - producto
      pnl_Tit01.Caption = "Reporte de Solicitudes en Evaluación Crediticia"
      pnl_Tit02.Caption = "Producto"
      pnl_Tit03.Caption = "Producto:"
      chk_TipVal.Caption = "Todos los Productos"
      Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   ElseIf moddat_g_int_TipPan = 2 Then
      'Reporte de solicitudes en evaluacion crediticia - consejero hipotecario
      pnl_Tit01.Caption = "Reporte de Solicitudes en Evaluación Crediticia"
      pnl_Tit02.Caption = "Consejero Hipotecario"
      pnl_Tit03.Caption = "Consejero:"
      chk_TipVal.Caption = "Todos los Consejeros Hipotecarios"
      Call moddat_gs_Carga_EjecMC(cmb_Produc, l_arr_Produc, 121)
   ElseIf moddat_g_int_TipPan = 3 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - producto
      pnl_Tit01.Caption = "Reporte de Solicitudes Observadas en Evaluación Crediticia"
      pnl_Tit02.Caption = "Producto"
      pnl_Tit03.Caption = "Producto:"
      chk_TipVal.Caption = "Todos los Productos"
      Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   ElseIf moddat_g_int_TipPan = 4 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - consejero hipotecario
      pnl_Tit01.Caption = "Reporte de Solicitudes Observadas en Evaluación Crediticia"
      pnl_Tit02.Caption = "Consejero Hipotecario"
      pnl_Tit03.Caption = "Consejero:"
      chk_TipVal.Caption = "Todos los Consejeros Hipotecarios"
      Call moddat_gs_Carga_EjecMC(cmb_Produc, l_arr_Produc, 121)
   ElseIf moddat_g_int_TipPan = 5 Then
      'Reporte de solicitudes con aprobacion condicionada en evalucion crediticia - producto
      pnl_Tit01.Caption = "Reporte de Solicitudes con Aprobación Condicionada en Evaluación Crediticia"
      pnl_Tit02.Caption = "Producto"
      pnl_Tit03.Caption = "Producto:"
      chk_TipVal.Caption = "Todos los Productos"
      Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   ElseIf moddat_g_int_TipPan = 6 Then
      'Reporte de solicitudes con aprobacion condicionada en evaluacion crediticia - consejero hipotecario
      pnl_Tit01.Caption = "Reporte de Solicitudes con Aprobación Condicionada en Evaluación Crediticia"
      pnl_Tit02.Caption = "Consejero Hipotecario"
      pnl_Tit03.Caption = "Consejero:"
      chk_TipVal.Caption = "Todos los Consejeros Hipotecarios"
      Call moddat_gs_Carga_EjecMC(cmb_Produc, l_arr_Produc, 121)
   End If
End Sub

Private Sub chk_TipVal_Click()
   If chk_TipVal.Value = 1 Then
      cmb_Produc.ListIndex = -1
      cmb_Produc.Enabled = False
   ElseIf chk_TipVal.Value = 0 Then
      cmb_Produc.Enabled = True
      Call gs_SetFocus(cmb_Produc)
   End If
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If chk_TipVal.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar un " & LCase(Trim(pnl_Tit03.Caption)), vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   If moddat_g_int_TipPan = 1 Then
      'Reporte de solicitudes en evaluacion crediticia - producto
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_Tramit(21, 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, "")
      Else
         Call modmip_gs_Exc_Tramit(21, 1, "", "")
      End If
   ElseIf moddat_g_int_TipPan = 2 Then
      'Reporte de solicitudes en evaluacion crediticia - consejero hipotecario
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_Tramit(21, 2, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, "")
      Else
         Call modmip_gs_Exc_Tramit(21, 2, "", "")
      End If
   ElseIf moddat_g_int_TipPan = 3 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - producto
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_Observ(21, 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo)
      Else
         Call modmip_gs_Exc_Observ(21, 1, "")
      End If
   ElseIf moddat_g_int_TipPan = 4 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - consejero hipotecario
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_Observ(21, 2, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo)
      Else
         Call modmip_gs_Exc_Observ(21, 2, "")
      End If
   ElseIf moddat_g_int_TipPan = 5 Then
      'Reporte de solicitudes con aprobacion condicionada en evalucion crediticia - producto
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_AprCon(21, 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo)
      Else
         Call modmip_gs_Exc_AprCon(21, 1, "")
      End If
   ElseIf moddat_g_int_TipPan = 6 Then
      'Reporte de solicitudes con aprobacion condicionada en evaluacion crediticia - consejero hipotecario
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Exc_AprCon(21, 2, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo)
      Else
         Call modmip_gs_Exc_AprCon(21, 2, "")
      End If
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If chk_TipVal.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar un " & LCase(Trim(pnl_Tit03.Caption)), vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   If moddat_g_int_TipPan = 1 Then
      'Reporte de solicitudes en evaluacion crediticia - producto
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_EvaIns("CRE_EVAHIP_01.RPT", 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_EvaIns("CRE_EVAHIP_01.RPT", 1, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_01.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_01.RPT"
      crp_Imprim.Action = 1
   ElseIf moddat_g_int_TipPan = 2 Then
      'Reporte de solicitudes en evaluacion crediticia - consejero hipotecario
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_EvaIns("CRE_EVAHIP_02.RPT", 2, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_EvaIns("CRE_EVAHIP_02.RPT", 2, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_02.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_02.RPT"
      crp_Imprim.Action = 1
   ElseIf moddat_g_int_TipPan = 3 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - producto
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_EvaObs("CRE_EVAHIP_04.RPT", 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_EvaObs("CRE_EVAHIP_04.RPT", 1, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_04.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_04.RPT"
      crp_Imprim.Action = 1
   ElseIf moddat_g_int_TipPan = 4 Then
      'Reporte de solicitudes observadas en evaluacion crediticia - consejero hipotecario
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_EvaObs("CRE_EVAHIP_05.RPT", 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_EvaObs("CRE_EVAHIP_05.RPT", 1, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_05.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_05.RPT"
      crp_Imprim.Action = 1
   ElseIf moddat_g_int_TipPan = 5 Then
      'Reporte de solicitudes con aprobacion condicionada en evalucion crediticia - producto
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_06.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_06.RPT"
      crp_Imprim.Action = 1
   ElseIf moddat_g_int_TipPan = 6 Then
      'Reporte de solicitudes con aprobacion condicionada en evaluacion crediticia - consejero hipotecario
      Screen.MousePointer = 11
      If chk_TipVal.Value = 0 Then
         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_07.RPT", 1, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 21, "")
      Else
         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_07.RPT", 1, "", 21, "")
      End If
      Screen.MousePointer = 0
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_07.RPT'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_07.RPT"
      crp_Imprim.Action = 1
   End If
End Sub

