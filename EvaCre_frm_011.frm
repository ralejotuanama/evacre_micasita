VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.MDIForm frm_MnuPri_01 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8070
   ClientLeft      =   5580
   ClientTop       =   2355
   ClientWidth     =   10230
   Icon            =   "EvaCre_frm_011.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   -2147483633
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
      Begin VB.CommandButton cmd_ConSol 
         Height          =   585
         Left            =   2430
         Picture         =   "EvaCre_frm_011.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Consulta de Solicitud de Crédito Hipotecario"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ConCre 
         Height          =   585
         Left            =   3030
         Picture         =   "EvaCre_frm_011.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consulta de Crédito Hipotecario"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_TipCam 
         Height          =   585
         Left            =   1830
         Picture         =   "EvaCre_frm_011.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Tipo de Cambio"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_SimCre 
         Height          =   585
         Left            =   1230
         Picture         =   "EvaCre_frm_011.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Simulación de Créditos Hipotecarios"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_CamCon 
         Height          =   585
         Left            =   630
         Picture         =   "EvaCre_frm_011.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   30
         Picture         =   "EvaCre_frm_011.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   7680
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
      _ExtentY        =   688
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   0
      BevelInner      =   1
      Begin Threed.SSPanel pnl_EntDat 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "lm_db_db1 - prod1"
         ForeColor       =   32768
         BackColor       =   -2147483633
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
      Begin Threed.SSPanel pnl_NumVer 
         Height          =   315
         Left            =   3990
         TabIndex        =   9
         Top             =   30
         Width           =   2100
         _Version        =   65536
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "rev. 008-1028.1"
         ForeColor       =   32768
         BackColor       =   -2147483633
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
      Begin Threed.SSPanel pnl_TipCam 
         Height          =   315
         Left            =   6120
         TabIndex        =   10
         Top             =   30
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Tipo Cambio: Compra: S/. 2.00 - Venta: S/. 2.01"
         ForeColor       =   32768
         BackColor       =   -2147483633
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
   End
   Begin VB.Menu mnuHip 
      Caption         =   "Creditos Hipotecarios"
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Recepción de Solicitudes"
         Index           =   1
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Evaluación Crediticia"
         Index           =   2
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Simulación de Créditos"
         Index           =   4
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada (Recepción de Solicitudes)"
         Index           =   6
      End
      Begin VB.Menu mnuHip_Opcion 
         Caption         =   "Levantamiento de Aprobación Condicionada (Evaluación Crediticia)"
         Index           =   7
      End
   End
   Begin VB.Menu mnuClt 
      Caption         =   "Clientes"
      Begin VB.Menu mnuClt_Opcion 
         Caption         =   "Mantenimiento de Clientes"
         Index           =   1
      End
      Begin VB.Menu mnuClt_Opcion 
         Caption         =   "Actualización Datos Cliente"
         Index           =   2
      End
      Begin VB.Menu mnuClt_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuClt_Opcion 
         Caption         =   "Base Negativa"
         Index           =   4
      End
      Begin VB.Menu mnuClt_Opcion 
         Caption         =   "Clientes PEP"
         Index           =   5
      End
   End
   Begin VB.Menu mnuEpr 
      Caption         =   "Empresas"
      Begin VB.Menu mnuEpr_Opcion 
         Caption         =   "Mantenimiento de Empresas Empleadoras"
         Index           =   1
      End
      Begin VB.Menu mnuEpr_Opcion 
         Caption         =   "Evaluación de Empresas Empleadoras"
         Index           =   2
      End
      Begin VB.Menu mnuEpr_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEpr_Opcion 
         Caption         =   "Mantenimiento de Empresas Promotoras y/o Constructoras"
         Index           =   4
      End
      Begin VB.Menu mnuEpr_Opcion 
         Caption         =   "Evaluación de Empresas Promotoras y/o Constructoras"
         Index           =   5
      End
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Consultas"
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Solicitud de Crédito Hipotecario"
         Index           =   1
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Crédito Hipotecario"
         Index           =   2
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCon_Opcion 
         Caption         =   "Consulta de Tipo de Cambio"
         Index           =   4
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Evaluación Crediticia (Producto)"
         Index           =   1
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Evaluación Crediticia (Consejero Hipotecario)"
         Index           =   2
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes en Evaluación Crediticia (Proyecto Inmobiliario)"
         Index           =   3
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Observación Pendiente (Producto)"
         Index           =   5
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Observación Pendiente (Consejero Hipotecario)"
         Index           =   6
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Aprobación Condicionada Pendiente (Producto)"
         Index           =   8
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Aprobación Condicionada Pendiente (Consejero Hipotecario)"
         Index           =   9
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Excepción Aprobada (Producto)"
         Index           =   11
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes con Excepción Aprobada (Consejero Hipotecario)"
         Index           =   12
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Evaluadas (Producto)"
         Index           =   14
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Evaluadas (Consejero Hipotecario)"
         Index           =   15
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Evaluadas (Proyecto Inmobiliario)"
         Index           =   16
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Evaluadas (Tipo de Evaluación)"
         Index           =   17
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes Evaluadas (Analista Créditos)"
         Index           =   18
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Fichas de Evaluación Crediticia"
         Index           =   20
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Cartera de Créditos"
         Index           =   21
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Clientes Atrasados"
         Index           =   22
      End
      Begin VB.Menu mnuRpt_Opcion 
         Caption         =   "Reporte de Solicitudes por Clientes"
         Index           =   23
      End
   End
End
Attribute VB_Name = "frm_MnuPri_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_CamCon_Click()
   If modgen_g_str_CodUsu <> "DESARROLLO" Then
      frm_IdeUsu_02.Show 1
   End If
End Sub

Private Sub cmd_ConCre_Click()
   If mnuCon_Opcion(2).Enabled Then
      Call mnuCon_Opcion_Click(2)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_ConSol_Click()
   If mnuCon_Opcion(1).Enabled Then
      Call mnuCon_Opcion_Click(1)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   End If
End Sub

Private Sub cmd_SimCre_Click()
   If mnuHip_Opcion(4).Enabled Then
      Call mnuHip_Opcion_Click(4)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_TipCam_Click()
   If mnuCon_Opcion(4).Enabled Then
      Call mnuCon_Opcion_Click(4)
   Else
      MsgBox "Usuario no tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub MDIForm_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_HabSeg
   Call moddat_gf_Cargar_AgrPrd
   Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If MsgBox("¿Está seguro de salir de la Plataforma?", vbExclamation + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
      Call gs_Desconecta_Servidor
      End
   Else
      Cancel = True
   End If
End Sub

Private Sub mnuClt_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Mantenimiento de Clientes
         moddat_g_int_FlgCre = 1
         frm_MntCli_51.Show 1
      
      Case 2
         'Actualizacio de Datos del Cliente
         moddat_g_int_FlgCre = 3
         frm_MntCli_51.Show 1
      
      Case 4
         'Base Negativa
         moddat_g_str_TipPar = 1
         frm_BasNeg_01.Show 1
      
      Case 5
         'Base PEP
         moddat_g_str_TipPar = 2
         frm_BasNeg_01.Show 1
   End Select
End Sub

Private Sub mnuCon_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Consulta de Solicitud de Crédito Hipotecario
         frm_Con_SolHip_51.Show 1
         
      Case 2
         'Consulta de Crédito Hipotecario
         moddat_g_int_FlgPre = 1
         frm_Con_CreHip_51.Show 1
         
      Case 4
         'Consulta de Tipo de Cambio
         frm_ConTCa_01.Show 1
   End Select
End Sub

Private Sub mnuRpt_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Reporte de Solicitudes en Evaluación Crediticia (Producto)
         moddat_g_int_TipPan = 1
         frm_Rpt_EvaCre_01.Show 1
      Case 2
         'Reporte de Solicitudes en Evaluación Crediticia (Consejero Hipotecario)
         moddat_g_int_TipPan = 2
         frm_Rpt_EvaCre_01.Show 1
      Case 3
         'Reporte de Solicitudes en Evaluación Crediticia (Proyecto Inmobiliario)
         frm_Rpt_EvaCre_03.Show 1
      Case 5
         'Reporte de Solicitudes en Evaluación Crediticia Observadas (Producto)
         moddat_g_int_TipPan = 3
         frm_Rpt_EvaCre_01.Show 1
      Case 6
         'Reporte de Solicitudes en Evaluación Crediticia Observadas (Consejero Hipotecario)
         moddat_g_int_TipPan = 4
         frm_Rpt_EvaCre_01.Show 1
      Case 8
         'Reporte de Solicitudes en Evaluación Crediticia con Aprobación Condicionada (Producto)
         moddat_g_int_TipPan = 5
         frm_Rpt_EvaCre_01.Show 1
      Case 9
         'Reporte de Solicitudes en Evaluación Crediticia con Aprobación Condicionada (Consejero Hipotecario)
         moddat_g_int_TipPan = 6
         frm_Rpt_EvaCre_01.Show 1
      Case 11
         'Reporte de Solicitudes con Excepción Aprobada en Evaluación Crediticia (Producto)
         moddat_g_int_TipPan = 1
         frm_Rpt_EvaCre_08.Show 1
      Case 12
         'Reporte de Solicitudes con Excepción Aprobada en Evaluación Crediticia (Consejero Hipotecario)
         moddat_g_int_TipPan = 2
         frm_Rpt_EvaCre_08.Show 1
      Case 14
         'Reporte de Solicitudes Evaluadas (Producto)
         moddat_g_int_TipPan = 1
         frm_Rpt_EvaCre_10.Show 1
      Case 15
         'Reporte de Solicitudes Evaluadas (Consejero Hipotecario)
         moddat_g_int_TipPan = 2
         frm_Rpt_EvaCre_10.Show 1
      Case 16
         'Reporte de Solicitudes Evaluadas (Proyecto Inmobiliario)
         frm_Rpt_EvaCre_12.Show 1
      Case 17
         'Reporte de Solicitudes Evaluadas (Tipo Evaluación Crediticia)
         moddat_g_int_TipPan = 3
         frm_Rpt_EvaCre_10.Show 1
      Case 18
         'Reporte de Solicitudes Evaluadas (Analista de Creditos)
         moddat_g_int_TipPan = 4
         frm_Rpt_EvaCre_10.Show 1
      Case 20
         'Ficha de Evaluación Crediticia
         frm_Rpt_FicEva_01.Show 1
      Case 21
         'Ficha de Evaluación Crediticia
         frm_Rpt_EvaCre_14.Show 1
      Case 22
         'Reporte de Carter de Clientes Atrasados
         frm_Rpt_EvaCre_15.Show 1
      Case 23
         'Reporte de Solicitudes por Clientes
         frm_Rpt_EvaCre_16.Show 1
   End Select
End Sub

Private Sub mnuEpr_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Mantenimiento de Empresas Empleadoras
         frm_MntEmp_51.Show 1
   End Select
End Sub

Private Sub mnuHip_Opcion_Click(Index As Integer)
   Select Case Index
      Case 1
         'Recepción de Solicitudes
         frm_RecSol_51.Show 1
         
      Case 2
         'Evaluación Crediticia
         frm_EvaCre_61.Show 1
         
      Case 4
         'Simulación de Créditos Hipotecarios
         If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
            MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         frm_SimCre_11.Show 1
         
      Case 6
         'Levantamiento de Aprobación Condicionada (Recepción de Solicitudes)
         frm_LevCon_11.Show 1
      
      Case 7
         'Levantamiento de Aprobación Condicionada (Evaluación Crediticia)
         frm_LevCon_13.Show 1
   End Select
End Sub

Private Sub fs_HabSeg()
Dim r_int_Posici     As Integer
Dim r_str_CodMen     As String
Dim r_dbl_TipVta     As Double
Dim r_dbl_TipCom     As Double
   
   'pnl_Seg_NomUsu.Caption = modgen_g_str_CodUsu
   pnl_NumVer.Caption = modgen_g_str_NumRev
   pnl_EntDat.Caption = moddat_g_str_NomEsq & " - " & UCase(moddat_g_str_EntDat)
   r_dbl_TipVta = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 1)
   r_dbl_TipCom = moddat_gf_ObtieneTipCamDia(1, 2, Format(date, "yyyymmdd"), 2)
   pnl_TipCam.Caption = "Tipo de Cambio: Compra: S/. " & Format(r_dbl_TipCom, "###0.0000") & " - Venta: S/. " & Format(r_dbl_TipVta, "###0.0000")
   
   'Desactivando todas las opciones
   For r_int_Posici = 1 To mnuHip_Opcion.Count
      If mnuHip_Opcion(r_int_Posici).Caption <> "-" Then
         mnuHip_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici

   For r_int_Posici = 1 To mnuClt_Opcion.Count
      If mnuClt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuClt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici

   For r_int_Posici = 1 To mnuCon_Opcion.Count
      If mnuCon_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCon_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuEpr_Opcion.Count
      If mnuEpr_Opcion(r_int_Posici).Caption <> "-" Then
         mnuEpr_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuCon_Opcion.Count
      If mnuCon_Opcion(r_int_Posici).Caption <> "-" Then
         mnuCon_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   For r_int_Posici = 1 To mnuRpt_Opcion.Count
      If mnuRpt_Opcion(r_int_Posici).Caption <> "-" Then
         mnuRpt_Opcion(r_int_Posici).Enabled = False
      End If
   Next r_int_Posici
   
   'Verificando si todas las Opciones están habilitadas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM SEG_PLTOPC "
   g_str_Parame = g_str_Parame & " WHERE PLTOPC_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTOPC_FLGMEN = 2 "
   g_str_Parame = g_str_Parame & " ORDER BY PLTOPC_CODMEN ASC, PLTOPC_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTOPC_CODMEN)
            Select Case r_str_CodMen
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCLT": mnuClt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUEPR": mnuEpr_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTOPC_CODSUB)).Visible = IIf(g_rst_Princi!PLTOPC_SITUAC = 1, True, False)
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Verificando por Plantilla de Acceso
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM SEG_PLTPLA "
   g_str_Parame = g_str_Parame & " WHERE PLTPLA_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTPLA_TIPUSU = '" & CStr(modgen_g_int_TipUsu) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTPLA_CODMEN ASC, PLTPLA_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTPLA_CODMEN)
            Select Case r_str_CodMen
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCLT": mnuClt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUEPR": mnuEpr_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTPLA_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificando por Personalización de Opciones
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM SEG_PLTUSU "
   g_str_Parame = g_str_Parame & " WHERE PLTUSU_CODPLT = '" & UCase(App.EXEName) & "' "
   g_str_Parame = g_str_Parame & "   AND PLTUSU_CODUSU = '" & CStr(modgen_g_str_CodUsu) & "' "
   g_str_Parame = g_str_Parame & " ORDER BY PLTUSU_CODMEN ASC, PLTUSU_CODSUB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
         Do While Not g_rst_Princi.EOF And r_str_CodMen = Trim(g_rst_Princi!PLTUSU_CODMEN)
            Select Case r_str_CodMen
               Case "MNUHIP": mnuHip_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCLT": mnuClt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUEPR": mnuEpr_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNUCON": mnuCon_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
               Case "MNURPT": mnuRpt_Opcion(CInt(g_rst_Princi!PLTUSU_CODSUB)).Enabled = True
            End Select
            
            g_rst_Princi.MoveNext
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
