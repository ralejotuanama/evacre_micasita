VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_BasNeg_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   1695
   ClientTop       =   2370
   ClientWidth     =   11760
   Icon            =   "EvaCre_frm_027.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7845
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11775
      _Version        =   65536
      _ExtentX        =   20770
      _ExtentY        =   13838
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_027.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_027.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11040
            Picture         =   "EvaCre_frm_027.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   1230
            Picture         =   "EvaCre_frm_027.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Consulta Registro"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   6315
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
         _ExtentY        =   11139
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
            Height          =   5955
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   10504
            _Version        =   393216
            Rows            =   12
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1950
            TabIndex        =   8
            Top             =   60
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Docum. Identidad"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   9420
            TabIndex        =   10
            Top             =   60
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   6720
            TabIndex        =   13
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ingreso"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   8070
            TabIndex        =   14
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Salida"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   480
            Left            =   630
            TabIndex        =   12
            Top             =   90
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Base Negativa de Clientes"
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
            Left            =   90
            Picture         =   "EvaCre_frm_027.frx":0D6C
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_BasNeg_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 1
   frm_BasNeg_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_Situac     As Integer

   grd_Listad.Col = 0
   moddat_g_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 5)
         
   grd_Listad.Col = 5
   r_int_Situac = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)

   If r_int_Situac = 2 Then
      If (moddat_g_str_TipPar = 1) Then
          MsgBox "Este registro ya ha sido retirado de la Base Negativa.", vbInformation, modgen_g_str_NomPlt
      ElseIf (moddat_g_str_TipPar = 2) Then
          MsgBox "Este registro ya ha sido retirado de la Base PEP.", vbInformation, modgen_g_str_NomPlt
      End If
      Exit Sub
   End If

   moddat_g_int_FlgAct = 2
   moddat_g_int_FlgGrb = 1
   frm_BasNeg_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Consul_Click()
   grd_Listad.Col = 0
   moddat_g_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 5)
         
   grd_Listad.Col = 2
   moddat_g_str_Descri = Format(CDate(grd_Listad.Text), "yyyymmdd")
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_BasNeg_04.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   If (moddat_g_str_TipPar = 1) Then
       pnl_Titulo.Caption = "Base Negativa de Clientes"
   ElseIf (moddat_g_str_TipPar = 2) Then
       pnl_Titulo.Caption = "Base PEP de Clientes"
   End If
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1845
   grd_Listad.ColWidth(1) = 4785
   grd_Listad.ColWidth(2) = 1365
   grd_Listad.ColWidth(3) = 1365
   grd_Listad.ColWidth(4) = 1905
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Consul.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT BASNEG_CLITDO, BASNEG_CLINDO, BASNEG_APEPAT, BASNEG_APEMAT, BASNEG_NOMBRE, BASNEG_FECNEG, BASNEG_FECLEV, BASNEG_SITUAC, PARDES_DESCRI "
   g_str_Parame = g_str_Parame & "   FROM CRE_BASNEG "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES ON PARDES_CODGRP = '013' AND PARDES_CODITE = BASNEG_SITUAC "
   g_str_Parame = g_str_Parame & "  WHERE BASNEG_TIPBAS = " & moddat_g_str_TipPar
   g_str_Parame = g_str_Parame & "  ORDER BY BASNEG_APEPAT ASC, BASNEG_APEMAT ASC, BASNEG_NOMBRE ASC, BASNEG_FECNEG DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!BASNEG_CLITDO) & " - " & Trim(g_rst_Princi!BASNEG_CLINDO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!BASNEG_APEPAT) & " " & Trim(g_rst_Princi!BASNEG_APEMAT) & " " & Trim(g_rst_Princi!BASNEG_NOMBRE)
      
      grd_Listad.Col = 2
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!BASNEG_FECNEG))
      
      If g_rst_Princi!BASNEG_FECLEV > 0 Then
         grd_Listad.Col = 3
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!BASNEG_FECLEV))
      End If
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!PARDES_DESCRI)   'moddat_gf_Consulta_ParDes("013", CStr(g_rst_Princi!BASNEG_SITUAC))
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!BASNEG_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Consul.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Consul_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub


