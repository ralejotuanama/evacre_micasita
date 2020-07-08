VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_EvaEmp_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   6690
   ClientLeft      =   1395
   ClientTop       =   1995
   ClientWidth     =   12990
   Icon            =   "EvaCre_frm_009.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6705
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13005
      _Version        =   65536
      _ExtentX        =   22939
      _ExtentY        =   11827
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
         Height          =   735
         Left            =   30
         TabIndex        =   12
         Top             =   5910
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
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
         Begin VB.CommandButton cmd_VerEva 
            Height          =   675
            Left            =   12180
            Picture         =   "EvaCre_frm_009.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Evaluación de Empresas"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4305
         Left            =   30
         TabIndex        =   7
         Top             =   1560
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
         _ExtentY        =   7594
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
            Height          =   3945
            Left            =   30
            TabIndex        =   4
            Top             =   330
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   6959
            _Version        =   393216
            Rows            =   20
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento de Identidad"
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
            Left            =   2100
            TabIndex        =   9
            Top             =   60
            Width           =   6045
            _Version        =   65536
            _ExtentX        =   10663
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
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
            Left            =   8130
            TabIndex        =   13
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ult. Evaluac."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   10830
            TabIndex        =   14
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situac. Evaluac."
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   9480
            TabIndex        =   17
            Top             =   60
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Emisión Eval."
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
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
            TabIndex        =   11
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
            Picture         =   "EvaCre_frm_009.frx":08D6
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   12915
         _Version        =   65536
         _ExtentX        =   22781
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
         Begin VB.ComboBox cmb_Clasif 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   3345
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10800
            Picture         =   "EvaCre_frm_009.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11490
            Picture         =   "EvaCre_frm_009.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12180
            Picture         =   "EvaCre_frm_009.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Clasificación:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   180
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frm_EvaEmp_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Clasif_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Clasif_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Clasif_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Clasif.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación de las Empresas a mostrar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Clasif)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_Clasif)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerEva_Click()
   grd_Listad.Col = 0
   
   moddat_g_int_TDoEmp = CInt(Left(grd_Listad.Text, 1))
   moddat_g_str_NDoEmp = Mid(grd_Listad.Text, 3)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_EvaEmp_02.Show 1
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call fs_Inicio
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt & " - Evaluación de Empresas"
   
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerEva_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 2040
   grd_Listad.ColWidth(1) = 6030
   grd_Listad.ColWidth(2) = 1350
   grd_Listad.ColWidth(3) = 1350
   grd_Listad.ColWidth(4) = 1740
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Clasif, 1, "016")
End Sub

Private Sub fs_Limpia()
   cmb_Clasif.ListIndex = -1
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_Clasif.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
   cmd_VerEva.Enabled = Not p_Habilita
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_CLASIF = " & CStr(cmb_Clasif.ItemData(cmb_Clasif.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_RAZSOC ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontraron Empresas registradas para esta Clasificación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Clasif)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   Call fs_Activa(False)
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!DATGEN_EMPTDO) & "-" & Trim(g_rst_Princi!DATGEN_EMPNDO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
      
      If g_rst_Princi!DATGEN_ULTEVA > 0 Then
         grd_Listad.Col = 2
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_ULTEVA))
      End If
      
      If g_rst_Princi!DATGEN_EMIEVA > 0 Then
         grd_Listad.Col = 3
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_EMIEVA))
      End If
      
      grd_Listad.Col = 4
      If g_rst_Princi!DATGEN_SITEVA = 1 Then
         grd_Listad.Text = "PENDIENTE"
      Else
         grd_Listad.Text = "EVALUADA"
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub SSPanel8_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel9_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub
