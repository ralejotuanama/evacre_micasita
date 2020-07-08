VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_RecSol_18 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4830
   ClientLeft      =   255
   ClientTop       =   2790
   ClientWidth     =   14820
   Icon            =   "EvaCre_frm_047.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14805
      _Version        =   65536
      _ExtentX        =   26114
      _ExtentY        =   8493
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin VB.CommandButton cmd_ConSol 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_047.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Consulta de Solicitud de Crédito Hipotecario"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14100
            Picture         =   "EvaCre_frm_047.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2835
         Left            =   30
         TabIndex        =   1
         Top             =   1920
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
         _ExtentY        =   5001
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
            Height          =   2475
            Left            =   30
            TabIndex        =   2
            Top             =   330
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   4366
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   3300
            TabIndex        =   4
            Top             =   60
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Rechazo"
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
            Left            =   1620
            TabIndex        =   5
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Rechazo"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   12570
            TabIndex        =   6
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Cliente"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   6240
            TabIndex        =   7
            Top             =   60
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Motivo de Rechazo"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   495
            Left            =   630
            TabIndex        =   9
            Top             =   60
            Width           =   5115
            _Version        =   65536
            _ExtentX        =   9022
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Lista de Solicitudes Rechazadas"
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
            Picture         =   "EvaCre_frm_047.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   11
            Top             =   60
            Width           =   12795
            _Version        =   65536
            _ExtentX        =   22569
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_RecSol_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli

   grd_Listad.ColWidth(0) = 1565
   grd_Listad.ColWidth(1) = 1685
   grd_Listad.ColWidth(2) = 2945
   grd_Listad.ColWidth(3) = 6335
   grd_Listad.ColWidth(4) = 1715
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   
   Call fs_Buscar
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)

   'Buscando Solicitudes Rechazadas como Cliente Titular
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
   
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
         
         grd_Listad.Col = 3
         grd_Listad.Text = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         
         grd_Listad.Col = 4
         grd_Listad.Text = moddat_gf_Consulta_ParDes("014", "1")
         
         g_rst_Princi.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscando Solicitudes Rechazadas como Cónyuge
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGNDO = '" & moddat_g_str_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
         
         grd_Listad.Col = 3
         grd_Listad.Text = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         
         grd_Listad.Col = 4
         grd_Listad.Text = moddat_gf_Consulta_ParDes("014", "2")
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

