VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_BasNeg_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   1890
   ClientTop       =   2160
   ClientWidth     =   11775
   Icon            =   "EvaCre_frm_029.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _Version        =   65536
      _ExtentX        =   20770
      _ExtentY        =   9499
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
         Height          =   2085
         Left            =   30
         TabIndex        =   1
         Top             =   1470
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
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
         Begin VB.TextBox txt_ObsIng 
            Height          =   945
            Left            =   2010
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1080
            Width           =   9615
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   3
            Top             =   90
            Width           =   9615
            _Version        =   65536
            _ExtentX        =   16960
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   2010
            TabIndex        =   4
            Top             =   420
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_MotIng 
            Height          =   315
            Left            =   2010
            TabIndex        =   5
            Top             =   750
            Width           =   9615
            _Version        =   65536
            _ExtentX        =   16960
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
         Begin VB.Label lbl_Etique 
            Caption         =   "Cliente:"
            Height          =   315
            Index           =   74
            Left            =   90
            TabIndex        =   9
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   8
            Top             =   420
            Width           =   1515
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Motivo Ingreso:"
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   7
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones de Ingreso:"
            Height          =   495
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   1080
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   10
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11070
            Picture         =   "EvaCre_frm_029.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1725
         Left            =   30
         TabIndex        =   12
         Top             =   3600
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
         _ExtentY        =   3043
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
         Begin VB.TextBox txt_ObsSal 
            Height          =   945
            Left            =   2010
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   720
            Width           =   9615
         End
         Begin Threed.SSPanel pnl_FecSal 
            Height          =   315
            Left            =   2010
            TabIndex        =   18
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin Threed.SSPanel pnl_MotSal 
            Height          =   315
            Left            =   2010
            TabIndex        =   19
            Top             =   390
            Width           =   9615
            _Version        =   65536
            _ExtentX        =   16960
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
         Begin VB.Label Label3 
            Caption         =   "Fecha Salida:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   60
            Width           =   1785
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones de Salida:"
            Height          =   465
            Index           =   82
            Left            =   90
            TabIndex        =   15
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Motivo Salida:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   390
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   17
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
            Height          =   225
            Left            =   720
            TabIndex        =   20
            Top             =   60
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   397
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   225
            Left            =   720
            TabIndex        =   21
            Top             =   330
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Consulta"
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
            Picture         =   "EvaCre_frm_029.frx":044E
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_BasNeg_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmd_Salida)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   If (moddat_g_str_TipPar = 1) Then
       pnl_Titulo.Caption = "Base Negativa de Clientes"
   ElseIf (moddat_g_str_TipPar = 2) Then
       pnl_Titulo.Caption = "Base PEP de Clientes"
   End If
End Sub

Private Sub fs_Buscar()
   pnl_FecSal.Caption = ""
   pnl_MotSal.Caption = ""
   txt_ObsSal.Text = ""

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_BASNEG "
   g_str_Parame = g_str_Parame & " WHERE BASNEG_TIPBAS = " & moddat_g_str_TipPar
   g_str_Parame = g_str_Parame & "   AND BASNEG_CLITDO = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND BASNEG_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND BASNEG_FECNEG = " & moddat_g_str_Descri & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      pnl_NomCli.Caption = CStr(g_rst_Genera!BASNEG_CLITDO) & " - " & Trim(g_rst_Genera!BASNEG_CLINDO) & " / " & Trim(g_rst_Genera!BASNEG_APEPAT) & " " & Trim(g_rst_Genera!BASNEG_APEMAT) & " " & Trim(g_rst_Genera!BASNEG_NOMBRE)
      pnl_FecIng.Caption = gf_FormatoFecha(CStr(g_rst_Genera!BASNEG_FECNEG))
      If (moddat_g_str_TipPar = 1) Then
          pnl_MotIng.Caption = moddat_gf_Consulta_ParDes("105", CStr(g_rst_Genera!BASNEG_CODNEG))
      ElseIf (moddat_g_str_TipPar = 2) Then
          pnl_MotIng.Caption = moddat_gf_Consulta_ParDes("116", CStr(g_rst_Genera!BASNEG_CODNEG))
      End If
      If Not IsNull(g_rst_Genera!BASNEG_OBSNEG) Then
         txt_ObsIng.Text = Trim(g_rst_Genera!BASNEG_OBSNEG)
      End If
   
      If g_rst_Genera!BASNEG_FECLEV > 0 Then
         pnl_FecSal.Caption = gf_FormatoFecha(CStr(g_rst_Genera!BASNEG_FECLEV))
         If (moddat_g_str_TipPar = 1) Then
             pnl_MotSal.Caption = moddat_gf_Consulta_ParDes("106", CStr(g_rst_Genera!BASNEG_CODLEV))
         ElseIf (moddat_g_str_TipPar = 2) Then
             pnl_MotSal.Caption = moddat_gf_Consulta_ParDes("117", CStr(g_rst_Genera!BASNEG_CODLEV))
         End If
         If Not IsNull(g_rst_Genera!BASNEG_OBSLEV) Then
            txt_ObsSal.Text = Trim(g_rst_Genera!BASNEG_OBSLEV)
         End If
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub txt_ObsIng_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSal_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
