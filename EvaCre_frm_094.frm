VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_EvaCre_72 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11595
   Icon            =   "EvaCre_frm_094.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3500
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   6174
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_094.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Aprobar Solicitud Jefe"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_094.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_094.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Comentario"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   630
            TabIndex        =   7
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
            Left            =   630
            TabIndex        =   8
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia"
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
            Picture         =   "EvaCre_frm_094.frx":0B9A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   1185
         Left            =   30
         TabIndex        =   14
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2090
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
         Begin VB.TextBox txt_Observ 
            Height          =   1035
            Left            =   1440
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Text            =   "EvaCre_frm_094.frx":0EA4
            Top             =   60
            Width           =   10035
         End
         Begin VB.Label Label10 
            Caption         =   "Comentario:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_72"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   txt_Observ.Text = ""
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Aprueb_Click()
Dim r_str_Parame    As String
Dim r_rst_Princi    As ADODB.Recordset
 
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " UPDATE TRA_EVACRE "
   r_str_Parame = r_str_Parame & "    SET EVACRE_FLGJFT = 1, "
   r_str_Parame = r_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "        SEGFECACT = " & Format(date, "yyyymmdd") & ",  "
   r_str_Parame = r_str_Parame & "        SEGHORACT = " & Format(Time, "HHmmss") & ",  "
   r_str_Parame = r_str_Parame & "        SEGPLTACT = '" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "        SEGTERACT = '" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "' "
   r_str_Parame = r_str_Parame & "  WHERE EVACRE_NUMSOL = " & moddat_g_str_NumSol
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 2) Then
      Exit Sub
   End If
   
   If Len(TRIM(txt_Observ.Text)) > 0 Then
      'Grabando en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 42, 0, TRIM(txt_Observ.Text), 0, 0) Then
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_Grabar_Click()
   If Len(TRIM(txt_Observ.Text)) = 0 Then
      MsgBox "Debe ingresar la Observación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 42, 0, TRIM(txt_Observ.Text), 0, 0) Then
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub
