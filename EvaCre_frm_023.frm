VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Rechaz_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   3270
   ClientTop       =   3945
   ClientWidth     =   10860
   Icon            =   "EvaCre_frm_023.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4185
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10905
      _Version        =   65536
      _ExtentX        =   19235
      _ExtentY        =   7382
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   780
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
            Left            =   10200
            Picture         =   "EvaCre_frm_023.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_023.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   1244
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
            Left            =   720
            TabIndex        =   17
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Recepción de Solicitudes"
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
            Left            =   720
            TabIndex        =   18
            Top             =   330
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Rechazo de Solicitud"
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
            Left            =   90
            Picture         =   "EvaCre_frm_023.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1845
         Left            =   30
         TabIndex        =   7
         Top             =   2280
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
         Begin VB.ComboBox cmb_MotRec 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9135
         End
         Begin VB.TextBox txt_Observ 
            Height          =   1395
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Text            =   "EvaCre_frm_023.frx":0B9A
            Top             =   390
            Width           =   9135
         End
         Begin VB.Label Label8 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label23 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
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
            Left            =   1620
            TabIndex        =   11
            Top             =   60
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9330
            TabIndex        =   13
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1620
            TabIndex        =   15
            Top             =   390
            Width           =   9135
            _Version        =   65536
            _ExtentX        =   16113
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
         Begin VB.Label Label11 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   7680
            TabIndex        =   14
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Rechaz_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_MotRec_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_MotRec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MotRec_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_MotRec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Motivo de Rechazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotRec)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de rechazar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_MotRec = cmb_MotRec.ItemData(cmb_MotRec.ListIndex)
   moddat_g_str_Observ = txt_Observ.Text
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli

   Select Case moddat_g_int_InsAct
      Case 11:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Recepción de Solicitudes"
      Case 21:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Evaluación Crediticia"
      Case 31:    pnl_TitPri.Caption = "Seguimiento de Solicitud de Crédito Hipotecario"
      Case 32:    pnl_TitPri.Caption = "Seguimiento de Solicitud de Crédito Hipotecario"
      Case 41:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Tasación del Inmueble"
      Case 42:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Evaluación de Seguros"
      Case 51:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Evaluación Legal"
      Case 61:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Pólizas de Seguro"
      Case 62:    pnl_TitPri.Caption = "Solicitud de Crédito Hipotecario - Trámites Cofide"
   End Select

   Call moddat_gs_Carga_MotRec(cmb_MotRec, moddat_g_int_InsAct)
   
   cmb_MotRec.ListIndex = -1
   txt_Observ.Text = ""
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub



