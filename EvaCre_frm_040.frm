VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_RecSol_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3780
   ClientLeft      =   1740
   ClientTop       =   3750
   ClientWidth     =   11580
   Icon            =   "EvaCre_frm_040.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   6694
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_040.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_040.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10920
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   10290
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
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
            Caption         =   "Registro de Excepciones"
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
            Picture         =   "EvaCre_frm_040.frx":0890
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
         Height          =   1485
         Left            =   30
         TabIndex        =   14
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2619
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
         Begin VB.ComboBox cmb_TipAut 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1110
            Width           =   10035
         End
         Begin VB.TextBox txt_Observ 
            Height          =   1035
            Left            =   1440
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Text            =   "EvaCre_frm_040.frx":0B9A
            Top             =   60
            Width           =   10035
         End
         Begin VB.Label Label8 
            Caption         =   "Autorizado por:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   1110
            Width           =   1245
         End
         Begin VB.Label Label10 
            Caption         =   "Descripción de Excepción:"
            Height          =   585
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RecSol_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipAut_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_TipAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipAut_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_NumExc     As Integer
   
   If Len(Trim(txt_Observ.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción de la Excepción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If
   
   If cmb_TipAut.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Autorización.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipAut)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_str_Observ = txt_Observ.Text
   moddat_g_int_TipAut = cmb_TipAut.ItemData(cmb_TipAut.ListIndex)
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Select Case moddat_g_int_CodIns
      Case 11: pnl_TitPri.Caption = "Créditos Hipotecarios - Recepción de Solicitudes"
      Case 21: pnl_TitPri.Caption = "Créditos Hipotecarios - Evaluación Crediticia"
      Case 41: pnl_TitPri.Caption = "Créditos Hipotecarios - Tasación del Inmueble"
      Case 42: pnl_TitPri.Caption = "Créditos Hipotecarios - Evaluación de Seguros"
      Case 51: pnl_TitPri.Caption = "Créditos Hipotecarios - Evaluación Legal"
      Case 61: pnl_TitPri.Caption = "Créditos Hipotecarios - Pólizas de Seguros"
      Case 62: pnl_TitPri.Caption = "Créditos Hipotecarios - Trámites COFIDE"
      Case 72: pnl_TitPri.Caption = "Créditos Hipotecarios - Autorización de Desembolso"
   End Select
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   txt_Observ.Text = ""
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipAut, 1, "243")
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipAut)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

