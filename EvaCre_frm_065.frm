VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RecSol_55 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   3750
   ClientLeft      =   4785
   ClientTop       =   2340
   ClientWidth     =   11580
   Icon            =   "EvaCre_frm_065.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3795
      Left            =   0
      TabIndex        =   7
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
         TabIndex        =   8
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_065.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_065.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   30
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cr�ditos Hipotecarios - Recepci�n de Solicitudes"
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
            TabIndex        =   11
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
            Picture         =   "EvaCre_frm_065.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   12
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
            Left            =   1265
            TabIndex        =   5
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
            Left            =   1260
            TabIndex        =   6
            Top             =   390
            Width           =   10095
            _Version        =   65536
            _ExtentX        =   17806
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   1485
         Left            =   30
         TabIndex        =   15
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
         Begin VB.ComboBox cmb_Motivo 
            Height          =   315
            ItemData        =   "EvaCre_frm_065.frx":0B9A
            Left            =   6060
            List            =   "EvaCre_frm_065.frx":0B9C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1110
            Visible         =   0   'False
            Width           =   5325
         End
         Begin VB.TextBox txt_Observ 
            Height          =   1035
            Left            =   1260
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Text            =   "EvaCre_frm_065.frx":0B9E
            Top             =   60
            Width           =   10125
         End
         Begin VB.ComboBox cmb_TipAut 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1110
            Width           =   3975
         End
         Begin VB.Label lbl_motivo 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   5370
            TabIndex        =   18
            Top             =   1170
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label10 
            Caption         =   "Descripci�n de Excepci�n:"
            Height          =   585
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label Label8 
            Caption         =   "Autorizado por:"
            Height          =   225
            Left            =   60
            TabIndex        =   16
            Top             =   1170
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RecSol_55"
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
      MsgBox "Debe ingresar la Descripci�n de la Excepci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If
   
   If cmb_TipAut.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Autorizaci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipAut)
      Exit Sub
   End If
   
   If modgen_g_int_FlgExc = 1 Then
      If cmb_Motivo.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Motivo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Motivo)
         Exit Sub
      End If
      
      moddat_g_int_MotExc = cmb_Motivo.ItemData(cmb_Motivo.ListIndex)
   End If
   
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
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
   modgen_g_int_FlgExc = 0
   
   Select Case moddat_g_int_CodIns
      Case 11: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Recepci�n de Solicitudes"
      Case 21: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Evaluaci�n Crediticia"
               modgen_g_int_FlgExc = 1
      Case 41: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Tasaci�n del Inmueble"
      Case 42: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Evaluaci�n de Seguros"
      Case 51: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Evaluaci�n Legal"
      Case 61: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - P�lizas de Seguros"
      Case 62: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Tr�mites COFIDE"
      Case 72: pnl_TitPri.Caption = "Solicitud de Cr�dito Hipotecario - Autorizaci�n de Desembolso"
   End Select
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   txt_Observ.Text = ""
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipAut, 1, "243")
   
   If modgen_g_int_FlgExc = 1 Then
      Call moddat_gs_Carga_LisIte_Combo(cmb_Motivo, 1, "042")
      lbl_motivo.Visible = True
      cmb_Motivo.Visible = True
   End If
End Sub
 

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipAut)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$�!�@#=?�+*" & Chr(10))
   End If
End Sub


