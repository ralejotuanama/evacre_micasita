VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_BasNeg_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   2220
   ClientTop       =   4185
   ClientWidth     =   11760
   Icon            =   "EvaCre_frm_028.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5385
      Left            =   0
      TabIndex        =   5
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
         TabIndex        =   12
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
            TabIndex        =   19
            Top             =   1080
            Width           =   9615
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   13
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
            TabIndex        =   15
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
            TabIndex        =   17
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
            Caption         =   "Observaciones de Ingreso:"
            Height          =   495
            Index           =   2
            Left            =   90
            TabIndex        =   20
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Motivo Ingreso:"
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   18
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   420
            Width           =   1515
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cliente:"
            Height          =   315
            Index           =   74
            Left            =   90
            TabIndex        =   14
            Top             =   90
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11070
            Picture         =   "EvaCre_frm_028.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_028.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1725
         Left            =   30
         TabIndex        =   7
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
         Begin VB.ComboBox cmb_MotSal 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   9615
         End
         Begin VB.TextBox txt_Observ 
            Height          =   945
            Left            =   2010
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   9615
         End
         Begin EditLib.fpDateTime ipp_FecSal 
            Height          =   315
            Left            =   2010
            TabIndex        =   0
            Top             =   60
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label7 
            Caption         =   "Motivo Salida:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones de Salida:"
            Height          =   465
            Index           =   82
            Left            =   90
            TabIndex        =   9
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Salida:"
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Top             =   60
            Width           =   1785
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
         Begin Threed.SSPanel pnl_Titulo1 
            Height          =   225
            Left            =   660
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_Titulo2 
            Height          =   225
            Left            =   660
            TabIndex        =   22
            Top             =   330
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Salida de Clientes de Base Negativa"
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
            Picture         =   "EvaCre_frm_028.frx":0890
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_BasNeg_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   If CDate(ipp_FecSal.Text) < CDate(pnl_FecIng.Caption) Then
      If (moddat_g_str_TipPar = 1) Then
          MsgBox "La fecha de salida de Base Negativa no puede ser menor a la fecha de ingreso.", vbExclamation, modgen_g_str_NomPlt
      ElseIf (moddat_g_str_TipPar = 2) Then
          MsgBox "La fecha de salida de Base PEP no puede ser menor a la fecha de ingreso.", vbExclamation, modgen_g_str_NomPlt
      End If
      Call gs_SetFocus(ipp_FecSal)
      Exit Sub
   End If

   If cmb_MotSal.ListIndex = -1 Then
      If (moddat_g_str_TipPar = 1) Then
          MsgBox "Debe seleccionar el motivo de salida de Base Negativa.", vbExclamation, modgen_g_str_NomPlt
      ElseIf (moddat_g_str_TipPar = 2) Then
          MsgBox "Debe seleccionar el motivo de salida de Base PEP.", vbExclamation, modgen_g_str_NomPlt
      End If
      Call gs_SetFocus(cmb_MotSal)
      Exit Sub
   End If
  
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando en Maestro de Ejecutivos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CRE_BASNEG_SALIDA ("
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecSal.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MotSal.ItemData(cmb_MotSal.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      g_str_Parame = g_str_Parame & moddat_g_str_TipPar & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      Screen.MousePointer = 0
   Loop
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   Call gs_CentraForm(Me)
         
   Call gs_SetFocus(ipp_FecSal)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   If (moddat_g_str_TipPar = 1) Then
       pnl_Titulo1.Caption = "Base Negativa de Clientes"
       pnl_Titulo2.Caption = "Salida de Clientes de Base Negativa"
       Call moddat_gs_Carga_LisIte_Combo(cmb_MotSal, 1, "106")
   ElseIf (moddat_g_str_TipPar = 2) Then
       pnl_Titulo1.Caption = "Base PEP de Clientes"
       pnl_Titulo2.Caption = "Salida de Clientes de Base PEP"
       Call moddat_gs_Carga_LisIte_Combo(cmb_MotSal, 1, "117")
   End If
End Sub

Private Sub fs_Limpia()
   ipp_FecSal.Text = Format(date, "dd/mm/yyyy")
   cmb_MotSal.ListIndex = -1
   txt_Observ.Text = ""
End Sub

Private Sub fs_Buscar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_BASNEG "
   g_str_Parame = g_str_Parame & " WHERE BASNEG_TIPBAS = " & moddat_g_str_TipPar
   g_str_Parame = g_str_Parame & "   AND BASNEG_CLITDO = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND BASNEG_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND BASNEG_SITUAC = 1 "

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
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmb_MotSal_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_MotSal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MotSal_Click
   End If
End Sub

Private Sub ipp_FecSal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MotSal)
   End If
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

Private Sub txt_ObsIng_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
