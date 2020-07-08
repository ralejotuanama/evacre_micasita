VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_BasNeg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4905
   ClientLeft      =   1455
   ClientTop       =   3330
   ClientWidth     =   11760
   Icon            =   "EvaCre_frm_026.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4905
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11775
      _Version        =   65536
      _ExtentX        =   20770
      _ExtentY        =   8652
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
         Height          =   675
         Left            =   30
         TabIndex        =   11
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_026.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11070
            Picture         =   "EvaCre_frm_026.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3375
         Left            =   30
         TabIndex        =   12
         Top             =   1470
         Width           =   11685
         _Version        =   65536
         _ExtentX        =   20611
         _ExtentY        =   5953
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
            Height          =   945
            Left            =   2010
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2370
            Width           =   9615
         End
         Begin VB.ComboBox cmb_MotIng 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2040
            Width           =   9615
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1380
            Width           =   3285
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   3285
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   3285
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3285
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   3285
         End
         Begin EditLib.fpDateTime ipp_FecIng 
            Height          =   315
            Left            =   2010
            TabIndex        =   5
            Top             =   1710
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
         Begin VB.Label Label3 
            Caption         =   "Fecha Ingreso:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones:"
            Height          =   225
            Index           =   82
            Left            =   90
            TabIndex        =   21
            Top             =   2370
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Motivo Ingreso:"
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Top             =   2040
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Número de DOI:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo de DOI:"
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   1380
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   1050
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   720
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   19
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
            TabIndex        =   20
            Top             =   90
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
            TabIndex        =   23
            Top             =   360
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Ingreso de Clientes a Base Negativa"
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
            Picture         =   "EvaCre_frm_026.frx":0890
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_BasNeg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de DOI.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de DOI.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
      If Len(Trim(txt_NumDoc.Text)) > 8 Then
         MsgBox "Debe ingresar el Número de DOI correctamente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   End If
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   If cmb_MotIng.ListIndex = -1 Then
      If (moddat_g_str_TipPar = 1) Then
          MsgBox "Debe seleccionar el motivo de ingreso a Negatividad.", vbExclamation, modgen_g_str_NomPlt
      ElseIf (moddat_g_str_TipPar = 2) Then
          MsgBox "Debe seleccionar el motivo de ingreso a PEP.", vbExclamation, modgen_g_str_NomPlt
      End If
      Call gs_SetFocus(cmb_MotIng)
      Exit Sub
   End If
  
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_BASNEG "
      g_str_Parame = g_str_Parame & " WHERE BASNEG_TIPBAS = " & moddat_g_str_TipPar
      g_str_Parame = g_str_Parame & "   AND BASNEG_CLITDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
      g_str_Parame = g_str_Parame & "   AND BASNEG_CLINDO = '" & txt_NumDoc.Text & "' "
      g_str_Parame = g_str_Parame & "   AND BASNEG_SITUAC = 1"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         If (moddat_g_str_TipPar = 1) Then
             MsgBox "El cliente ya ha sido registrado como Negativo.", vbExclamation, modgen_g_str_NomPlt
         ElseIf (moddat_g_str_TipPar = 2) Then
             MsgBox "El cliente ya ha sido registrado como PEP.", vbExclamation, modgen_g_str_NomPlt
         End If
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      'Validando por Fecha de Ingreso a Base Negativa
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_BASNEG "
      g_str_Parame = g_str_Parame & " WHERE BASNEG_TIPBAS = " & moddat_g_str_TipPar
      g_str_Parame = g_str_Parame & "   AND BASNEG_CLITDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
      g_str_Parame = g_str_Parame & "   AND BASNEG_CLINDO = '" & txt_NumDoc.Text & "' "
      g_str_Parame = g_str_Parame & "   AND BASNEG_FECNEG = " & Format(CDate(ipp_FecIng.Text), "yyyymmdd") & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         If (moddat_g_str_TipPar = 1) Then
             MsgBox "El cliente ya ha sido registrado como Negativo con esta fecha.", vbExclamation, modgen_g_str_NomPlt
         ElseIf (moddat_g_str_TipPar = 2) Then
             MsgBox "El cliente ya ha sido registrado como PEP con esta fecha.", vbExclamation, modgen_g_str_NomPlt
         End If
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando en Maestro de Ejecutivos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      g_str_Parame = "USP_CRE_BASNEG_INGRESO ("
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIng.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MotIng.ItemData(cmb_MotIng.ListIndex)) & ", "
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
   Me.Caption = modgen_g_str_NomPlt
   Screen.MousePointer = 0
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipDoc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   If (moddat_g_str_TipPar = 1) Then
       pnl_Titulo1.Caption = "Base Negativa de Clientes"
       pnl_Titulo2.Caption = "Ingreso de Clientes a Base Negativa"
       Call moddat_gs_Carga_LisIte_Combo(cmb_MotIng, 1, "105")
   ElseIf (moddat_g_str_TipPar = 2) Then
       pnl_Titulo1.Caption = "Base PEP de Clientes"
       pnl_Titulo2.Caption = "Ingreso de Clientes a Base PEP"
       Call moddat_gs_Carga_LisIte_Combo(cmb_MotIng, 1, "116")
   End If
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
   cmb_MotIng.ListIndex = -1
   txt_Observ.Text = ""
End Sub

Private Sub cmb_MotIng_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_MotIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MotIng_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub ipp_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MotIng)
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIng)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & ".,- _")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
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

