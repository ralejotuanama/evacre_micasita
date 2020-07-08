VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_MntCli_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   4995
   ClientLeft      =   2880
   ClientTop       =   2235
   ClientWidth     =   7305
   Icon            =   "EvaCre_frm_016.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4995
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7305
      _Version        =   65536
      _ExtentX        =   12885
      _ExtentY        =   8811
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
         Height          =   795
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   5790
            Picture         =   "EvaCre_frm_016.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   420
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5100
            Picture         =   "EvaCre_frm_016.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   6480
            Picture         =   "EvaCre_frm_016.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   420
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3345
         Left            =   30
         TabIndex        =   15
         Top             =   1590
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   5900
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
         Begin VB.CommandButton cmd_LimBus 
            Height          =   675
            Left            =   6480
            Picture         =   "EvaCre_frm_016.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   675
            Left            =   5790
            Picture         =   "EvaCre_frm_016.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.TextBox txt_LisCli 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "EvaCre_frm_016.frx":1076
            Top             =   1200
            Width           =   7065
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   60
            Width           =   2775
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   90
            Left            =   30
            TabIndex        =   16
            Top             =   1080
            Width           =   7125
            _Version        =   65536
            _ExtentX        =   12568
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres"
            Height          =   285
            Left            =   30
            TabIndex        =   19
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno"
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
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
            TabIndex        =   21
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
            Picture         =   "EvaCre_frm_016.frx":107C
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgBus  As Integer

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = " & "'" & txt_NumDoc.Text & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If MsgBox("No se encontró registrado al Cliente. ¿Desea crear al Cliente?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   'Doc. Identidad
   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_NumDoc = txt_NumDoc.Text
   
   frm_MntCli_02.Show 1
End Sub

Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   txt_LisCli.Text = ""
   
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = "SELECT * FROM CLI_BUSCLI WHERE "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEPAT) LIKE '" & r_str_ApePat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEMAT) LIKE '" & r_str_ApeMat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_NOMBRE) LIKE '" & r_str_Nombre & "' ORDER BY "
   g_str_Parame = g_str_Parame & "BUSCLI_APEPAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_APEMAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_NOMBRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   l_int_FlgBus = 2
   Do While Not g_rst_Princi.EOF
      txt_LisCli.Text = txt_LisCli.Text & CStr(g_rst_Princi!BUSCLI_TIPDOC) & "-"
      txt_LisCli.Text = txt_LisCli.Text & Mid(Trim(g_rst_Princi!BUSCLI_NUMDOC) & Space(12), 1, 12) & Space(3)
      txt_LisCli.Text = txt_LisCli.Text & Trim(g_rst_Princi!BUSCLI_APEPAT) & " " & Trim(g_rst_Princi!BUSCLI_APEMAT) & " " & Trim(g_rst_Princi!BUSCLI_NOMBRE)
      txt_LisCli.Text = txt_LisCli.Text & Chr(13) & Chr(10)
      
      g_rst_Princi.MoveNext
   Loop
   l_int_FlgBus = 1
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_LimBus_Click()
   Call fs_Limpia_BusCli
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   l_int_FlgBus = 1
   
   Screen.MousePointer = 0
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

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_LisCli_KeyPress(KeyAscii As Integer)
   If l_int_FlgBus = 1 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call fs_Limpia_BusCli
End Sub

Private Sub fs_Limpia_BusCli()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   txt_LisCli.Text = ""
End Sub

