VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_MntEmp_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   1365
   ClientTop       =   2415
   ClientWidth     =   12525
   Icon            =   "EvaCre_frm_035.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12525
      _Version        =   65536
      _ExtentX        =   22093
      _ExtentY        =   10980
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
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
            TabIndex        =   13
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Empresas"
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
            Picture         =   "EvaCre_frm_035.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4605
         Left            =   30
         TabIndex        =   14
         Top             =   1560
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
         _ExtentY        =   8123
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
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_LimBus 
            Height          =   675
            Left            =   11700
            Picture         =   "EvaCre_frm_035.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   675
            Left            =   11010
            Picture         =   "EvaCre_frm_035.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar Clientes"
            Top             =   30
            Width           =   675
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin Threed.SSPanel SSPanel24 
            Height          =   90
            Left            =   30
            TabIndex        =   15
            Top             =   1080
            Width           =   12375
            _Version        =   65536
            _ExtentX        =   21828
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCli 
            Height          =   3375
            Left            =   60
            TabIndex        =   10
            Top             =   1200
            Width           =   12345
            _ExtentX        =   21775
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Left            =   30
            TabIndex        =   18
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Razón Social:"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   19
         Top             =   750
         Width           =   12435
         _Version        =   65536
         _ExtentX        =   21934
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11040
            Picture         =   "EvaCre_frm_035.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10350
            Picture         =   "EvaCre_frm_035.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   11730
            Picture         =   "EvaCre_frm_035.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   390
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_MntEmp_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         txt_RazSoc.Enabled = True
         
         txt_NomCom.Enabled = False
         txt_NomCom.Text = ""
         
         Call gs_SetFocus(txt_RazSoc)
      Else
         txt_RazSoc.Enabled = False
         txt_RazSoc.Text = ""
         
         txt_NomCom.Enabled = True
         Call gs_SetFocus(txt_NomCom)
      End If
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

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

   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_NumDoc = txt_NumDoc.Text

   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      moddat_g_int_FlgGrb = 1
   Else
      moddat_g_int_FlgGrb = 2
   End If
   
   frm_MntEmp_02.Show 1
End Sub

Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String
   Dim r_str_CadBus  As String

   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If Len(Trim(txt_RazSoc.Text)) = 0 Then
         MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_RazSoc)
         Exit Sub
      End If
   Else
      If Len(Trim(txt_NomCom.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomCom)
         Exit Sub
      End If
   End If
   
   Call gs_LimpiaGrid(grd_LisCli)
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      r_str_CadBus = "%" & txt_RazSoc.Text & "%"
   Else
      r_str_CadBus = "%" & txt_NomCom.Text & "%"
   End If
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & "RTRIM(DATGEN_RAZSOC) LIKE '" & r_str_CadBus & "' "
      g_str_Parame = g_str_Parame & "ORDER BY DATGEN_RAZSOC ASC"
   Else
      g_str_Parame = g_str_Parame & "RTRIM(DATGEN_NOMCOM) LIKE '" & r_str_CadBus & "' "
      g_str_Parame = g_str_Parame & "ORDER BY DATGEN_NOMCOM ASC"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado empresas para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisCli.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_LisCli.Rows = grd_LisCli.Rows + 1
      grd_LisCli.Row = grd_LisCli.Rows - 1
      
      grd_LisCli.Col = 0
      grd_LisCli.Text = CStr(g_rst_Princi!DATGEN_EMPTDO) & "-" & Trim(g_rst_Princi!DATGEN_EMPNDO & "")
      
      grd_LisCli.Col = 1
      
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         grd_LisCli.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")
      Else
         grd_LisCli.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCli.Redraw = True
   Call gs_UbiIniGrid(grd_LisCli)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_LimBus_Click()
   Call fs_Limpia_BusCli
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call cmd_Limpia_Click
   Call fs_Limpia_BusCli
      
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_LisCli.ColWidth(0) = 2000
   grd_LisCli.ColWidth(1) = 10000
   
   grd_LisCli.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCli.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")
   
   cmb_TipBus.Clear
   cmb_TipBus.AddItem "POR RAZON SOCIAL"
   cmb_TipBus.ItemData(cmb_TipBus.NewIndex) = 1
   
   cmb_TipBus.AddItem "POR NOMBRE COMERCIAL"
   cmb_TipBus.ItemData(cmb_TipBus.NewIndex) = 2
   
End Sub

Private Sub fs_Limpia_BusCli()
   cmb_TipBus.ListIndex = -1
   
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   
   txt_RazSoc.Enabled = False
   txt_NomCom.Enabled = False
   
   Call gs_LimpiaGrid(grd_LisCli)
End Sub

Private Sub grd_LisCli_DblClick()
   Dim r_int_TipDoc     As Integer
   Dim r_str_NumDoc     As String

   If grd_LisCli.Rows > 0 Then
      grd_LisCli.Col = 0
      
      r_int_TipDoc = CInt(Left(grd_LisCli.Text, 1))
      r_str_NumDoc = Mid(grd_LisCli.Text, 3)
   
      Call gs_RefrescaGrid(grd_LisCli)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_int_TipDoc)
      txt_NumDoc.Text = r_str_NumDoc
      
      Call cmd_Buscar_Click
      Call gs_SetFocus(grd_LisCli)
   End If
End Sub

Private Sub grd_LisCli_SelChange()
   If grd_LisCli.Rows > 2 Then
      grd_LisCli.RowSel = grd_LisCli.Row
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- _.,;:)(/&%$@#*+")
   End If
End Sub

Private Sub txt_NomCom_GotFocus()
   Call gs_SelecTodo(txt_NomCom)
End Sub

Private Sub txt_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- _.,;:)(/&%$@#*+")
   End If
End Sub


