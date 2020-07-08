VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_RecSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   1380
   ClientTop       =   1950
   ClientWidth     =   14880
   Icon            =   "EvaCre_frm_030.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14865
      _Version        =   65536
      _ExtentX        =   26220
      _ExtentY        =   15637
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
         Height          =   7245
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
         _ExtentY        =   12779
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   3750
            TabIndex        =   5
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   5250
            TabIndex        =   6
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   6750
            TabIndex        =   7
            Top             =   60
            Width           =   5385
            _Version        =   65536
            _ExtentX        =   9499
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   12120
            TabIndex        =   8
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Registro"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   13620
            TabIndex        =   9
            Top             =   60
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Selec."
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   14685
            _ExtentX        =   25903
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   14775
         _Version        =   65536
         _ExtentX        =   26061
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
            TabIndex        =   2
            Top             =   60
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Recepción de Solicitudes"
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
            Picture         =   "EvaCre_frm_030.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   60
         TabIndex        =   11
         Top             =   8040
         Width           =   14715
         _Version        =   65536
         _ExtentX        =   25956
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
         Begin VB.CommandButton cmd_DatSol 
            Height          =   675
            Left            =   30
            Picture         =   "EvaCre_frm_030.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Datos de la Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RecSol 
            Height          =   675
            Left            =   13290
            Picture         =   "EvaCre_frm_030.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Enviar Solicitudes"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   14010
            Picture         =   "EvaCre_frm_030.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_RecSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_DatSol_Click()
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 1
   moddat_g_str_CodPrd = Mid(grd_Listad.Text, 1, 3)
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 7
   moddat_g_str_CodSub = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_CSgSol_02.Show 1
End Sub

Private Sub cmd_RecSol_Click()
   Dim r_str_Cadena  As String
   Dim r_str_NumSoA  As String
   Dim r_str_IdeCli  As String
   Dim r_str_NomCli  As String

   Dim r_int_Contad  As Integer
   Dim r_str_NumSol  As String
   Dim r_int_DiaTra  As String
   Dim r_str_FecSol  As String
   Dim r_int_FlgSel  As Integer
   
   r_str_Cadena = ""
   
   grd_Listad.Redraw = False
   
   r_int_FlgSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 5
      
      If grd_Listad.Text = "X" Then
         r_int_FlgSel = 1
         Exit For
      End If
   Next r_int_Contad

   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)

   If r_int_FlgSel = 0 Then
      MsgBox "Debe seleccionar las Solicitudes a enviar a Evaluación Crediticia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de recepcionar las Solicitudes?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 5
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Col = 1
         r_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
         r_str_NumSoA = grd_Listad.Text
      
         grd_Listad.Col = 2
         r_str_IdeCli = grd_Listad.Text
      
         grd_Listad.Col = 3
         r_str_NomCli = grd_Listad.Text
      
         grd_Listad.Col = 4
         r_str_FecSol = grd_Listad.Text
         
         Call moddat_gs_FecSis
         r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(r_str_FecSol))
      
         'Creando Nueva Ocurrencia en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(r_str_NumSol, 11, 15, 0, "", 0, 0) Then
            Exit Sub
         End If
         
         'Modificando Registro en Instancia Actual
         If Not moddat_gf_Modifica_Seguim(r_str_NumSol, 11, r_int_DiaTra, 1, 7) Then
            Exit Sub
         End If
         
                  
         'Creando Registro de Nueva Instancia
         If Not moddat_gf_Inserta_Seguim(r_str_NumSol, 21) Then
            Exit Sub
         End If
         
         'Creando Nueva Ocurrencia en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(r_str_NumSol, 21, 11, 0, "", 0, 0) Then
            Exit Sub
         End If
         
         'Actualizando en Maestro de Solicitudes
         If Not modatecli_gf_ActIns_SolMae(r_str_NumSol, 21) Then
            Exit Sub
         End If
         
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
   
   MsgBox "Se recepcionaron las solicitudes con éxito.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3675
   grd_Listad.ColWidth(1) = 1515
   grd_Listad.ColWidth(2) = 1505
   grd_Listad.ColWidth(3) = 5365
   grd_Listad.ColWidth(4) = 1505
   grd_Listad.ColWidth(5) = 825
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
End Sub


Private Sub fs_Buscar()
   'Obtener Tasa de Interes de Producto
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 11 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         
         'Obteniendo Información del Cliente
         grd_Listad.Col = 3
         grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
         
         grd_Listad.Col = 4
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
         
         grd_Listad.Col = 5
         grd_Listad.Text = ""
         
         grd_Listad.Col = 6
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
         
         grd_Listad.Col = 8
         grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
         
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
   Else
      cmd_DatSol.Enabled = False
      cmd_RecSol.Enabled = False
      MsgBox "No se encontraron Solicitudes registradas pendientes de Envio a Evaluación Crediticia.", vbInformation, modgen_g_str_NomPlt
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 5
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

