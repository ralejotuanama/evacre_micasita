VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_EvaCre_67 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   3135
   ClientLeft      =   6090
   ClientTop       =   3885
   ClientWidth     =   7995
   Icon            =   "EvaCre_frm_093.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   5715
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
         Left            =   60
         TabIndex        =   1
         Top             =   2310
         Width           =   7890
         _Version        =   65536
         _ExtentX        =   13917
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
         Begin VB.TextBox txt_NroAct 
            Height          =   315
            Left            =   2280
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Text            =   "Text"
            Top             =   390
            Width           =   3495
         End
         Begin EditLib.fpDateTime ipp_FecCom 
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Top             =   60
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
         Begin VB.Label Label12 
            Caption         =   "Fecha Comité Créditos:"
            Height          =   285
            Left            =   90
            TabIndex        =   5
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Número Acta Comité:"
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   435
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   7890
         _Version        =   65536
         _ExtentX        =   13917
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
            Left            =   690
            TabIndex        =   7
            Top             =   30
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   690
            TabIndex        =   8
            Top             =   330
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia - Registro de Comité de Créditos"
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
            Left            =   60
            Picture         =   "EvaCre_frm_093.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   825
         Left            =   60
         TabIndex        =   9
         Top             =   1440
         Width           =   7890
         _Version        =   65536
         _ExtentX        =   13917
         _ExtentY        =   1455
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
            Left            =   1140
            TabIndex        =   10
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   5760
            TabIndex        =   11
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "29/08/2012"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Left            =   1140
            TabIndex        =   12
            Top             =   405
            Width           =   6645
            _Version        =   65536
            _ExtentX        =   11721
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Height          =   300
            Left            =   90
            TabIndex        =   15
            Top             =   465
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   4200
            TabIndex        =   14
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   120
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   16
         Top             =   765
         Width           =   7890
         _Version        =   65536
         _ExtentX        =   13917
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
            Left            =   7270
            Picture         =   "EvaCre_frm_093.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_093.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   If CDate(ipp_FecCom.Text) > date Then
      MsgBox "La fecha de Aprobación del Comité no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecCom)
      Exit Sub
   End If
   
   If CDate(ipp_FecCom.Text) > date Then
      MsgBox "La fecha de Aprobación del Comité no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecCom)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
        
      g_str_Parame = "usp_tra_evacre_comite ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecCom.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NroAct.Text & "', "
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
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Screen.MousePointer = 0
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop

    'Creando Nueva Ocurrencia en Detalle de Seguimiento
    If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 32, 0, "", 0, 0) Then
        Exit Sub
    End If
    
    MsgBox "Registro de comite de créditos correctamente.", vbInformation, modgen_g_str_NomPlt
    
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
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
      
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_FecCom.Text = Format(date, "dd/mm/yyyy")
   txt_NroAct.Text = ""
   pnl_NumSol.Caption = ""
   pnl_Client.Caption = ""
   pnl_FecIng.Caption = ""
   
End Sub

Private Sub ipp_FecCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NroAct)
   End If
End Sub
 
Private Sub txt_NroAct_GotFocus()
   Call gs_SelecTodo(txt_NroAct)
End Sub

Private Sub txt_NroAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ()/" & Chr(10))
   End If
End Sub

Private Sub fs_Buscar_DatEva()
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT EVACRE_NUMSOL, TO_DATE(EVACRE_FECCOM,'YYYY/MM/DD') AS EVACRE_FECCOM, EVACRE_NROACT "
    g_str_Parame = g_str_Parame & "FROM TRA_EVACRE WHERE  "
    g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        Exit Sub
    End If
   
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
        g_rst_Princi.MoveFirst
       
        ipp_FecCom.Text = IIf(IsNull(g_rst_Princi!EVACRE_FECCOM), Format(date, "dd/mm/yyyy"), (g_rst_Princi!EVACRE_FECCOM))
        txt_NroAct.Text = IIf(IsNull(g_rst_Princi!EVACRE_NROACT), "", (Trim(g_rst_Princi!EVACRE_NROACT & "")))
 
    End If
   
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
End Sub

