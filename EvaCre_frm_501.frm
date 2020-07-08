VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_EvaCre_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   6075
   ClientTop       =   1905
   ClientWidth     =   6315
   Icon            =   "EvaCre_frm_501.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _Version        =   65536
      _ExtentX        =   11192
      _ExtentY        =   5318
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
         TabIndex        =   5
         Top             =   60
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
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
            Height          =   270
            Left            =   630
            TabIndex        =   6
            Top             =   30
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de Cartera de Clientes x Proyecto"
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
         Begin VB.Image Image2 
            Height          =   480
            Left            =   90
            Picture         =   "EvaCre_frm_501.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   780
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
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
            Left            =   5640
            Picture         =   "EvaCre_frm_501.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_501.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1485
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
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
         Begin VB.CheckBox chk_Proyec 
            Caption         =   "Todos los Proyectos"
            Height          =   285
            Left            =   1530
            TabIndex        =   12
            Top             =   450
            Width           =   1845
         End
         Begin VB.ComboBox cmb_CodPry 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   120
            Width           =   4605
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   780
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   1110
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin VB.Label Label4 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   180
            TabIndex        =   13
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Top             =   1170
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   150
            TabIndex        =   9
            Top             =   810
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaCre_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Proyec()      As moddat_tpo_Genera



Private Sub chk_Proyec_Click()
   
   If chk_Proyec.Value = 1 Then
      'cmb_CodPry.ListIndex = -1
      cmb_CodPry.Enabled = False
   Else
      cmb_CodPry.Enabled = True
   End If

End Sub

Private Sub cmd_ExpExc_Click()

   If chk_Proyec.Value = False Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto Inmobiliario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
  Call fs_GenExc
  
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel         As Excel.Application
   Dim r_int_ConVer        As Integer
   Dim r_str_FecIni        As String
   Dim r_str_FecFin        As String
   Dim r_dbl_TipCam        As Double
   
   r_dbl_TipCam = 2.809
   
   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
         
   Screen.MousePointer = 11
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN, CRE_HIPMAE, CRE_SOLMAE "
   g_str_Parame = g_str_Parame & "WHERE DATGEN_TIPDOC = HIPMAE_TDOCLI AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = HIPMAE_NDOCLI AND "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "HIPMAE_PRYINM = '" & l_arr_Proyec(cmb_CodPry.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & r_str_FecIni & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE ASC"


   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
        
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Range(.Cells(1, 24), .Cells(1, 25)).Merge
      .Range(.Cells(2, 24), .Cells(2, 25)).Merge
      .Range(.Cells(1, 24), .Cells(2, 25)).Font.Bold = True
      .Cells(1, 24) = "Dpto. de Tecnología e Informática"
      .Cells(2, 24) = "Desarrollo de Sistemas"

      .Range(.Cells(4, 1), .Cells(4, 11)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(4, 1) = "CARTERA DE CLIENTES"
            
      If chk_Proyec.Value = 0 Then
         .Range(.Cells(6, 1), .Cells(6, 5)).Merge
         .Cells(6, 1).Font.Bold = True
         .Cells(6, 1) = "PROYECTO INMOBILIARIO: " & l_arr_Proyec(cmb_CodPry.ListIndex + 1).Genera_Nombre
      Else
         .Range(.Cells(6, 1), .Cells(6, 3)).Merge
         .Cells(6, 1).Font.Bold = True
         .Cells(6, 1) = "TODOS LOS PROYECTOS"
      End If
   
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "NUMERO OPERACION"
      .Cells(7, 3) = "NUMERO SOLICITUD"
      .Cells(7, 4) = "CONSEJ. HIPOT."
      .Cells(7, 5) = "PRODUCTO"
      .Cells(7, 6) = "SUB PRODUCTO"
      .Cells(7, 7) = "TIPO EVALUACION"
      .Cells(7, 8) = "ACTIV. ECON."
      .Cells(7, 9) = "DOI CLIENTE"
      .Cells(7, 10) = "NOMBRE CLIENTE"
      .Cells(7, 11) = "FEC. DESEMBOLSO"
      .Cells(7, 12) = "PLAZO AÑOS"
      .Cells(7, 13) = "MONEDA"
      .Cells(7, 14) = "COM. VENT. DOLARES"
      .Cells(7, 15) = "APOR. PROP. DOLARES"
      .Cells(7, 16) = "COM. VENT. SOLES"
      .Cells(7, 17) = "APOR. PROP. SOLES"
      .Cells(7, 18) = "MONTO PREST. ORIGINAL"
      .Cells(7, 19) = "MONTO PREST. SOLES"
      .Cells(7, 20) = "TASA INTERES"
      .Cells(7, 21) = "FEC. PREPAGO"
      .Cells(7, 22) = "TIPO GARANTIA"
      .Cells(7, 23) = "MONEDA GARANTIA"
      .Cells(7, 24) = "MONTO GAR. ORIGINAL"
      .Cells(7, 25) = "MONTO GAR. SOLES"

             
      .Range(.Cells(7, 1), .Cells(7, 25)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 25)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 25)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 25)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 25)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 25)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 25)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 4
      '.Columns("B").HorizontalAlignment = xlHAlignLeft
     
      .Columns("B").ColumnWidth = 17
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 16
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 45
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 65
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 20
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 32
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 10
      .Columns("I").HorizontalAlignment = xlHAlignCenter
            
      .Columns("J").ColumnWidth = 45
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Columns("L").ColumnWidth = 11
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Columns("M").ColumnWidth = 18
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Columns("N").ColumnWidth = 17
      .Columns("N").NumberFormat = "###,###,##0.00"
      
      .Columns("O").ColumnWidth = 18
      .Columns("O").NumberFormat = "###,###,##0.00"
            
      .Columns("P").ColumnWidth = 15
      .Columns("P").NumberFormat = "###,###,##0.00"
      
      .Columns("Q").ColumnWidth = 16
      .Columns("Q").NumberFormat = "###,###,##0.00"
      
      .Columns("R").ColumnWidth = 20
      .Columns("R").NumberFormat = "###,###,##0.00"
      
      .Columns("S").ColumnWidth = 18
      .Columns("S").NumberFormat = "###,###,##0.00"
      
      .Columns("T").ColumnWidth = 12
      .Columns("T").NumberFormat = "###,###,##0.00"
      
      .Columns("U").ColumnWidth = 12
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      
      .Columns("V").ColumnWidth = 24
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      
      .Columns("W").ColumnWidth = 18
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      
      .Columns("X").ColumnWidth = 18
      .Columns("X").NumberFormat = "###,###,##0.00"
      
      .Columns("Y").ColumnWidth = 16
      .Columns("Y").NumberFormat = "###,###,##0.00"


      r_int_ConVer = 8
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
      
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 7
         .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(Trim(g_rst_Princi!HIPMAE_NUMOPE))
         .Cells(r_int_ConVer, 3) = gf_Formato_NumSol(Trim(g_rst_Princi!HIPMAE_NUMSOL))
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!HIPMAE_CONHIP)
         .Cells(r_int_ConVer, 5) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         .Cells(r_int_ConVer, 6) = moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
         .Cells(r_int_ConVer, 7) = moddat_gf_Consulta_ParDes("038", g_rst_Princi!SOLMAE_TIPEVA)
         .Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("008", g_rst_Princi!DATGEN_OCUPAC)
         .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         .Cells(r_int_ConVer, 10) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
         .Cells(r_int_ConVer, 11) = gf_FormatoFecha(g_rst_Princi!HIPMAE_FECDES)
         .Cells(r_int_ConVer, 12) = g_rst_Princi!HIPMAE_PLAANO
                  
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = "SOLES"
         Else
            .Cells(r_int_ConVer, 13) = "DOLARES AMERICANOS"
         End If
         
         .Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPMAE_CVTDOL, "###,###,##0.00")
         .Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPMAE_APODOL, "###,###,##0.00")
         .Cells(r_int_ConVer, 16) = Format(g_rst_Princi!HIPMAE_CVTSOL, "###,###,##0.00")
         .Cells(r_int_ConVer, 17) = Format(g_rst_Princi!HIPMAE_APOSOL, "###,###,##0.00")
         .Cells(r_int_ConVer, 18) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
         
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            .Cells(r_int_ConVer, 19) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
         Else
            .Cells(r_int_ConVer, 19) = Format(g_rst_Princi!HIPMAE_MTOPRE * r_dbl_TipCam, "###,###,##0.00")
         End If

         .Cells(r_int_ConVer, 20) = Format(g_rst_Princi!HIPMAE_TASINT, "###,###,##0.00")
         .Cells(r_int_ConVer, 21) = gf_FormatoFecha(g_rst_Princi!HIPMAE_FECPPG)
         .Cells(r_int_ConVer, 22) = moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPMAE_TIPGAR)
         
         If g_rst_Princi!HIPMAE_MONGAR = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = "SOLES"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = "DOLARES AMERICANOS"
         End If
         
         .Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
         
         If g_rst_Princi!HIPMAE_MONGAR = 1 Then
            .Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
         Else
            .Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPMAE_MTOGAR * r_dbl_TipCam, "###,###,##0.00")
         End If
         
         g_rst_Princi.MoveNext
         
         r_int_ConVer = r_int_ConVer + 1
   
      Loop
   
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 25)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 25)).Font.Size = 8
      
      .Range(.Cells(1, 24), .Cells(1, 24)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(2, 24), .Cells(2, 24)).HorizontalAlignment = xlHAlignRight
      
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()

   Call moddat_gs_Carga_Proyec(cmb_CodPry, l_arr_Proyec)
   cmb_CodPry.ListIndex = -1
   chk_Proyec.Value = 0
   ipp_FecFin.Text = date
   ipp_FecIni.Text = date


End Sub



