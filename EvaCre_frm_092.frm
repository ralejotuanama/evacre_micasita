VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_EvaCre_16 
   Caption         =   "Form5"
   ClientHeight    =   2280
   ClientLeft      =   5955
   ClientTop       =   4800
   ClientWidth     =   7815
   LinkTopic       =   "Form5"
   ScaleHeight     =   2280
   ScaleWidth      =   7815
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   4207
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
         TabIndex        =   1
         Top             =   30
         Width           =   7725
         _Version        =   65536
         _ExtentX        =   13626
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
            Height          =   285
            Left            =   660
            TabIndex        =   2
            Top             =   60
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Registradas"
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
            Picture         =   "EvaCre_frm_092.frx":0000
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   7725
         _Version        =   65536
         _ExtentX        =   13626
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_092.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7110
            Picture         =   "EvaCre_frm_092.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   795
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   7725
         _Version        =   65536
         _ExtentX        =   13626
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   60
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            Text            =   "01/01/2008"
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
            TabIndex        =   8
            Top             =   390
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
            Text            =   "01/01/2008"
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
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   150
            TabIndex        =   9
            Top             =   120
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaCre_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ExpExc_Click()
   fs_ExpRep
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub


Private Sub fs_ExpRep()
   Dim r_obj_Excel As Excel.Application
   Dim r_lng_fecini As Long
   Dim r_lng_fecfin As Long
   Dim r_int_IniFil As Integer
   Dim r_int_totcli As Integer
   Dim r_int_totsol As Integer
   
   r_int_IniFil = 8
   r_int_totcli = 0
   r_int_totsol = 0
   
   r_lng_fecini = Val(ipp_FecIni.Year)
   If Val(ipp_FecIni.Month) < 10 Then
      r_lng_fecini = r_lng_fecini & 0 & ipp_FecIni.Month
   Else
      r_lng_fecini = r_lng_fecini & ipp_FecIni.Month
   End If
   If Val(ipp_FecIni.Day) < 10 Then
      r_lng_fecini = r_lng_fecini & 0 & ipp_FecIni.Day
   Else
      r_lng_fecini = r_lng_fecini & ipp_FecIni.Day
   End If
   
   r_lng_fecfin = Val(ipp_FecFin.Year)
   If Val(ipp_FecFin.Month) < 10 Then
      r_lng_fecfin = r_lng_fecfin & 0 & ipp_FecFin.Month
   Else
      r_lng_fecfin = r_lng_fecfin & ipp_FecFin.Month
   End If
   If Val(ipp_FecFin.Day) < 10 Then
      r_lng_fecfin = r_lng_fecfin & 0 & ipp_FecFin.Day
   Else
      r_lng_fecfin = r_lng_fecfin & ipp_FecFin.Day
   End If

   'OBTENER LISTADO DE CONSEJEROS
   g_str_Parame = "SELECT S.SOLMAE_CONHIP ,Y.NROCLI,COUNT(S.SOLMAE_CONHIP) AS NROSOL "
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE S,CLI_DATGEN C, "
   g_str_Parame = g_str_Parame & "( SELECT SOLMAE_CONHIP,COUNT(SOLMAE_CONHIP) as NROCLI "
   g_str_Parame = g_str_Parame & "FROM ( "
   g_str_Parame = g_str_Parame & "SELECT DATGEN_NUMDOC,SOLMAE_CONHIP "
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & "LEFT JOIN CLI_DATGEN ON (SOLMAE_TITTDO=DATGEN_TIPDOC "
   g_str_Parame = g_str_Parame & "AND SOLMAE_TITNDO=DATGEN_NUMDOC) "
   g_str_Parame = g_str_Parame & "WHERE SOLMAE_FECSOL >=" & r_lng_fecini & " "
   g_str_Parame = g_str_Parame & "AND SOLMAE_FECSOL <=" & r_lng_fecfin & " "
   g_str_Parame = g_str_Parame & "GROUP BY DATGEN_NUMDOC,SOLMAE_CONHIP "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP) "
   g_str_Parame = g_str_Parame & "GROUP BY SOLMAE_CONHIP ) Y "
   g_str_Parame = g_str_Parame & "WHERE "
   g_str_Parame = g_str_Parame & "S.SOLMAE_TITTDO=C.DATGEN_TIPDOC  "
   g_str_Parame = g_str_Parame & "AND S.SOLMAE_TITNDO=C.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "AND S.SOLMAE_CONHIP =Y.SOLMAE_CONHIP "
   g_str_Parame = g_str_Parame & "AND S.SOLMAE_FECSOL >=" & r_lng_fecini & " "
   g_str_Parame = g_str_Parame & "AND S.SOLMAE_FECSOL <=" & r_lng_fecfin & " "
   g_str_Parame = g_str_Parame & "GROUP BY S.SOLMAE_CONHIP, Y.NROCLI "
   g_str_Parame = g_str_Parame & "ORDER BY S.SOLMAE_CONHIP"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados de los consejeros.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'OBTENER NRO DE SOLICITUDES
   g_str_Parame = "SELECT DATGEN_NUMDOC,SOLMAE_CONHIP "
   g_str_Parame = g_str_Parame & ",TRIM(DATGEN_APEPAT)||' '|| TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS NOMBRE "
   g_str_Parame = g_str_Parame & ",COUNT (*) AS NROVCS "
   g_str_Parame = g_str_Parame & "FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & "LEFT JOIN CLI_DATGEN ON (SOLMAE_TITTDO=DATGEN_TIPDOC AND SOLMAE_TITNDO=DATGEN_NUMDOC) "
   g_str_Parame = g_str_Parame & "WHERE SOLMAE_FECSOL >=" & r_lng_fecini & " "
   g_str_Parame = g_str_Parame & "AND SOLMAE_FECSOL <=" & r_lng_fecfin & " "
   g_str_Parame = g_str_Parame & "GROUP BY DATGEN_NUMDOC,SOLMAE_CONHIP, TRIM(DATGEN_APEPAT)"
   g_str_Parame = g_str_Parame & "||' '|| TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
                       
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
     
   With r_obj_Excel.ActiveSheet
      'Unir celdas
      .Range("B5") = "REPORTE DE SOLICITUDES REGISTRADAS"
      .Range("B5:E5").Font.Underline = True
      .Range("A5:C5").Merge
      .Range("A5:C5").HorizontalAlignment = xlHAlignCenter
      .Range("B5:E5").Font.Bold = True
      .Range("A1:A6").Font.Bold = True
         
      .Range("A1:J600").Font.Name = "Arial"
      .Range("A1:J600").Font.Size = 9
      
      .Columns("A").ColumnWidth = 14.57
      .Columns("B").ColumnWidth = 46
      .Columns("C").ColumnWidth = 11.29
      .Columns("D").ColumnWidth = 11.57
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
   
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
   
      .Cells(r_int_IniFil, 1) = "USUARIO"
      .Cells(r_int_IniFil, 1).Font.Bold = True
      r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      
      r_int_IniFil = r_int_IniFil + 2
      
      .Cells(r_int_IniFil, 1) = "NRO DOI"
      .Cells(r_int_IniFil, 2) = "NOMBRE Y APELLIDOS"
      .Cells(r_int_IniFil, 3) = "INGRESADOS"
      .Cells(r_int_IniFil, 1).Font.Bold = True
      .Cells(r_int_IniFil, 2).Font.Bold = True
      .Cells(r_int_IniFil, 3).Font.Bold = True
      .Cells(r_int_IniFil, 2).HorizontalAlignment = xlHAlignCenter
      r_int_IniFil = r_int_IniFil + 1
      
         g_rst_Listas.MoveFirst
         Do While Not g_rst_Listas.EOF
            If Trim(g_rst_Princi!SOLMAE_CONHIP) = Trim(g_rst_Listas!SOLMAE_CONHIP) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 1) = g_rst_Listas!DATGEN_NUMDOC
               r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 2) = Trim(g_rst_Listas!NOMBRE)
               r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 3) = Trim(g_rst_Listas!NROVCS)
               r_int_IniFil = r_int_IniFil + 1
            End If
            
            g_rst_Listas.MoveNext
         Loop
         r_int_IniFil = r_int_IniFil + 2
         g_rst_Princi.MoveNext
      Loop
   
      r_int_IniFil = r_int_IniFil + 2
      
      
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
   
      .Cells(r_int_IniFil, 3) = "CLIENTES"
      .Cells(r_int_IniFil, 4) = "SOLICITUDES"
      .Cells(r_int_IniFil, 3).Font.Bold = True
      .Cells(r_int_IniFil, 4).Font.Bold = True
      r_int_IniFil = r_int_IniFil + 1
   
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         .Cells(r_int_IniFil, 2).HorizontalAlignment = xlHAlignRight
         r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
         r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 3) = g_rst_Princi!NROCLI
         r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 4) = g_rst_Princi!NROSOL
         r_int_totcli = r_int_totcli + g_rst_Princi!NROCLI
         r_int_totsol = r_int_totsol + g_rst_Princi!NROSOL
         r_int_IniFil = r_int_IniFil + 1
      g_rst_Princi.MoveNext
      Loop
    
   .Cells(r_int_IniFil, 2) = "TOTAL"
   .Cells(r_int_IniFil, 2).HorizontalAlignment = xlHAlignRight
   r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 3) = r_int_totcli
   r_obj_Excel.ActiveSheet.Cells(r_int_IniFil, 4) = r_int_totsol
   .Cells(r_int_IniFil, 2).Font.Bold = True
   .Cells(r_int_IniFil, 3).Font.Bold = True
   .Cells(r_int_IniFil, 4).Font.Bold = True
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
          
   Screen.MousePointer = 0
      
   r_obj_Excel.Visible = True
      
   Set r_obj_Excel = Nothing
   
End Sub
 
Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_FecIni.Text = date
   ipp_FecFin.Text = date
End Sub

