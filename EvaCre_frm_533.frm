VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_EvaCre_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   6405
   ClientTop       =   4455
   ClientWidth     =   5355
   Icon            =   "EvaCre_frm_533.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      _Version        =   65536
      _ExtentX        =   9446
      _ExtentY        =   4260
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
         Top             =   60
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            TabIndex        =   2
            Top             =   30
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de Créditos Atrasados"
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
            Picture         =   "EvaCre_frm_533.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   780
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            Picture         =   "EvaCre_frm_533.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4650
            Picture         =   "EvaCre_frm_533.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   30
         TabIndex        =   6
         Top             =   1470
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   8
            Top             =   420
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaCre_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_PerMes           As String
Dim l_str_PerAno           As String

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_SegExc     As String
Dim r_str_NumSol     As String
Dim r_int_Contad     As Integer
           
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT HIPCIE_NUMOPE, PRODUC_DESCRI, HIPCIE_TDOCLI, HIPCIE_NDOCLI, DATGEN_APEPAT, "
    g_str_Parame = g_str_Parame & "       DATGEN_APEMAT, DATGEN_NOMBRE, HIPCIE_FECDES, HIPCIE_TIPEVA, HIPCIE_TIPMON, "
    g_str_Parame = g_str_Parame & "       HIPCIE_MTOPRE, HIPCIE_SALCAP, HIPCIE_SALCON, HIPCIE_MTOPRE, HIPCIE_SALCAP, "
    g_str_Parame = g_str_Parame & "       HIPCIE_SALCON, HIPCIE_TIPGAR, HIPCIE_MONGAR, HIPCIE_MTOGAR, HIPCIE_DIAMOR, "
    g_str_Parame = g_str_Parame & "       (SELECT X.PARDES_DESCRI FROM MNT_PARDES X WHERE PARDES_CODGRP = '008' "
    g_str_Parame = g_str_Parame & "           AND PARDES_CODITE = (SELECT ACTECO_CODACT FROM CLI_ACTECO "
    g_str_Parame = g_str_Parame & "                                 Where ACTECO_CLITDO = DATGEN_TIPDOC "
    g_str_Parame = g_str_Parame & "                                   AND ACTECO_CLINDO = DATGEN_NUMDOC AND ACTECO_ORDACT = 1)) AS OCUPACION "
    g_str_Parame = g_str_Parame & " FROM CRE_HIPCIE A, CLI_DATGEN B, CRE_PRODUC C "
    g_str_Parame = g_str_Parame & "WHERE HIPCIE_TDOCLI = DATGEN_TIPDOC "
    g_str_Parame = g_str_Parame & "   AND HIPCIE_NDOCLI = DATGEN_NUMDOC "
    g_str_Parame = g_str_Parame & "   AND HIPCIE_CODPRD = PRODUC_CODIGO "
    g_str_Parame = g_str_Parame & "   AND HIPCIE_DIAMOR > 30 "
    g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & l_str_PerMes
    g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & l_str_PerAno
    g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE ASC "
       
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE RCC"
   
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 18), .Cells(2, 18)).Font.Bold = True
      .Cells(1, 18) = "Dpto. de Tecnología e Informática"
      .Cells(2, 18) = "Desarrollo de Sistemas"
           
      .Range(.Cells(4, 1), .Cells(4, 13)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(4, 1) = "REPORTE DE CREDITOS ATRASADOS " & l_str_PerAno & "-" & Format(l_str_PerMes, "00")
   
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "NUMERO OPERACION"
      .Cells(7, 3) = "PRODUCTO"
      .Cells(7, 4) = "DOI CLIENTE"
      .Cells(7, 5) = "NOMBRE CLIENTE"
      .Cells(7, 6) = "FECHA DESEMBOLSO"
      .Cells(7, 7) = "TIPO EVALUACION"
      .Cells(7, 8) = "OCUPACION"
      .Cells(7, 9) = "MTO. PRESTAMO ($/.)"
      .Cells(7, 10) = "MTO. PRESTAMO (US$/.)"
      .Cells(7, 11) = "SALDO TOTAL ($/.)"
      .Cells(7, 12) = "SALDO TOTAL (US$/.)"
      .Cells(7, 13) = "TIPO GARANTIA"
      .Cells(7, 14) = "MTO. GARANTIA ($/.)"
      .Cells(7, 15) = "MTO. GARANTIA (US$/.)"
      .Cells(7, 16) = "DIAS ATRASO"
      .Cells(7, 17) = "EXCEPCION"
      .Cells(7, 18) = "COMENTARIOS"
       
      .Range(.Cells(7, 1), .Cells(7, 18)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 18)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 18)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 18)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 18)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 18)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 40
      '.Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 17
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 38
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 17
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 19
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 17
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 25
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 17
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("O").ColumnWidth = 19
      .Columns("O").NumberFormat = "###,###,##0.00"
      .Columns("P").ColumnWidth = 11
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 9
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 100
      '.Columns("R").HorizontalAlignment = xlHAlignCenter
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 8
   
      Do While Not g_rst_Princi.EOF
         'Buscando datos de la Garantía en Registro de Hipotecas
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 7
         .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(Trim(g_rst_Princi!HIPCIE_NUMOPE))
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PRODUC_DESCRI)
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!HIPCIE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPCIE_NDOCLI)
         .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
         .Cells(r_int_ConVer, 6) = gf_FormatoFecha(Trim(g_rst_Princi!HIPCIE_FECDES))
         .Cells(r_int_ConVer, 7) = moddat_gf_Consulta_ParDes("038", Trim(g_rst_Princi!HIPCIE_TIPEVA))
         .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!OCUPACION) 'Trim(g_rst_Princi!PARDES_DESCRI)
         
         If Trim(g_rst_Princi!HIPCIE_TIPMON) = 1 Then
            .Cells(r_int_ConVer, 9) = Format(Trim(g_rst_Princi!HIPCIE_MTOPRE), "###,###,##0.00")
            .Cells(r_int_ConVer, 11) = Format(Trim(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON), "###,###,##0.00")
         ElseIf Trim(g_rst_Princi!HIPCIE_TIPMON) = 2 Then
            .Cells(r_int_ConVer, 10) = Format(Trim(g_rst_Princi!HIPCIE_MTOPRE), "###,###,##0.00")
            .Cells(r_int_ConVer, 12) = Format(Trim(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON), "###,###,##0.00")
         End If
         
         .Cells(r_int_ConVer, 13) = moddat_gf_Consulta_ParDes("241", Trim(g_rst_Princi!HIPCIE_TIPGAR))
         
         If Trim(g_rst_Princi!HIPCIE_MONGAR) = 1 Then
            .Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
         ElseIf Trim(g_rst_Princi!HIPCIE_MONGAR) = 2 Then
            .Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
         End If
         
         .Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!HIPCIE_DIAMOR)
         
         r_str_NumSol = modsec_gf_Buscar_NumSol(Trim(g_rst_Princi!HIPCIE_NUMOPE))
         r_str_SegExc = gf_BusExc(r_str_NumSol)
         
         If r_str_SegExc <> "" Then
            .Cells(r_int_ConVer, 17) = "SI"
            .Cells(r_int_ConVer, 18) = Trim(r_str_SegExc)
            
            For r_int_Contad = 1 To Len(r_str_SegExc) Step 1
               If Len(r_str_SegExc) <= (150 * r_int_Contad) Then
                  .Range("R" & r_int_ConVer).RowHeight = 15 * r_int_Contad
                  .Range("R" & r_int_ConVer).WrapText = True
                  .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 18)).VerticalAlignment = xlVAlignCenter
                  Exit For
               End If
            Next
         End If
         
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 18)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 18)).Font.Size = 8
      
      .Range(.Cells(1, 18), .Cells(1, 18)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(2, 18), .Cells(2, 18)).HorizontalAlignment = xlHAlignRight
   End With
     
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function gf_BusExc(ByVal P_NumSol As String) As String
   gf_BusExc = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGEXC "
   g_str_Parame = g_str_Parame & "WHERE SEGEXC_CODINS = 21 "
   g_str_Parame = g_str_Parame & "AND SEGEXC_NUMSOL = " & P_NumSol & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGEXC_NUMEXC DESC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      gf_BusExc = Trim(g_rst_Listas!SEGEXC_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
Dim r_int_PerMes  As Integer
Dim r_int_PerAno  As Integer

   r_int_PerMes = Month(date)
   r_int_PerAno = Year(date)
   
   If Month(date) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If
 
   Call gs_BuscarCombo_Item(cmb_PerMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
End Sub
