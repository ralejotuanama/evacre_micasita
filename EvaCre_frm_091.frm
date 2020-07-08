VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10230
   ClientLeft      =   1815
   ClientTop       =   525
   ClientWidth     =   11625
   LinkTopic       =   "Form4"
   ScaleHeight     =   10230
   ScaleWidth      =   11625
   Begin Threed.SSPanel SSPanel1 
      Height          =   10245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18071
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   3465
         Left            =   30
         TabIndex        =   1
         Top             =   4620
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6112
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
         Begin TabDlg.SSTab tab_Seguim 
            Height          =   3345
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   5900
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "EvaCre_frm_091.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label7"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel5"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Descar"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Observ"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "EvaCre_frm_091.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txt_ObsExc"
            Tab(1).Control(1)=   "grd_LisExc"
            Tab(1).Control(2)=   "SSPanel9"
            Tab(1).Control(3)=   "SSPanel11"
            Tab(1).Control(4)=   "SSPanel12"
            Tab(1).Control(5)=   "SSPanel15"
            Tab(1).Control(6)=   "SSPanel16"
            Tab(1).Control(7)=   "pnl_DesExc"
            Tab(1).Control(8)=   "pnl_TipAut"
            Tab(1).Control(9)=   "Label6"
            Tab(1).Control(10)=   "Label3"
            Tab(1).Control(11)=   "Label4"
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "EvaCre_frm_091.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label12"
            Tab(2).Control(1)=   "Label14"
            Tab(2).Control(2)=   "Label15"
            Tab(2).Control(3)=   "pnl_InsCon"
            Tab(2).Control(4)=   "SSPanel20"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel18"
            Tab(2).Control(7)=   "grd_LisCon"
            Tab(2).Control(8)=   "SSPanel17"
            Tab(2).Control(9)=   "txt_LevCon"
            Tab(2).Control(10)=   "txt_ObsCon"
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "EvaCre_frm_091.frx":0054
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "EvaCre_frm_091.frx":0058
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "EvaCre_frm_091.frx":005C
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "EvaCre_frm_091.frx":0060
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "EvaCre_frm_091.frx":0064
               Top             =   2640
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   8
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.21
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   855
               Left            =   30
               TabIndex        =   9
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   60
               TabIndex        =   10
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Ocurrencia"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   2400
               TabIndex        =   11
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Ocurrencia"
               ForeColor       =   16777215
               BackColor       =   16384
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
               Left            =   1230
               TabIndex        =   12
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Ocurrencia"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel pnl_DesOcu 
               Height          =   315
               Left            =   1320
               TabIndex        =   13
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   855
               Left            =   -74970
               TabIndex        =   14
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -74940
               TabIndex        =   15
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -69330
               TabIndex        =   16
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -73770
               TabIndex        =   17
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepción"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -72600
               TabIndex        =   18
               Top             =   360
               Width           =   3285
               _Version        =   65536
               _ExtentX        =   5794
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   45
               Left            =   -74970
               TabIndex        =   19
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73680
               TabIndex        =   20
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_TipAut 
               Height          =   315
               Left            =   -73650
               TabIndex        =   21
               Top             =   2970
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   45
               Left            =   -74970
               TabIndex        =   22
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   -74970
               TabIndex        =   23
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -74940
               TabIndex        =   24
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -65610
               TabIndex        =   25
               Top             =   360
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situación"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -72210
               TabIndex        =   26
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
               ForeColor       =   16777215
               BackColor       =   16384
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
            Begin Threed.SSPanel pnl_InsCon 
               Height          =   315
               Left            =   -73680
               TabIndex        =   27
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   36
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   34
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   33
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   31
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   30
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   29
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   28
               Top             =   2640
               Width           =   1035
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   7530
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   6960
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   38
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito Hipotecario"
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
            Left            =   690
            TabIndex        =   39
            Top             =   330
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Evaluación Crediticia"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6390
            Top             =   150
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "EvaCre_frm_091.frx":0068
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   30
         TabIndex        =   40
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
         Begin VB.CommandButton cmd_ConCre 
            Height          =   585
            Left            =   1920
            Picture         =   "EvaCre_frm_091.frx":0372
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Condiciones Crediticias"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CalIng 
            Height          =   585
            Left            =   1320
            Picture         =   "EvaCre_frm_091.frx":067C
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Calificación de Ingresos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   4320
            Picture         =   "EvaCre_frm_091.frx":0986
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Mantenimiento de Clientes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   5520
            Picture         =   "EvaCre_frm_091.frx":0C90
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   4920
            Picture         =   "EvaCre_frm_091.frx":10D2
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_091.frx":13DC
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPer 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_091.frx":181E
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Verificaciones Personales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   2520
            Picture         =   "EvaCre_frm_091.frx":1B28
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Impresión de Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerLab 
            Height          =   585
            Left            =   630
            Picture         =   "EvaCre_frm_091.frx":1F6A
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Verificaciones Laborales"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   50
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   51
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   52
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         End
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9450
            TabIndex        =   53
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/9999"
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
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   56
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   54
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2055
         Left            =   30
         TabIndex        =   57
         Top             =   8130
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3625
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1965
            Left            =   30
            TabIndex        =   58
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2325
         Left            =   30
         TabIndex        =   59
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4101
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   2205
            Left            =   60
            TabIndex        =   60
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3889
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            Tab             =   6
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "EvaCre_frm_091.frx":2274
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_091.frx":2290
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "EvaCre_frm_091.frx":22AC
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(7)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Patrimonio"
            TabPicture(3)   =   "EvaCre_frm_091.frx":22C8
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(4)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Referencias Personales"
            TabPicture(4)   =   "EvaCre_frm_091.frx":22E4
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Inmueble"
            TabPicture(5)   =   "EvaCre_frm_091.frx":2300
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(2)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos del Crédito"
            TabPicture(6)   =   "EvaCre_frm_091.frx":231C
            Tab(6).ControlEnabled=   -1  'True
            Tab(6).Control(0)=   "Label5"
            Tab(6).Control(0).Enabled=   0   'False
            Tab(6).Control(1)=   "grd_Listad(5)"
            Tab(6).Control(1).Enabled=   0   'False
            Tab(6).Control(2)=   "txt_ObsSol"
            Tab(6).Control(2).Enabled=   0   'False
            Tab(6).ControlCount=   3
            TabCaption(7)   =   "Docum. Recibidos"
            TabPicture(7)   =   "EvaCre_frm_091.frx":2338
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   1290
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Text            =   "EvaCre_frm_091.frx":2354
               Top             =   1470
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   0
               Left            =   -74940
               TabIndex        =   62
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   1
               Left            =   -74940
               TabIndex        =   63
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   6
               Left            =   -74940
               TabIndex        =   64
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1095
               Index           =   5
               Left            =   60
               TabIndex        =   65
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   1931
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   2
               Left            =   -74970
               TabIndex        =   66
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   3
               Left            =   -74940
               TabIndex        =   67
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   4
               Left            =   -74940
               TabIndex        =   68
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   7
               Left            =   -74940
               TabIndex        =   69
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   90
               TabIndex        =   70
               Top             =   1470
               Width           =   1155
            End
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

