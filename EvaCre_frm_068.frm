VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_EvaCre_62 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10230
   ClientLeft      =   2610
   ClientTop       =   2130
   ClientWidth     =   11610
   Icon            =   "EvaCre_frm_068.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
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
            TabPicture(0)   =   "EvaCre_frm_068.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
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
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "EvaCre_frm_068.frx":0028
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
            Tab(1).Control(9)=   "pnl_Motivo"
            Tab(1).Control(10)=   "lbl_motivo"
            Tab(1).Control(11)=   "Label4"
            Tab(1).Control(12)=   "Label3"
            Tab(1).Control(13)=   "Label6"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "EvaCre_frm_068.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_LevCon"
            Tab(2).Control(1)=   "txt_ObsCon"
            Tab(2).Control(2)=   "SSPanel17"
            Tab(2).Control(3)=   "grd_LisCon"
            Tab(2).Control(4)=   "SSPanel18"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel20"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "Label12"
            Tab(2).Control(9)=   "Label14"
            Tab(2).Control(10)=   "Label15"
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "EvaCre_frm_068.frx":0060
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "EvaCre_frm_068.frx":0064
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73770
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "EvaCre_frm_068.frx":0068
               Top             =   1980
               Width           =   10065
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "EvaCre_frm_068.frx":006C
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "EvaCre_frm_068.frx":0070
               Top             =   1980
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
                  Size            =   8.25
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
               Cols            =   6
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
               Left            =   -73770
               TabIndex        =   20
               Top             =   1650
               Width           =   10050
               _Version        =   65536
               _ExtentX        =   17727
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
               Left            =   -73770
               TabIndex        =   21
               Top             =   2970
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_Motivo 
               Height          =   315
               Left            =   -68880
               TabIndex        =   66
               Top             =   2970
               Width           =   5205
               _Version        =   65536
               _ExtentX        =   9181
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69630
               TabIndex        =   65
               Top             =   3030
               Width           =   705
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   36
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   34
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   33
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   255
               Left            =   -74940
               TabIndex        =   31
               Top             =   3030
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   30
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   29
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   28
               Top             =   1980
               Width           =   1155
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
            Left            =   4770
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
            Left            =   4110
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Threed.SSPanel pnl_AprCon 
            Height          =   555
            Left            =   8460
            TabIndex        =   38
            Top             =   60
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CLIENTE CON APROBACION CONDICIONADA PENDIENTE"
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   690
            TabIndex        =   50
            Top             =   30
            Width           =   2865
            _Version        =   65536
            _ExtentX        =   5054
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
            TabIndex        =   51
            Top             =   330
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
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
            Left            =   3570
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
         Begin Threed.SSPanel pnl_Reingr 
            Height          =   555
            Left            =   5400
            TabIndex        =   64
            Top             =   60
            Visible         =   0   'False
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   979
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Outline         =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "EvaCre_frm_068.frx":0074
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   30
         TabIndex        =   39
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
         Begin VB.CommandButton cmd_Aprueb_ValJef 
            Height          =   585
            Left            =   9180
            Picture         =   "EvaCre_frm_068.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Aprobar Solicitud por el Jefe de Créditos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   585
            Left            =   8580
            Picture         =   "EvaCre_frm_068.frx":0688
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   8010
            Picture         =   "EvaCre_frm_068.frx":0ACA
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_AprCon 
            Height          =   585
            Left            =   7440
            Picture         =   "EvaCre_frm_068.frx":0DD4
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Aprobación con Condición"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SolRec 
            Height          =   585
            Left            =   6870
            Picture         =   "EvaCre_frm_068.frx":10DE
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Solicitudes Rechazadas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   6300
            Picture         =   "EvaCre_frm_068.frx":13E8
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Mantenimiento de Clientes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   585
            Left            =   5730
            Picture         =   "EvaCre_frm_068.frx":16F2
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_LevCon 
            Height          =   585
            Left            =   5160
            Picture         =   "EvaCre_frm_068.frx":19FC
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Levantar Condición de Aprobación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Excepc 
            Height          =   585
            Left            =   4590
            Picture         =   "EvaCre_frm_068.frx":1E3E
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Registro de Excepción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   4020
            Picture         =   "EvaCre_frm_068.frx":2148
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Registro de Observación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   3450
            Picture         =   "EvaCre_frm_068.frx":258A
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Impresión de Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Comite 
            Height          =   585
            Left            =   2880
            Picture         =   "EvaCre_frm_068.frx":29CC
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Registro de Comité de Créditos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ConCre 
            Height          =   585
            Left            =   2310
            Picture         =   "EvaCre_frm_068.frx":2E0E
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Condiciones Crediticias"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CalIng 
            Height          =   585
            Left            =   1740
            Picture         =   "EvaCre_frm_068.frx":3118
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Calificación de Ingresos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerMic 
            Height          =   585
            Left            =   1170
            Picture         =   "EvaCre_frm_068.frx":3422
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "MicroEmpresarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerLab 
            Height          =   585
            Left            =   600
            Picture         =   "EvaCre_frm_068.frx":3CEC
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Verificaciones Laborales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPer 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_068.frx":3FF6
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Verificaciones Personales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_068.frx":4300
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   45
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2055
         Left            =   30
         TabIndex        =   48
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
            TabIndex        =   49
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
         TabIndex        =   52
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
            TabIndex        =   53
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3889
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "EvaCre_frm_068.frx":4742
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "EvaCre_frm_068.frx":475E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "EvaCre_frm_068.frx":477A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(7)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Patrimonio"
            TabPicture(3)   =   "EvaCre_frm_068.frx":4796
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(4)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Referencias Personales"
            TabPicture(4)   =   "EvaCre_frm_068.frx":47B2
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Inmueble"
            TabPicture(5)   =   "EvaCre_frm_068.frx":47CE
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(2)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos del Crédito"
            TabPicture(6)   =   "EvaCre_frm_068.frx":47EA
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "txt_ObsSol"
            Tab(6).Control(1)=   "grd_Listad(5)"
            Tab(6).Control(2)=   "Label5"
            Tab(6).ControlCount=   3
            TabCaption(7)   =   "Docum. Recibidos"
            TabPicture(7)   =   "EvaCre_frm_068.frx":4806
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   54
               Text            =   "EvaCre_frm_068.frx":4822
               Top             =   1470
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   0
               Left            =   60
               TabIndex        =   55
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
               TabIndex        =   56
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
               TabIndex        =   57
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
               Left            =   -74940
               TabIndex        =   58
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
               TabIndex        =   59
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
               TabIndex        =   60
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
               TabIndex        =   61
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
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74910
               TabIndex        =   63
               Top             =   1470
               Width           =   1155
            End
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_AprCon     As Integer
Dim l_int_FlgRec     As Integer
Dim l_dbl_IngDP1     As Double
Dim l_dbl_IngDp2     As Double
Dim l_dbl_IngDp3     As Double
Dim l_dbl_IngDp4     As Double
Dim l_dbl_IngAd1     As Double
Dim l_dbl_IngAd2     As Double
Dim l_dbl_IngTot     As Double
Dim l_dbl_OblMen     As Double
Dim l_dbl_IngNet     As Double
Dim l_dbl_CuoSol     As Double
Dim l_dbl_CuoMPr     As Double
Dim l_dbl_TipCam     As Double
Dim l_str_FecRIn     As String
Dim l_str_FecCal     As String
Dim l_str_EmpSeg     As String
Dim l_dbl_MtoPre_Cal As Double
Dim l_int_PlaAno_Cal As Integer
Dim l_int_TipSeg_Cal As Integer
Dim l_int_PerGra_Cal As Integer
Dim l_int_CuoDbl_Cal As Integer
Dim l_int_FlgVDm     As Integer
Dim l_int_FlgVLb     As Integer
Dim l_int_FlgIng     As Integer
Dim l_int_FlgApr     As Integer
Dim l_int_FlgCmt     As Integer
Dim l_int_TipEva     As Integer
Dim l_str_PerAct     As String
Dim l_str_CodPry     As String

Private Sub cmd_AprCon_Click()
Dim r_int_DiaTra     As Integer
Dim r_str_Cadena     As String
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_int_Resul      As Integer
Dim r_str_CodMod     As String
Dim r_str_CodPrd     As String
Dim r_str_DesMod     As String

   r_str_CodMod = ""
   r_str_CodPrd = ""
   r_str_DesMod = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.SOLMAE_CODPRD, A.SOLMAE_CODMOD, B.EVACRE_FLGJFT, "
   r_str_Parame = r_str_Parame & "        TRIM(C.SUBPRD_DESCRI) AS MICROEMPRESARIO, "
   r_str_Parame = r_str_Parame & "        (CASE WHEN INSTR(C.SUBPRD_DESCRI,'MICROEMPRESARIO') > 0 THEN 1 ELSE 0 END) AS FLAG_MICRO "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE A "
   r_str_Parame = r_str_Parame & "   LEFT JOIN TRA_EVACRE B ON B.EVACRE_NUMSOL = A.SOLMAE_NUMERO "
   r_str_Parame = r_str_Parame & "   LEFT JOIN CRE_SUBPRD C ON C.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND C.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   r_str_Parame = r_str_Parame & "  WHERE A.SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      r_str_CodMod = Trim(r_rst_Princi!SOLMAE_CODMOD)
      r_str_CodPrd = Trim(r_rst_Princi!SOLMAE_CODPRD)
      r_str_DesMod = moddat_gf_Buscar_NomMod(Trim(r_str_CodPrd), r_str_CodMod)
   End If
   
   'Validacion de parametros de gastos de cierre
   If InStr(r_str_DesMod, "TERMINADO") = 0 Then
      r_int_Resul = gf_Valida_GastoCierre(r_str_CodPrd, l_str_CodPry)
   
      If r_int_Resul = 1 Then
         MsgBox "El proyecto asociado a la solicitud no tiene empresa de peritaje asignada, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      ElseIf r_int_Resul = 2 Then
         MsgBox "La empresa de peritaje asociado al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      ElseIf r_int_Resul = 3 Then
         'MsgBox "El proyecto asociado a la solicitud no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         'Exit Sub
         'MsgBox "Los gastos de cierre no se calcularán porque no se han registrado los parámetros de notaría, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
      ElseIf r_int_Resul = 4 Then
         MsgBox "La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   'Validacion de registro de informacion de creditos
   If grd_LisEva.Rows = 0 Then
       MsgBox "No se ha registrado información de la Evaluación de Créditos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_CalIng)
       Exit Sub
   End If
   If l_int_FlgVDm = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerPer)
       Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerLab)
       Exit Sub
   End If
   If l_int_FlgIng = 0 Then
       MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_CalIng)
       Exit Sub
   End If
   If l_int_FlgApr = 0 Then
       MsgBox "Debe ingresar la información de Condiciones del Crédito.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
   If l_int_FlgCmt = 0 Then
        MsgBox "Debe ingresar la información del Registro de Comité de Créditos.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmd_Comite)
        Exit Sub
   End If
   If CDate(l_str_FecCal) < CDate(l_str_FecRIn) Then
       MsgBox "La Fecha de Cálculo de Monto de Préstamo es menor a la Fecha de Cálculo de Ingresos. Por favor actualice los Cálculos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
   If moddat_g_int_TipMon <> 1 Then
       If CDate(l_str_FecCal) < date Then
           MsgBox "La Fecha de Cálculo de Monto de Préstamo no es igual a la Fecha Actual. Por favor actualice el Cálculo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmd_ConCre)
           Exit Sub
       End If
   End If
   
   'Validacion de la jefatura
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "210") = "" Then
      'If r_rst_Princi!FLAG_MICRO = 0 Then
         If r_rst_Princi!EVACRE_FLGJFT = 0 Then
            MsgBox "Esta solicitud no fue aprobada por su jefe de créditos.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      'End If
   End If
   
   'Verificacion de observaciones pendientes
   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 21) Then
     MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   'Validacion de numero de excepciones
   If fs_Valida_Excepc(moddat_g_str_NumSol, Year(CDate(l_str_PerAct)), Month(CDate(l_str_PerAct)), 21) = False Then
      MsgBox "La solicitud sobrepasa la regla de excepción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Verificando que Cliente no haya sido ingresado como Cliente Negativo
   If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   frm_RecSol_56.Show 1

   If moddat_g_int_FlgAct_1 = 1 Then
      Exit Sub
   End If
   
   'Creando Aprobación Condicionada
   If Not moddat_gf_Inserta_AprCon(moddat_g_str_NumSol, 21, moddat_g_str_Observ) Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 21)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 21, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando Datos de Aprobación en Solicitud de Crédito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLMAE_APRUEBA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDP1) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp2) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp3) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp4) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngAd1) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngAd2) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngTot) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_OblMen) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoSol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoMPr) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TipCam) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PlaAno_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_TipSeg_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_CuoDbl_Cal) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 31) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 31) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - APROBACION CONDICIONADA (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ & Chr(13)
   
   Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
   
   'GENERANDO GASTOS DE CIERRE
   If InStr(r_str_DesMod, "TERMINADO") = 0 Then                   'Si no es bien terminado
      
      If Not fs_Genera_GastoCierre(moddat_g_str_NumSol) Then
         Exit Sub
      End If
      
      'Enviar Correo de Asignación de Gastos de Cierre
      modgen_g_str_Mail_Asunto = "ASIGNACION DE GASTOS DE CIERRE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   End If
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct = 2
   Call cmd_Imprim_Click
   Unload Me
End Sub

Private Sub cmd_Aprueb_Click()
Dim r_int_DiaTra     As Integer
Dim r_str_Cadena     As String
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_int_Resul      As Integer
Dim r_str_CodMod     As String
Dim r_str_CodPrd     As String
Dim r_str_DesMod     As String

   r_str_CodMod = ""
   r_str_CodPrd = ""
   r_str_DesMod = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.SOLMAE_CODPRD, A.SOLMAE_CODMOD, B.EVACRE_FLGJFT, TRIM(C.SUBPRD_DESCRI) AS MICROEMPRESARIO, "
   r_str_Parame = r_str_Parame & "        (CASE WHEN INSTR(C.SUBPRD_DESCRI,'MICROEMPRESARIO') > 0 THEN 1 ELSE 0 END) AS FLAG_MICRO "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE A "
   r_str_Parame = r_str_Parame & "   LEFT JOIN TRA_EVACRE B ON B.EVACRE_NUMSOL = A.SOLMAE_NUMERO "
   r_str_Parame = r_str_Parame & "   LEFT JOIN CRE_SUBPRD C ON C.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND C.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   r_str_Parame = r_str_Parame & "  WHERE A.SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      r_str_CodMod = Trim(r_rst_Princi!SOLMAE_CODMOD)
      r_str_CodPrd = Trim(r_rst_Princi!SOLMAE_CODPRD)
      r_str_DesMod = moddat_gf_Buscar_NomMod(Trim(r_str_CodPrd), r_str_CodMod)
   End If
   
   'Validacion de parametros asociados a los gastos de cierre
   If InStr(r_str_DesMod, "TERMINADO") = 0 Then
      r_int_Resul = gf_Valida_GastoCierre(r_str_CodPrd, l_str_CodPry)
      
      If r_int_Resul = 1 Then
         MsgBox "El proyecto asociado a la solicitud no tiene empresa de peritaje asignado, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      ElseIf r_int_Resul = 2 Then
         MsgBox "El proyecto asociado a la solicitud no tiene notaría asignada, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      ElseIf r_int_Resul = 3 Then
         'MsgBox "La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         'Exit Sub
         'MsgBox "Los gastos de cierre no se calcularán porque no se han registrado los parámetros de notaría, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         If MsgBox("La notaria asociada al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor coordinar con el área legal la actualización de la información en caso contrario no se generaran los gastos de cierre." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
      ElseIf r_int_Resul = 4 Then
         MsgBox "La empresa de peritaje asociado al proyecto no tiene registrado los parámetros necesarios para el cálculo de los gastos de cierre, favor actualizar información en la plataforma de Operaciones.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   'Validacion de registro de informacion de creditos
   If grd_LisEva.Rows = 0 Then
      MsgBox "No se ha registrado información de la Evaluación de Créditos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalIng)
      Exit Sub
   End If
   If l_int_FlgVDm = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerPer)
      Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerLab)
      Exit Sub
   End If
   If l_int_FlgIng = 0 Then
        MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(cmd_CalIng)
        Exit Sub
   End If
   If l_int_FlgApr = 0 Then
       MsgBox "Debe ingresar la información de Condiciones del Crédito.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
   If l_int_FlgCmt = 0 Then
       MsgBox "Debe ingresar la información del Registro de Comité de Créditos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_Comite)
       Exit Sub
   End If
   If CDate(l_str_FecCal) < CDate(l_str_FecRIn) Then
       MsgBox "La Fecha de Cálculo de Monto de Préstamo es menor a la Fecha de Cálculo de Ingresos. Por favor actualice los Cálculos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
   If moddat_g_int_TipMon <> 1 Then
      If CDate(l_str_FecCal) < date Then
         MsgBox "La Fecha de Cálculo de Monto de Préstamo no es igual a la Fecha Actual. Por favor actualice el Cálculo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_ConCre)
         Exit Sub
      End If
   End If
   
   'Validacion de la Jefatura
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "210") = "" Then
      'If r_rst_Princi!FLAG_MICRO = 0 Then
         If r_rst_Princi!EVACRE_FLGJFT = 0 Then
            MsgBox "Esta solicitud no fue aprobada por su jefe de créditos.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      'End If
   End If
   
   'Verificacion de observaciones pendientes
   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 21) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Verificando observaciones pendientes
   If fs_Valida_Excepc(moddat_g_str_NumSol, Year(CDate(l_str_PerAct)), Month(CDate(l_str_PerAct)), 21) = False Then
      MsgBox "La solicitud sobrepasa la regla de excepción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Verificando que Cliente no haya sido ingresado como Cliente Negativo
   If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 21)))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 21, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
      
   'Actualizando Datos de Aprobación en Solicitud de Crédito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLMAE_APRUEBA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDP1) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp2) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp3) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngDp4) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngAd1) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngAd2) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngTot) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_OblMen) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoSol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CuoMPr) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TipCam) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PlaAno_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_TipSeg_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra_Cal) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_CuoDbl_Cal) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 31) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 31) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - APROBACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
     
   Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
   
   'GENERANDO GASTOS DE CIERRE
   If InStr(r_str_DesMod, "TERMINADO") = 0 Then                   'Si no es bien terminado
      If Not fs_Genera_GastoCierre(moddat_g_str_NumSol) Then
         Exit Sub
      End If
   
      'Enviar Correo de Asignación de Gastos de Cierre
      modgen_g_str_Mail_Asunto = "ASIGNACION DE GASTOS DE CIERRE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, False, False, False)
   End If
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct = 2
   Call cmd_Imprim_Click
   Unload Me
End Sub

Private Function fs_Asigna_GasAdm(p_NumSol As String, p_Codigo As Integer, p_Import As Double) As Integer
   fs_Asigna_GasAdm = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_ASIGNA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Left(p_Codigo, 2) & ", "
      g_str_Parame = g_str_Parame & Right(p_Codigo, 1) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(p_Import)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_Asigna_GasAdm = True
End Function

Private Function fs_Genera_GastoCierre(ByVal p_NumSol As String) As Integer
Dim r_rst_Genera        As ADODB.Recordset
Dim r_str_Princi        As String
Dim r_dbl_ValPre        As Double
Dim r_dbl_ValInm        As Double
Dim r_dbl_ValEst        As Double
Dim r_str_CodPry        As String
Dim r_dbl_ValGci        As Double
Dim r_dbl_MtoGci        As Double
Dim r_str_CodMod        As String

Dim r_dbl_GasTas        As Double
Dim r_dbl_GasNot        As Double
Dim r_dbl_BloReg        As Double
Dim r_dbl_RegMin        As Double
Dim r_dbl_RegHip        As Double
Dim r_dbl_ImpItf        As Double
Dim r_str_Operac        As String
Dim r_dbl_PorITF        As Double
Dim r_int_Contad        As Integer

   fs_Genera_GastoCierre = False
   
   r_dbl_ValPre = 0
   r_dbl_ValInm = 0
   r_dbl_ValEst = 0
   r_dbl_MtoGci = 0
   r_dbl_ValGci = 0
   r_str_CodPry = ""
   r_str_CodMod = ""
   
   r_dbl_GasTas = 0
   r_dbl_GasNot = 0
   r_dbl_BloReg = 0
   r_dbl_RegMin = 0
   r_dbl_RegHip = 0
   r_dbl_ImpItf = 0
   r_str_Operac = ""

   r_str_Princi = ""
   r_str_Princi = r_str_Princi & " SELECT SOLMAE_PREMTO, SOLMAE_MTOINM, SOLMAE_MTOEST, SOLMAE_MTOGCI, SOLMAE_CODPRD, SOLMAE_CODMOD, SOLINM_PRYCOD "
   r_str_Princi = r_str_Princi & "   FROM CRE_SOLMAE "
   r_str_Princi = r_str_Princi & "        INNER JOIN CRE_SOLINM ON SOLINM_NUMSOL = SOLMAE_NUMERO "
   r_str_Princi = r_str_Princi & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(r_str_Princi, g_rst_Princi, 3) Then
      fs_Genera_GastoCierre = False
   End If
   
   If Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      r_dbl_ValPre = g_rst_Princi!SOLMAE_PREMTO
      r_dbl_ValInm = g_rst_Princi!SOLMAE_MTOINM
      r_dbl_ValEst = g_rst_Princi!SOLMAE_MTOEST
      r_dbl_ValGci = g_rst_Princi!SOLMAE_MTOGCI
      r_str_CodPry = g_rst_Princi!SOLINM_PRYCOD
      r_str_CodMod = g_rst_Princi!SOLMAE_CODMOD
            
      r_dbl_MtoGci = gf_Genera_Gastos_Cierre(moddat_g_str_CodPrd, r_str_CodPry, Format(r_str_CodMod, "000"), r_dbl_ValInm, r_dbl_ValEst, r_dbl_ValPre, r_dbl_GasTas, r_dbl_GasNot, r_dbl_BloReg, r_dbl_RegMin, r_dbl_RegHip, r_dbl_ImpItf)
      r_dbl_MtoGci = Format(CDbl(r_dbl_MtoGci), "###,###,##0.00")
      
      If CDbl(r_dbl_MtoGci) <> CDbl(r_dbl_ValGci) Then
         If fs_Valida_FinGCi(p_NumSol) Then
            If Not fs_Actualiza_GasCie(moddat_g_str_NumSol, r_dbl_MtoGci) Then
               fs_Genera_GastoCierre = False
               Exit Function
            Else
               fs_Genera_GastoCierre = True
            End If
         End If
      End If
      
      '******************************************** INGRESO DE GASTOS DE CIERRE ***********************************************************
      'GASTOS DE TASACION
      If r_dbl_GasTas > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 111, r_dbl_GasTas) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'GASTOS NOTARIALES
      If r_dbl_GasNot > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 121, r_dbl_GasNot) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'ITF
      If r_dbl_GasNot > 0 And r_dbl_ImpItf > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 131, r_dbl_ImpItf) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'Si es Bien Terminado
      'GASTOS REGISTRALES - BLOQUEO REGISTRAL
      If r_dbl_BloReg > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 161, r_dbl_BloReg) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'GASTOS REGISTRALES - MINUTA DE COMPRA VENTA
      If r_dbl_RegMin > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 171, r_dbl_RegMin) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'GASTOS REGISTRALES - INCRIPCION DE GARANTIA
      If r_dbl_RegHip > 0 Then
         If Not fs_Asigna_GasAdm(moddat_g_str_NumSol, 181, r_dbl_RegHip) Then
            fs_Genera_GastoCierre = False
         End If
      End If
      
      'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 24, 0, "", 0, 0) Then 'moddat_g_int_CodIns
         fs_Genera_GastoCierre = False
      End If
      
      '******************************************* PAGO DE LOS GASTOS DE CIERRE **********************************************************
   '   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "210")
   '   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
      
      If fs_Valida_FinGCi(p_NumSol) Then
         'Espera 1 minuto para añadir los pagos
         r_int_Contad = 1
         
         Do While r_int_Contad < 61
            DoEvents
            r_int_Contad = r_int_Contad + 1
         Loop
         
         r_str_Operac = "99999"
         
         'Obteniendo ITF
         r_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
         
         'GASTOS NOTARIALES
         If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, Left(121, 2), moddat_g_int_TipMon, r_dbl_GasNot, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
            fs_Genera_GastoCierre = False
         End If
         
         'ITF
         If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, Left(131, 2), moddat_g_int_TipMon, r_dbl_ImpItf, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
            fs_Genera_GastoCierre = False
         End If
         
         'Si es Bien Terminado
         'GASTOS REGISTRALES - BLOQUEO REGISTRAL
         If r_dbl_BloReg > 0 Then
            If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, Left(161, 2), moddat_g_int_TipMon, r_dbl_BloReg, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
               fs_Genera_GastoCierre = False
            End If
         End If
         
         'GASTOS REGISTRALES - MINUTA DE COMPRA VENTA
         If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, Left(171, 2), moddat_g_int_TipMon, r_dbl_RegMin, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
            fs_Genera_GastoCierre = False
         End If
         
         'GASTOS REGISTRALES - INCRIPCION DE GARANTIA
         If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, Left(181, 2), moddat_g_int_TipMon, r_dbl_RegHip, r_dbl_PorITF, Format(CDate(Now), "yyyymmdd"), r_str_Operac) Then
            fs_Genera_GastoCierre = False
         End If
         
         'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 31, 25, 0, "", 0, 0) Then 'moddat_g_int_CodIns
            fs_Genera_GastoCierre = False
         End If
      End If
  End If
  
  g_rst_Princi.Close
  fs_Genera_GastoCierre = True
End Function

Private Sub cmd_Aprueb_ValJef_Click()
   If l_int_FlgVDm = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerPer)
       Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerLab)
       Exit Sub
   End If
   If l_int_FlgIng = 0 Then
       MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_CalIng)
       Exit Sub
   End If
   If l_int_FlgApr = 0 Then
       MsgBox "Debe ingresar la información de Condiciones del Crédito.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
   
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "210") <> "" Then
      moddat_g_int_FlgAct_2 = 1
      frm_EvaCre_72.Show 1
      
      If moddat_g_int_FlgAct_1 = 2 Then
         'Cargando Datos de Seguimiento
         Call fs_Buscar_LisOcu
      End If
   Else
      MsgBox "Esta opción está permitida solo al jefe de crédito.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_CalIng_Click()

   If l_int_FlgVDm = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerPer)
      Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerLab)
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_2 = 1
   frm_EvaCre_64.Show 1

   If moddat_g_int_FlgAct_2 = 2 Then
      Screen.MousePointer = 11
      
      'Buscando Información de Evaluación ya registrada
      Call fs_Buscar_DatEva
      
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Comite_Click()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   If l_int_FlgVDm = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerPer)
       Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
       MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_VerLab)
       Exit Sub
   End If
   If l_int_FlgIng = 0 Then
       MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_CalIng)
       Exit Sub
   End If
   If l_int_FlgApr = 0 Then
       MsgBox "Debe ingresar la información de Condiciones del Crédito.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmd_ConCre)
       Exit Sub
   End If
    
   moddat_g_int_FlgAct = 1
   
   frm_EvaCre_67.Show 1
    
   If moddat_g_int_FlgAct = 2 Then
       Screen.MousePointer = 11
       Call fs_Buscar_DatEva                               'Buscando Información de Evaluación ya registrada
       
       'Buscar Información del Crédito
       Call modmip_gs_DatCre(grd_Listad(5), r_arr_Mtz)
       txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
       moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
       moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
       moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
       l_int_TipEva = r_arr_Mtz(0).DatCom_TipEva
       l_str_EmpSeg = r_arr_Mtz(0).DatCom_EsgDes
   
       Call fs_Buscar_LisOcu
       Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ConCre_Click()
   If l_int_FlgVDm = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerPer)
      Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerLab)
      Exit Sub
   End If
   If l_int_FlgIng = 0 Then
      MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalIng)
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_2 = 1
   
   frm_EvaCre_65.Show 1

   If moddat_g_int_FlgAct_2 = 2 Then
      Screen.MousePointer = 11
      
      'Buscando Información de Evaluación ya registrada
      Call fs_Buscar_DatEva
      
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_DatCli_Click()
   moddat_g_int_FlgAct = 1
   moddat_g_int_FlgGrb = 2
   
   frm_MntCli_52.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Call gs_LimpiaGrid(grd_Listad(0))
      Call gs_LimpiaGrid(grd_Listad(1))
      Call gs_LimpiaGrid(grd_Listad(7))
   
      'Buscar Información de Solicitud de Crédito
      moddat_g_int_CygTDo = 0
      moddat_g_str_CygNDo = ""
   
      Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
      Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
      Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(7))         'Buscar Información del Apoderado
   End If

End Sub

Private Sub cmd_Excepc_Click()
Dim r_int_NumExc     As Integer
  
   moddat_g_str_Observ = ""
   moddat_g_int_TipAut = 0
   moddat_g_int_FlgAct_1 = 1
   frm_RecSol_55.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
   
      'Generando Número de Excepción
      r_int_NumExc = 0
      
      g_str_Parame = "SELECT COUNT(SEGEXC_NUMSOL) AS NUMREG FROM TRA_SEGEXC WHERE "
      g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         r_int_NumExc = g_rst_Princi!NUMREG
      End If
         
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      r_int_NumExc = r_int_NumExc + 1
      
      'Grabando en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 18, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Grabando en Detalle de Excepciones
      If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 21, r_int_NumExc, moddat_g_str_Observ, moddat_g_int_TipAut) Then
         Exit Sub
      End If
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - EXCEPCION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      'antes no se le enviaba
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
      
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      Call fs_Buscar_LisExc      'Buscando Excepciones
   
      'Si no hay Excepciones aplicadas
      If grd_LisExc.Rows = 0 Then
         tab_Seguim.TabVisible(1) = False
      Else
         tab_Seguim.TabVisible(1) = True
      End If
      
      Screen.MousePointer = 0
   End If
End Sub

Private Function fs_Valida_Excepc(ByVal p_NumSol As String, ByVal p_PerAnio As Integer, ByVal p_PerMes As String, ByVal p_CodIns As Integer) As Boolean
Dim r_dbl_PorExc     As Double
Dim r_int_ConExp     As Integer
Dim r_int_SinExp     As Integer
   
   fs_Valida_Excepc = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(SUBSTR(PARVAL_DESCRI,INSTR(PARVAL_DESCRI,'-')+2)) PORCENTAJE "
   g_str_Parame = g_str_Parame & "  FROM MNT_PARVAL "
   g_str_Parame = g_str_Parame & " WHERE PARVAL_CODGRP = 606 "
   g_str_Parame = g_str_Parame & "   AND SUBSTR(PARVAL_DESCRI,1,4) = '" & p_PerAnio & "'"
   g_str_Parame = g_str_Parame & "   AND SUBSTR(PARVAL_DESCRI,5,2) = '" & Format(p_PerMes, "00") & "'"
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      MsgBox "No existe porcebtaje de Excepción para este periodo.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   
   If Not g_rst_Listas.EOF Then
      r_dbl_PorExc = CDbl(g_rst_Listas!PORCENTAJE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   'Calcula Solicitudes con y sin excepcion del mes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT NRO_CRED_EXCEP, COUNT(NRO_CRED_EXCEP) AS CANTIDAD "
   g_str_Parame = g_str_Parame & "  FROM ( SELECT SOLMAE_FECSOL, CASE WHEN (NVL((  SELECT COUNT(*)  "
   g_str_Parame = g_str_Parame & "                                                   FROM TRA_SEGEXC "
   g_str_Parame = g_str_Parame & "                                                        LEFT JOIN MNT_PARDES ON (SEGEXC_MOTEXC=PARDES_CODITE AND PARDES_CODGRP = 42 ) "
   g_str_Parame = g_str_Parame & "                                                   WHERE SEGEXC_NUMSOL = SOL.SOLMAE_NUMERO AND SEGEXC_CODINS = " & p_CodIns & " ),0)) > 0 THEN 1 ELSE 0 END AS NRO_CRED_EXCEP "
   g_str_Parame = g_str_Parame & "           FROM CRE_SOLMAE SOL "
   g_str_Parame = g_str_Parame & "                INNER JOIN TRA_SEGUIM ON SEGUIM_NUMSOL = SOL.SOLMAE_NUMERO AND SEGUIM_CODINS = 21 AND SEGUIM_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                  AND SUBSTR(SEGUIM_FECFIN,1,4) = '" & p_PerAnio & "'"
   g_str_Parame = g_str_Parame & "                  AND SUBSTR(SEGUIM_FECFIN,5,2) = '" & Format(p_PerMes, "00") & "'"
   g_str_Parame = g_str_Parame & "          WHERE SOLMAE_SITUAC = 1 )"
   g_str_Parame = g_str_Parame & " GROUP BY NRO_CRED_EXCEP "
   g_str_Parame = g_str_Parame & " ORDER BY NRO_CRED_EXCEP "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      fs_Valida_Excepc = True
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   
   Do While Not g_rst_Listas.EOF
      If g_rst_Listas!NRO_CRED_EXCEP = 0 Then
         r_int_SinExp = g_rst_Listas!CANTIDAD
      ElseIf g_rst_Listas!NRO_CRED_EXCEP = 1 Then
         r_int_ConExp = g_rst_Listas!CANTIDAD
      End If
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   If r_int_SinExp > 0 Then
      If ((CInt(r_int_ConExp) + 1) / CInt(r_int_SinExp)) / 100 <= r_dbl_PorExc Then
         fs_Valida_Excepc = True
      End If
   Else
      fs_Valida_Excepc = True
   End If
End Function

Private Sub cmd_Imprim_Click()
Dim r_rst_MaeCli        As ADODB.Recordset
Dim r_rst_MaeSol        As ADODB.Recordset
Dim r_rst_SegCon        As ADODB.Recordset
Dim r_rst_Coment        As ADODB.Recordset
Dim r_str_Tit_CodPri    As String
Dim r_str_Tit_CodSec    As String
Dim r_int_Tit_TDoPri    As Integer
Dim r_str_Tit_NDoPri    As String
Dim r_str_Tit_EmpPri    As String
Dim r_int_Tit_TDoSec    As Integer
Dim r_str_Tit_NDoSec    As String
Dim r_str_Tit_EmpSec    As String
Dim r_str_Cyg_CodPri    As String
Dim r_str_Cyg_CodSec    As String
Dim r_int_Cyg_TDoPri    As Integer
Dim r_str_Cyg_NDoPri    As String
Dim r_str_Cyg_EmpPri    As String
Dim r_int_Cyg_TDoSec    As Integer
Dim r_str_Cyg_NDoSec    As String
Dim r_str_Cyg_EmpSec    As String
Dim r_int_FlgCon        As Integer
Dim r_str_AprCon        As String
Dim r_int_TipEva        As Integer
Dim r_str_Comentario    As String
   
   If grd_LisEva.Rows = 0 Then
      MsgBox "No se ha registrado información de la Evaluación de Créditos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalIng)
      Exit Sub
   End If
   If l_int_FlgVDm = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Personales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerPer)
      Exit Sub
   End If
   If l_int_FlgVLb = 0 Then
      MsgBox "Debe ingresar la información de las Verificaciones Laborales.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_VerLab)
      Exit Sub
   End If
   If l_int_FlgIng = 0 Then
      MsgBox "Debe ingresar la información de Calificación de Ingresos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_CalIng)
      Exit Sub
   End If
   If l_int_FlgApr = 0 Then
      MsgBox "Debe ingresar la información de Condiciones del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ConCre)
      Exit Sub
   End If
   If l_int_FlgCmt = 0 Then
         MsgBox "Debe ingresar la información del Registro de Comité de Créditos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_Comite)
         Exit Sub
   End If
   If CDate(l_str_FecCal) < CDate(l_str_FecRIn) Then
      MsgBox "La Fecha de Cálculo de Monto de Préstamo es menor a la Fecha de Cálculo de Ingresos. Por favor actualice los Cálculos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ConCre)
      Exit Sub
   End If
   If moddat_g_int_TipMon <> 1 Then
      If CDate(l_str_FecCal) < date Then
         MsgBox "La Fecha de Cálculo de Monto de Préstamo no es igual a la Fecha Actual. Por favor actualice el Cálculo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_ConCre)
         Exit Sub
      End If
   End If
   If moddat_gf_Valida_Observ(moddat_g_str_NumSol, 21) Then
      MsgBox "La solicitud presenta Observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Ficha de Evaluación Crediticia?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM RPT_FICCRE WHERE FICCRE_NOMTER = '" & modgen_g_str_NombPC & "' AND FICCRE_NUMSOL = '" & gf_Formato_NumSol(moddat_g_str_NumSol) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_FICCRE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'Para obtener Actividades Económicas del Cliente Titular
   r_str_Tit_CodPri = ""
   r_str_Tit_CodSec = ""
   r_int_Tit_TDoPri = 0
   r_str_Tit_NDoPri = ""
   r_str_Tit_EmpPri = ""
   r_int_Tit_TDoSec = 0
   r_str_Tit_NDoSec = ""
   r_str_Tit_EmpSec = ""
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!ActEco_OrdAct = 1 Then
            r_str_Tit_CodPri = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
         Else
            r_str_Tit_CodSec = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
         End If
            
         Select Case g_rst_Princi!ACTECO_CODACT
            Case 11
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  r_int_Tit_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                  r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
               Else
                  r_int_Tit_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                  r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
               End If
               
            Case 21
               If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Tit_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                  Else
                     r_int_Tit_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                     r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                  End If
               End If
               
            Case 41
               If g_rst_Princi!ActEco_OrdAct = 1 Then
                  r_int_Tit_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                  r_str_Tit_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
               Else
                  r_int_Tit_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                  r_str_Tit_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
               End If
         End Select
            
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscando Calificación de Empresas (Actividad Principal - Titular)
   r_str_Tit_EmpPri = ""
   
   If r_int_Tit_TDoPri > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Tit_TDoPri) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Tit_NDoPri & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         r_str_Tit_EmpPri = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   'Buscando Calificación de Empresas (Actividad Secundaria - Titular)
   r_str_Tit_EmpSec = ""
   
   If r_int_Tit_TDoSec > 0 Then
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Tit_TDoSec) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Tit_NDoSec & "' "
    
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         r_str_Tit_EmpSec = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   'Cónyuge
   r_str_Cyg_CodPri = ""
   r_str_Cyg_CodSec = ""
   r_int_Cyg_TDoPri = 0
   r_str_Cyg_NDoPri = ""
   r_str_Cyg_EmpPri = ""
   r_int_Cyg_TDoSec = 0
   r_str_Cyg_NDoSec = ""
   r_str_Cyg_EmpSec = ""
   
   If moddat_g_int_CygTDo > 0 Then
      'Para obtener Actividades Económicas del Cliente Cónyuge
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' "
      g_str_Parame = g_str_Parame & "ORDER BY ACTECO_ORDACT ASC"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!ActEco_OrdAct = 1 Then
               r_str_Cyg_CodPri = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
            Else
               r_str_Cyg_CodSec = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ACTECO_CODACT))
            End If
               
            Select Case g_rst_Princi!ACTECO_CODACT
               Case 11
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Dep_TipDoc
                     r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  Else
                     r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Dep_TipDoc
                     r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
                  End If
                  
               Case 21
                  If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
                     If g_rst_Princi!ActEco_OrdAct = 1 Then
                        r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     Else
                        r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Ind_TipDoc_Emp
                        r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
                     End If
                  End If
                  
               Case 41
                  If g_rst_Princi!ActEco_OrdAct = 1 Then
                     r_int_Cyg_TDoPri = g_rst_Princi!ActEco_Acc_TipDoc
                     r_str_Cyg_NDoPri = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  Else
                     r_int_Cyg_TDoSec = g_rst_Princi!ActEco_Acc_TipDoc
                     r_str_Cyg_NDoSec = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
                  End If
            End Select
               
            g_rst_Princi.MoveNext
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      'Buscando Calificación de Empresas (Actividad Principal - Cónyuge)
      r_str_Cyg_EmpPri = ""
      
      If r_int_Cyg_TDoPri > 0 Then
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Cyg_TDoPri) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Cyg_NDoPri & "' "
       
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_str_Cyg_EmpPri = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End If
   
      'Buscando Calificación de Empresas (Actividad Secundaria - Cónyuge)
      r_str_Cyg_EmpSec = ""
      
      If r_int_Cyg_TDoSec > 0 Then
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(r_int_Cyg_TDoSec) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & r_str_Cyg_NDoSec & "' "
       
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            r_str_Cyg_EmpSec = Trim(g_rst_Princi!DATGEN_NOMCOM & "") & " / " & Trim(g_rst_Princi!DATGEN_RAZSOC)
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End If
   End If

   'Abriendo Cursor sobre Maestro de Clientes
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MaeCli, 3) Then
       Exit Sub
   End If
   
   r_rst_MaeCli.MoveFirst
   
   'Abriendo Cursor sobre Maestro de Solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MaeSol, 3) Then
       Exit Sub
   End If
   
   r_rst_MaeSol.MoveFirst
   r_int_TipEva = r_rst_MaeSol!SOLMAE_TIPEVA
   
   'Abriendo Cursos sobre Aprobación Condicionada
   r_int_FlgCon = 0
   r_str_AprCon = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGCON_CODINS = 21"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SegCon, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_SegCon.BOF And r_rst_SegCon.EOF) Then
      r_rst_SegCon.MoveFirst
      
      r_int_FlgCon = 1
      r_str_AprCon = Trim(r_rst_SegCon!SEGCON_OBSCON & "")
   End If
   
   'Obtener el comentario de datos de una solicitud
   g_str_Parame = ""
   g_str_Parame = "SELECT SEGDET_OBSERV FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 AND SEGDET_CODOCU = 17 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Coment, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Coment.BOF And r_rst_Coment.EOF) Then
      r_rst_Coment.MoveFirst
      r_str_Comentario = Trim(r_rst_Coment!SEGDET_OBSERV & "")
   End If
   
   'Abriendo Cursos sobre Evaluación Crediticia
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   g_str_Parame = "INSERT INTO RPT_FICCRE ("
   g_str_Parame = g_str_Parame & "FICCRE_NOMTER, "
   g_str_Parame = g_str_Parame & "FICCRE_NUMSOL, "
   g_str_Parame = g_str_Parame & "FICCRE_PRODUC, "
   g_str_Parame = g_str_Parame & "FICCRE_SUBPRD, "
   g_str_Parame = g_str_Parame & "FICCRE_TIPEVA, "
   g_str_Parame = g_str_Parame & "FICCRE_CLINOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CLIDOC, "
   g_str_Parame = g_str_Parame & "FICCRE_CLIEST, "
   g_str_Parame = g_str_Parame & "FICCRE_CYGNOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CYGDOC, "
   g_str_Parame = g_str_Parame & "FICCRE_MONCRE, "
   g_str_Parame = g_str_Parame & "FICCRE_MONSIM, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLCVT, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLINI, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLPRE, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLPLA, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLGRA, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLTSG, "
   g_str_Parame = g_str_Parame & "FICCRE_TASINT, "
   g_str_Parame = g_str_Parame & "FICCRE_VERDOM, "
   g_str_Parame = g_str_Parame & "FICCRE_REFPER, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGLIQ, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_TOTDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INGDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_INIDEU, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_CUOSOL, "
   g_str_Parame = g_str_Parame & "FICCRE_TOT_CUOMPR, "
   g_str_Parame = g_str_Parame & "FICCRE_APRPRE, "
   g_str_Parame = g_str_Parame & "FICCRE_APRPLA, "
   g_str_Parame = g_str_Parame & "FICCRE_APRGRA, "
   g_str_Parame = g_str_Parame & "FICCRE_APRTSG, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEFLG, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEFEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIEENT, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL0, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL1, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL2, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL3, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECL4, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_TDEUMN, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_TDEUME, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_RIECOM, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_ACTPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_EMPPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_VERPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_ACTSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_EMPSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_VERSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGADI, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_TIT_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEFLG, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEFEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIEENT, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL0, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL1, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL2, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL3, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECL4, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_TDEUMN, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_TDEUME, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_RIECOM, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_ACTPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_EMPPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_VERPRI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_ACTSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_EMPSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_VERSEC, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGADI, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_OBLMEN, "
   g_str_Parame = g_str_Parame & "FICCRE_CYG_INGNET, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOENT, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOMON, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOMTO, "
   g_str_Parame = g_str_Parame & "FICCRE_AHOTPO, "
   g_str_Parame = g_str_Parame & "FICCRE_FLGCON, "
   g_str_Parame = g_str_Parame & "FICCRE_APRCON, "
   g_str_Parame = g_str_Parame & "FICCRE_SOLCUOEXT, "
   g_str_Parame = g_str_Parame & "FICCRE_APRCUOEXT, "
   g_str_Parame = g_str_Parame & "FICCRE_COMENT) "
   
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(moddat_g_str_NumSol) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_MaeSol!SOLMAE_CODPRD) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_SubPrd(r_rst_MaeSol!SOLMAE_CODPRD, r_rst_MaeSol!SOLMAE_CODSUB) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_MaeSol!SOLMAE_TIPEVA)) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NomCli & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & "-" & Trim(moddat_g_str_NumDoc) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("205", r_rst_MaeCli!DATGEN_ESTCIV) & "', "
   
   If moddat_g_int_CygTDo > 0 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_CygTDo) & "-" & Trim(moddat_g_str_CygNDo) & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", r_rst_MaeSol!SOLMAE_TIPMON) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", r_rst_MaeSol!SOLMAE_TIPMON) & "', "
   
   If r_rst_MaeSol!SOLMAE_TIPMON = 2 Then
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_COMVTA_DOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_APOPRO_DOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOPRE_DOL) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_COMVTA_SOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_APOPRO_SOL) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOPRE_SOL) & ", "
   End If
   
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_PLAANO) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_PERGRA) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(r_rst_MaeSol!SOLMAE_ESGDES, r_rst_MaeSol!SOLMAE_TIPSEG) & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_TASINT) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("067", CStr(g_rst_Princi!EVACRE_TIPVDM)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_OBSVDM) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_REFPER) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGTOT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OBLMEN) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_MTODEU) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGNET) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_RINGDE) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_RINIDE) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CUOSOL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CUOMPR) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_MTOPRE_CAL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(r_rst_MaeSol!SOLMAE_ESGDES, g_rst_Princi!EVACRE_TIPSEG_CAL) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_TIT_CRIFLG)) & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_TIT_CRIFEC)) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRIENT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL0) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL1) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL2) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL3) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_CRICL4) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_TOTDMN) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_TIT_TOTDME) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_TIT_CRIOBS) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_CodPri) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_EmpPri) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP1) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_TIT_TIPVE1)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_TIT_LABVE1) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_CodSec) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(r_str_Tit_EmpSec) & "', "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP2) & ", "
   
   If Len(Trim(g_rst_Princi!EVACRE_TIT_LABVE2)) > 0 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_TIT_TIPVE2)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_TIT_LABVE2) & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGAD1) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OMETIT) & ", "
   g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INTTIT) & ", "
   
   If (g_rst_Princi!EVACRE_CYG_CRIFLG > 0) And (g_rst_Princi!EVACRE_INTCYG > 0) Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG)) & "', "
      g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC)) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRIENT) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL0) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL1) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL2) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL3) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_CRICL4) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_TOTDMN) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_CYG_TOTDME) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!EVACRE_CYG_CRIOBS) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_CodPri) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_EmpPri) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP3) & ", "
      
      If Len(Trim(g_rst_Princi!EVACRE_CYG_LABVE1)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_CYG_TIPVE1)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_CYG_LABVE1) & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_CodSec) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(r_str_Cyg_EmpSec) & "', "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGDP4) & ", "
      
      If Len(Trim(g_rst_Princi!EVACRE_CYG_LABVE2)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("068", CStr(g_rst_Princi!EVACRE_CYG_TIPVE2)) & Chr(10) & Chr(13) & Trim(g_rst_Princi!EVACRE_CYG_LABVE2) & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INGAD2) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_OMECYG) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!EVACRE_INTCYG) & ", "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   If r_rst_MaeSol!SOLMAE_TIPEVA = 2 Then
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("505", r_rst_MaeSol!SOLMAE_INSFIN) & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", r_rst_MaeSol!SOLMAE_MONAHO) & "', "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MTOAHO) & ", "
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_MESAHO) & ", "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   g_str_Parame = g_str_Parame & CStr(r_int_FlgCon) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_AprCon & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT) & ", "
   If IsNull(r_rst_MaeSol!SOLMAE_CUOEXT_CAL) Then
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT) & ", "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_MaeSol!SOLMAE_CUOEXT_CAL) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & CStr(r_str_Comentario) & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "Error al insertar registro en tabla RPT_FICCRE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_rst_MaeCli.Close
   Set r_rst_MaeCli = Nothing
   
   r_rst_MaeSol.Close
   Set r_rst_MaeSol = Nothing
   
   r_rst_SegCon.Close
   Set r_rst_SegCon = Nothing
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "RPT_FICCRE"
   crp_Imprim.SelectionFormula = "{RPT_FICCRE.FICCRE_NUMSOL} = '" & gf_Formato_NumSol(moddat_g_str_NumSol) & "' AND {RPT_FICCRE.FICCRE_NOMTER} = '" & modgen_g_str_NombPC & "'"
   
   If r_int_TipEva = 1 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_01.RPT"
   ElseIf r_int_TipEva = 2 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_02.RPT"
   Else
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_FICEVA_03.RPT"
   End If
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_LevCon_Click()
   moddat_g_int_CodIns = 11
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_57.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      If Not moddat_gf_Inserta_LevCon(moddat_g_str_NumSol, 11, moddat_g_str_Observ) Then
         Exit Sub
      End If
      
      'Buscando Aprobaciones Condicionadas
      Call fs_Buscar_LisCon
   
      'Si no hay Aprobaciones Condicionadas Pendiente
      If l_int_AprCon = 0 Then
         pnl_AprCon.Visible = False
         cmd_LevCon.Enabled = False
      End If
      
      Screen.MousePointer = 0
      
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - LEVANTAMIENTO DE CONDICIONES DE APROBACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ & Chr(13)
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
   End If
End Sub

Private Sub cmd_NueObs_Click()
 Dim r_int_NumObs     As Integer
   
   moddat_g_int_TipObs = 0
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_54.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      If moddat_g_int_TipObs = 1 Then
         'Generando Número de Observación
         r_int_NumObs = 0
            
         g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
         g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
         g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 AND "
         g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
         g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
             Exit Sub
         End If
      
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            g_rst_Princi.MoveFirst
            Do While Not g_rst_Princi.EOF
               r_int_NumObs = r_int_NumObs + 1
               g_rst_Princi.MoveNext
            Loop
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         r_int_NumObs = r_int_NumObs + 1
   
         'Grabando en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 21, CStr(r_int_NumObs), moddat_g_str_Observ, 1, 0) Then
            Exit Sub
         End If
         
         'Actualizando en Instancia si es una Observación
         If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 21, 0, 3, 2) Then
            Exit Sub
         End If
   
         'Enviando Correo Electrónico
         modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - OBSERVACION (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      Else
         'Grabando en Detalle de Seguimiento
         If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 17, 0, moddat_g_str_Observ, 0, 0) Then
            Exit Sub
         End If
         
         modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - COMENTARIO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      End If
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      'antes no se enviaba al jefe
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
      
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Rechaz_Click()
Dim r_int_DiaTra     As Integer
Dim r_str_CodIns     As String
Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 21
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(moddat_gf_FecIng_Ins(moddat_g_str_NumSol, 21)))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 21, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 21, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = "EVALUACION CREDITICIA - RECHAZO (Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, "", 0, True, False, True)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub cmd_SolRec_Click()
   frm_RecSol_53.Show 1
End Sub

Private Sub cmd_VerLab_Click()
   moddat_g_int_FlgAct_2 = 1
   frm_EvaCre_66.Show 1

   If moddat_g_int_FlgAct_2 = 2 Then
      Screen.MousePointer = 11
      'Buscando Información de Evaluación ya registrada
      Call fs_Buscar_DatEva
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_VerMic_Click()
   'moddat_g_int_FlgAct_2 = 1
   Screen.MousePointer = 11
   frm_EvaCre_68.Show 1
   Screen.MousePointer = 0
End Sub

Private Sub cmd_VerPer_Click()
   moddat_g_int_FlgAct_2 = 1
   frm_EvaCre_63.Show 1
   
   If moddat_g_int_FlgAct_2 = 2 Then
      Screen.MousePointer = 11
      'Buscando Información de Evaluación ya registrada
      Call fs_Buscar_DatEva
      'Cargando Datos de Seguimiento
      Call fs_Buscar_LisOcu
      Screen.MousePointer = 0
   End If
End Sub

Private Sub fs_Reingr()
'NRO DE VECES REINGRESANDO LA SOLICITUD DE UN CLIENTE
Dim r_int_nrovcs As Integer
   
   r_int_nrovcs = 0

   g_str_Parame = "SELECT COUNT(DATGEN_NUMDOC) AS CNTSOL  "
   g_str_Parame = g_str_Parame & " FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " LEFT JOIN CLI_DATGEN ON (SOLMAE_TITTDO=DATGEN_TIPDOC AND SOLMAE_TITNDO=DATGEN_NUMDOC) "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_TITNDO=" & moddat_g_str_NumDoc & " "
   g_str_Parame = g_str_Parame & " AND SOLMAE_TITTDO=" & moddat_g_int_TipDoc & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_int_nrovcs = Val(g_rst_Princi!CNTSOL)
   
   If r_int_nrovcs = 0 Or r_int_nrovcs = 1 Then
      pnl_Reingr.Visible = False
   Else
      pnl_Reingr.Caption = "El cliente presenta " & r_int_nrovcs & " reingresos"
      pnl_Reingr.Visible = True
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Me.Caption = modgen_g_str_NomPlt
   moddat_g_int_CodIns = 21
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call fs_Reingr
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
    
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(7))         'Buscar Información del Apoderado
   
   Call modmip_gs_DatCre(grd_Listad(5), r_arr_Mtz)                                      'Buscar Información del Crédito
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
   l_int_TipEva = r_arr_Mtz(0).DatCom_TipEva
   l_str_EmpSeg = r_arr_Mtz(0).DatCom_EsgDes
   
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Información del Inmueble
   Call fs_DatPat          'Datos del Patrimonio
   Call fs_DatRef          'Referencias Personales
   'Call fs_DatCre          'Datos del Crédito
   Call fs_SolDoc          'Documentos Recibidos
   
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   'Si no hay Aprobaciones Condicionadas Pendiente
   If l_int_AprCon = 0 Then
      pnl_AprCon.Visible = False
      cmd_LevCon.Enabled = False
   End If
   
   'Verificando si Cliente tiene Solicitudes Rechazadas anteriormente
   If Not ff_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      cmd_SolRec.Enabled = False
   End If
   
   'Busca si es Microempresario
   If fs_MicEmp = False Then
      cmd_VerMic.Enabled = False
   End If
   
   'Cargando Datos de Evaluación Crediticia
   Call fs_Buscar_DatEva
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
Function fs_MicEmp() As Boolean
   
   fs_MicEmp = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM CRE_SUBPRD  "
   g_str_Parame = g_str_Parame & "  WHERE SUBPRD_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "    AND SUBPRD_CODSUB = '" & moddat_g_str_CodSub & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If InStr(Trim(g_rst_Princi!SUBPRD_DESCRI), "MICROEMPRESARIO") Then
         fs_MicEmp = True
      End If
   End If
End Function
Private Sub fs_DatRef()
   Call gs_LimpiaGrid(grd_Listad(3))

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SOLREF_TIPREF ASC, SOLREF_NUMREF ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!SOLREF_TIPPAR <> 8 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1:       grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0:                             grd_Listad(3).Text = "Referencia: "
            
            grd_Listad(3).Col = 1
            
            If g_rst_Princi!SOLREF_TIPREF = 1 Then
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
            ElseIf g_rst_Princi!SOLREF_TIPREF = 2 Then
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
            Else
               grd_Listad(3).Text = moddat_gf_Consulta_ParDes("271", CStr(g_rst_Princi!SOLREF_TIPPAR))
            End If
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Apellidos y Nombres"
            grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
         
            If Len(Trim(g_rst_Princi!SOLREF_TELEFO & "")) > 0 Then
               grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
               grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Teléfono"
               grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
            End If
         
            If Len(Trim(g_rst_Princi!SOLREF_CELULA & "")) > 0 Then
               grd_Listad(3).Rows = grd_Listad(3).Rows + 1:    grd_Listad(3).Row = grd_Listad(3).Rows - 1
               grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Celular"
               grd_Listad(3).Col = 1:                          grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
            End If
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         End If
   
         g_rst_Princi.MoveNext
      Loop
      
      grd_Listad(3).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(4).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Fecha de Adquisición (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                    grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Dirección (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:          grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                                grd_Listad(4).Text = "INMUEBLES"
      grd_Listad(4).Col = 1:                                grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Número de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Línea Crédito (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Pago Mínimo (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "TARJETAS DE CREDITO"
      grd_Listad(4).Col = 1:                             grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:          grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                                grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Número de Operación (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Monto del Préstamo (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "DEUDAS"
      grd_Listad(4).Col = 1:                          grd_Listad(4).Text = "NO REGISTRA"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "GASTOS MENSUALES"
      grd_Listad(4).Col = 1:                             grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(6).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1:    grd_Listad(6).Row = grd_Listad(6).Rows - 1
   
      grd_Listad(6).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Function fs_UserEjecutivo(p_CodUsu As String, p_CodEje As String) As String
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
    
   fs_UserEjecutivo = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.EJETIP_CODEJE FROM CRE_EJETIP A, CRE_EJECMC B "
   r_str_Parame = r_str_Parame & "  WHERE A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   r_str_Parame = r_str_Parame & "    AND A.EJETIP_TIPEJE = " & p_CodEje
   r_str_Parame = r_str_Parame & "    AND A.EJETIP_TIPEJE = " & p_CodEje
   r_str_Parame = r_str_Parame & "    AND B.EJECMC_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "    AND UPPER(TRIM(A.EJETIP_CODEJE)) = '" & UCase(Trim(p_CodUsu)) & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   fs_UserEjecutivo = Trim(r_rst_Princi!EJETIP_CODEJE & "")
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Function

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de Cónyuge
    For r_int_Contad = 0 To 5
       grd_Listad(r_int_Contad).ColWidth(0) = 2900:    grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
       grd_Listad(r_int_Contad).ColWidth(1) = 7950:    grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
       
       Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
    Next r_int_Contad
   
    grd_Listad(6).ColWidth(0) = 10850:     grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
    
    grd_Listad(7).ColWidth(0) = 2900:      grd_Listad(7).ColAlignment(0) = flexAlignLeftCenter
    grd_Listad(7).ColWidth(1) = 7950:      grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter
    
    Call gs_LimpiaGrid(grd_Listad(7))
    
    'Lista de Ocurrencias
    grd_LisOcu.ColWidth(0) = 1155
    grd_LisOcu.ColWidth(1) = 1185
    grd_LisOcu.ColWidth(2) = 8595
    grd_LisOcu.ColWidth(3) = 0
    grd_LisOcu.ColWidth(4) = 0
    
    grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
    grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
    grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter
    
    Call gs_LimpiaGrid(grd_LisOcu)
    
    pnl_DesOcu.Caption = ""
    txt_Observ.Text = ""
    txt_Descar.Text = ""
    
    'Lista de Excepciones
    grd_LisExc.ColWidth(0) = 1175
    grd_LisExc.ColWidth(1) = 1175
    grd_LisExc.ColWidth(2) = 3275
    grd_LisExc.ColWidth(3) = 5325
    grd_LisExc.ColWidth(4) = 0
    grd_LisExc.ColWidth(5) = 0
    
    grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
    grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
    grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
    grd_LisExc.ColAlignment(3) = flexAlignLeftCenter
    
    Call gs_LimpiaGrid(grd_LisExc)
    
    pnl_DesExc.Caption = ""
    txt_ObsExc.Text = ""
    pnl_TipAut.Caption = ""
    pnl_Motivo.Caption = ""
    
    'Lista de Aprobaciones Condicionadas
    grd_LisCon.ColWidth(0) = 2735
    grd_LisCon.ColWidth(1) = 6605
    grd_LisCon.ColWidth(2) = 1625
    grd_LisCon.ColWidth(3) = 0
    
    grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
    grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
    grd_LisCon.ColAlignment(2) = flexAlignLeftCenter
    
    Call gs_LimpiaGrid(grd_LisCon)
    
    pnl_InsCon.Caption = ""
    txt_ObsCon.Text = ""
    txt_LevCon.Text = ""
    
    'Lista de Datos de Evaluación
    grd_LisEva.ColWidth(0) = 3300
    grd_LisEva.ColWidth(1) = 7940
    
    grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
    grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
    
    'inicializar variables
    l_int_FlgVDm = 0
    l_int_FlgVLb = 0
    l_int_FlgIng = 0
    l_int_FlgApr = 0
    l_int_FlgCmt = 0
    
    l_str_PerAct = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
    l_str_CodPry = gf_Obtener_Proyec(moddat_g_str_NumSol)
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21    "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisOcu.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
      grd_LisOcu.Row = grd_LisOcu.Rows - 1
      
      'Número de Observación
      'grd_LisOcu.Col = 0
      'grd_LisOcu.Text = Format(g_rst_Princi!SEGDET_NUMOBS, "000")
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripción Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGFECACT > 0 Then
         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         
         grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & r_str_FecOcu
         grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call grd_LisOcu_Click
End Sub

Private Sub fs_Buscar_LisExc()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   
'   g_str_Parame = "SELECT * FROM TRA_SEGEXC WHERE "
'   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
'   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   g_str_Parame = modgen_gf_Buscar_Excepc
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepción
      grd_LisExc.Col = 5
      grd_LisExc.Text = g_rst_Princi!PARDES_DESCRI
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisExc)
   Call grd_LisExc_Click
End Sub

Private Sub fs_Buscar_LisCon()
   l_int_AprCon = 0
   
   Call gs_LimpiaGrid(grd_LisCon)
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripción Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCon)
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisCon_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 0
      pnl_InsCon.Caption = grd_LisCon.Text
   
      grd_LisCon.Col = 1
      txt_ObsCon.Text = grd_LisCon.Text
      
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      Call gs_RefrescaGrid(grd_LisCon)
   End If
End Sub

Private Sub grd_LisCon_SelChange()
   If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_LisExc_Click()
   Dim r_str_FecExc     As String
   Dim r_str_HorExc     As String
   Dim r_str_InsExc     As String

   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 0
      r_str_FecExc = grd_LisExc.Text
      
      grd_LisExc.Col = 1
      r_str_HorExc = grd_LisExc.Text
      
      grd_LisExc.Col = 2
      r_str_InsExc = grd_LisExc.Text
      
      pnl_DesExc.Caption = "Día: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
             
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         grd_LisExc.Col = 5
         pnl_Motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
         pnl_Motivo.Visible = True
         lbl_motivo.Visible = True
      Else
         pnl_Motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_Motivo.Caption = ""
      End If
       
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
      pnl_Motivo.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
'   Call grd_LisExc_Click
End Sub

Private Sub grd_LisOcu_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisOcu.Rows > 0 Then
      grd_LisOcu.Col = 0
      r_str_FecOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 1
      r_str_HorOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 2
      r_str_DesOcu = grd_LisOcu.Text
      
      pnl_DesOcu.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
'   If grd_LisOcu.Rows > 2 Then
'      grd_LisOcu.RowSel = grd_LisOcu.Row
'   End If
'
'   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub


Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub


Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Buscar_DatEva()
   Call gs_LimpiaGrid(grd_LisEva)
   
   l_dbl_IngDP1 = 0
   l_dbl_IngDp2 = 0
   l_dbl_IngDp3 = 0
   l_dbl_IngDp4 = 0
   l_dbl_IngAd1 = 0
   l_dbl_IngAd2 = 0
   l_dbl_IngTot = 0
   l_dbl_OblMen = 0
   l_dbl_IngNet = 0
   l_dbl_CuoSol = 0
   l_dbl_CuoMPr = 0
   l_str_FecRIn = ""
   l_dbl_TipCam = 0
   l_dbl_MtoPre_Cal = 0
   l_int_PlaAno_Cal = 0
   l_int_TipSeg_Cal = 0
   l_int_PerGra_Cal = 0
   l_int_CuoDbl_Cal = 0
   l_str_FecCal = ""
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT B.*, "
   g_str_Parame = g_str_Parame & "        (CASE WHEN INSTR(C.SUBPRD_DESCRI,'MICROEMPRESARIO') > 0 THEN 1 ELSE 0 END) AS FLAG_MICRO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_EVACRE B ON B.EVACRE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_SUBPRD C ON C.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND C.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   g_str_Parame = g_str_Parame & "  WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   cmd_Aprueb_ValJef.Visible = False
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "210") <> "" Then
      'SE ACTIVA SI NO ES MICROEMPRESARIO
      'If g_rst_Princi!FLAG_MICRO = 0 Then
         cmd_Aprueb_ValJef.Visible = True
      'End If
   End If
   
   g_rst_Princi.MoveFirst
   l_dbl_IngDP1 = g_rst_Princi!EVACRE_INGDP1
   l_dbl_IngDp2 = g_rst_Princi!EVACRE_INGDP2
   l_dbl_IngDp3 = g_rst_Princi!EVACRE_INGDP3
   l_dbl_IngDp4 = g_rst_Princi!EVACRE_INGDP4
   l_dbl_IngAd1 = g_rst_Princi!EVACRE_INGAD1
   l_dbl_IngAd2 = g_rst_Princi!EVACRE_INGAD2
   l_dbl_IngTot = g_rst_Princi!EVACRE_INGTOT
   l_dbl_OblMen = g_rst_Princi!EVACRE_OBLMEN
   l_dbl_IngNet = g_rst_Princi!EVACRE_INGNET
   l_dbl_CuoSol = g_rst_Princi!EVACRE_CUOSOL
   l_dbl_CuoMPr = g_rst_Princi!EVACRE_CUOMPR
   l_str_FecRIn = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_FECING))
   
   l_dbl_TipCam = g_rst_Princi!EVACRE_TIPCAM
   l_dbl_MtoPre_Cal = g_rst_Princi!EVACRE_MTOPRE_CAL
   l_int_PlaAno_Cal = g_rst_Princi!EVACRE_PLAANO_CAL
   l_int_TipSeg_Cal = g_rst_Princi!EVACRE_TIPSEG_CAL
   l_int_PerGra_Cal = g_rst_Princi!EVACRE_PERGRA_CAL
   l_int_CuoDbl_Cal = g_rst_Princi!EVACRE_CUODBL_CAL
   l_str_FecCal = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_FECCAL))
   
   If Not (IsNull(g_rst_Princi!EVACRE_FECCOM) And IsNull(g_rst_Princi!EVACRE_NROACT)) Then
       l_int_FlgCmt = 1
   End If
   
   'Llenando Grid
   If g_rst_Princi!EVACRE_INGTOT > 0 Then
      l_int_FlgIng = 1
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso Líquido"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGTOT, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Obligaciones Mensuales"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OBLMEN, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso Neto"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota Máxima Aprob."
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Cuota Máximo Aprob. (M. Prest.)"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2) & " (Tipo de Cambio: S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
      End If
   End If
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      l_int_FlgApr = 1
      
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Monto Préstamo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Plazo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " Años "
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Período de Gracia Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota Extraordinaria Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = IIf(g_rst_Princi!EVACRE_CUODBL_CAL = 0, "NO", moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!EVACRE_CUODBL_CAL)))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Tipo de Seguro Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Tipo Cambio de Aprobación"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 14, 4)
      End If
   End If
   
   'Verificación Domiciliaria
   If Len(Trim(g_rst_Princi!EVACRE_TIPVDM & "")) > 0 Then
      l_int_FlgVDm = 1
      If g_rst_Princi!EVACRE_TIPVDM > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificación Domiciliaria"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("067", CStr(g_rst_Princi!EVACRE_TIPVDM))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_OBSVDM & "")
      End If
   End If
   
   'Central de Riesgo
   If Not IsNull(g_rst_Princi!EVACRE_TIT_CRIFLG) Then
      If g_rst_Princi!EVACRE_TIT_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - Titular Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_TIT_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_TIT_CRIFEC))
      End If
      
      If g_rst_Princi!EVACRE_TIT_CRIFLG = 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_TIT_CRIENT)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificación"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = "Nor: " & Format(g_rst_Princi!EVACRE_TIT_CRICL0, "##0.00") & "%" & " / " & "Cpp: " & Format(g_rst_Princi!EVACRE_TIT_CRICL1, "##0.00") & "%" & " / " & "Def: " & Format(g_rst_Princi!EVACRE_TIT_CRICL2, "##0.00") & "%" & " / " & "Dud: " & Format(g_rst_Princi!EVACRE_TIT_CRICL3, "##0.00") & "%" & " / " & "Per: " & Format(g_rst_Princi!EVACRE_TIT_CRICL4, "##0.00") & "%"
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda MN"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_TOTDMN, 12, 2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda ME"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_TOTDME, 12, 2)
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (1)"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN1)) & ")"

         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE1, 12, 2)
        
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN2 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN2)) & ")"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE2, 12, 2)
            
            
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN3 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN3)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE3, 12, 2)
            
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN4 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN4)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE4, 12, 2)

         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN5 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN5)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE5, 12, 2)

         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN6 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN6)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_LIMDE6, 12, 2)

         End If
      End If
      
      'Central de Riesgo (Cónyuge)
      If g_rst_Princi!EVACRE_CYG_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - Cónyuge Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC))
         
         If g_rst_Princi!EVACRE_CYG_CRIFLG = 1 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_CYG_CRIENT)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificación"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = "Nor: " & Format(g_rst_Princi!EVACRE_CYG_CRICL0, "##0.00") & "%" & " / " & "Cpp: " & Format(g_rst_Princi!EVACRE_CYG_CRICL1, "##0.00") & "%" & " / " & "Def: " & Format(g_rst_Princi!EVACRE_CYG_CRICL2, "##0.00") & "%" & " / " & "Dud: " & Format(g_rst_Princi!EVACRE_CYG_CRICL3, "##0.00") & "%" & " / " & "Per: " & Format(g_rst_Princi!EVACRE_CYG_CRICL4, "##0.00") & "%"
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda MN"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_TOTDMN, 12, 2)
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda ME"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_TOTDME, 12, 2)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (1)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN1)) & ")"

            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN1, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE1, 12, 2)

            
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN2 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN2)) & ")"

               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN2, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE2, 12, 2)
           
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN3 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN3)) & ")"

               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN3, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE3, 12, 2)

            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN4 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN4)) & ")"

               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN4, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE4, 12, 2)

               
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN5 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN5)) & ")"

               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN5, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE5, 12, 2)
               
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN6 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN6)) & ")"
   
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = ""
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = " LINEA UTILIZADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN6, 12, 2) & " - " & " LINEA ASIGNADA: " & "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_LIMDE6, 12, 2)

            End If
         End If
      End If
   End If
   
   'Referencias Personales
   If Len(Trim(g_rst_Princi!EVACRE_REFPER & "")) > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificación Referencias"
      grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_REFPER & "")
   End If
   
   'Actividades Económicas
   If Not IsNull(g_rst_Princi!EVACRE_TIT_TIPVE1) Then
      If g_rst_Princi!EVACRE_TIT_TIPVE1 > 0 Then
         l_int_FlgVLb = 1
         
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE1 & "")
      End If
      
      If g_rst_Princi!EVACRE_TIT_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Titular - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_TIT_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_TIT_LABVE2 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE1 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Cónyuge - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE1 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Cónyuge - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE2 & "")
      End If
   End If
   
   'Ingresos
   If g_rst_Princi!EVACRE_INGDP1 > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso Líquido Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP2, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD1, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1 + g_rst_Princi!EVACRE_INGDP2 + g_rst_Princi!EVACRE_INGAD1, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Obligaciones Mensuales Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMETIT, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso Neto Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTTIT, 12, 2)
   
      If g_rst_Princi!EVACRE_INGDP3 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Líquido Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP4, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD2, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3 + g_rst_Princi!EVACRE_INGDP4 + g_rst_Princi!EVACRE_INGAD2, 12, 2)
      End If
   
      If g_rst_Princi!EVACRE_OMECYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Obligaciones Mensuales Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMECYG, 12, 2)
      End If
      
      If g_rst_Princi!EVACRE_INTCYG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Neto Cónyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTCYG, 12, 2)
      End If
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTODEU, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Ingreso / Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINGDE, 12, 2)
   
      grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Inicial / Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINIDE, 12, 2) & "%"
   End If
   
   Call gs_UbiIniGrid(grd_LisEva)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function ff_Buscar_SolRec(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Integer
   ff_Buscar_SolRec = False
   
   'Buscando Solicitudes Rechazadas como Cliente Titular
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      ff_Buscar_SolRec = True
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Function
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscando Solicitudes Rechazadas como Cónyuge
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGNDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      ff_Buscar_SolRec = True
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Function
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Function fs_Actualiza_GasCie(ByVal p_NumSol As String, ByVal p_GasCie As Double) As Integer

Dim r_rst_Grabar  As ADODB.Recordset
   
   fs_Actualiza_GasCie = False
   
   g_str_Parame = "UPDATE CRE_SOLMAE SET "
   g_str_Parame = g_str_Parame & "   SOLMAE_MTOGCI = " & p_GasCie & " , "
   g_str_Parame = g_str_Parame & "   SOLMAE_MTOPRE_SOL = SOLMAE_PREMTO + " & CDbl(p_GasCie) & ", "
   g_str_Parame = g_str_Parame & "   SOLMAE_MTOPRE_MPR = SOLMAE_PREMTO + " & CDbl(p_GasCie) & ""
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND SOLMAE_MTOGCI > 0 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Function
   End If
   
   fs_Actualiza_GasCie = True
End Function
Private Function fs_Valida_FinGCi(ByVal p_NumSol As String) As Integer

Dim r_rst_Genera  As ADODB.Recordset
Dim r_str_Parame  As String
   
   fs_Valida_FinGCi = False
   
   r_str_Parame = r_str_Parame & " SELECT SOLMAE_MTOGCI FROM CRE_SOLMAE  "
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   r_str_Parame = r_str_Parame & "    AND SOLMAE_MTOGCI > 0 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      If r_rst_Genera!SOLMAE_MTOGCI > 0 Then
         fs_Valida_FinGCi = True
      Else
         fs_Valida_FinGCi = False
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing

   End If
End Function
