VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_LevCon_14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   885
   ClientTop       =   1470
   ClientWidth     =   11610
   Icon            =   "EvaCre_frm_051.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   17754
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
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "EvaCre_frm_051.frx":000C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(4)=   "SSPanel5"
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(9)=   "txt_Observ"
            Tab(0).Control(10)=   "txt_Descar"
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "EvaCre_frm_051.frx":0028
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label6"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label4"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "lbl_motivo"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "pnl_Motivo"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "pnl_TipAut"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "pnl_DesExc"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "SSPanel16"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "SSPanel15"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "SSPanel12"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "SSPanel11"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).Control(11)=   "SSPanel9"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "grd_LisExc"
            Tab(1).Control(12).Enabled=   0   'False
            Tab(1).Control(13)=   "txt_ObsExc"
            Tab(1).Control(13).Enabled=   0   'False
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobaci�n Condicionada"
            TabPicture(2)   =   "EvaCre_frm_051.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_ObsCon"
            Tab(2).Control(1)=   "txt_LevCon"
            Tab(2).Control(2)=   "SSPanel17"
            Tab(2).Control(3)=   "grd_LisCon"
            Tab(2).Control(4)=   "SSPanel18"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel20"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "Label15"
            Tab(2).Control(9)=   "Label14"
            Tab(2).Control(10)=   "Label12"
            Tab(2).ControlCount=   11
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "EvaCre_frm_051.frx":0060
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "EvaCre_frm_051.frx":0064
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   1230
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "EvaCre_frm_051.frx":0068
               Top             =   1980
               Width           =   10035
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "EvaCre_frm_051.frx":006C
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
               Text            =   "EvaCre_frm_051.frx":0070
               Top             =   1980
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   -74970
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
               Left            =   -74970
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
               Left            =   -74940
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
               Left            =   -72600
               TabIndex        =   11
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Ocurrencia"
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
               Left            =   -73770
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
               Left            =   -73680
               TabIndex        =   13
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               Left            =   30
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
               Left            =   60
               TabIndex        =   15
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepci�n"
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
               Left            =   5670
               TabIndex        =   16
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripci�n Excepci�n"
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
               Left            =   1230
               TabIndex        =   17
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepci�n"
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
               Left            =   2400
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
               Left            =   30
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
               Left            =   1230
               TabIndex        =   20
               Top             =   1650
               Width           =   10035
               _Version        =   65536
               _ExtentX        =   17701
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               Left            =   1230
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
               Caption         =   "Situaci�n"
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
               Caption         =   "Condiciones de Aprobaci�n"
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
               Caption         =   "D�a: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
               Left            =   6060
               TabIndex        =   65
               Top             =   2970
               Width           =   5205
               _Version        =   65536
               _ExtentX        =   9181
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "MOTIVO"
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
               Left            =   5400
               TabIndex        =   66
               Top             =   3030
               Width           =   705
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   36
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observaci�n:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   34
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripci�n:"
               Height          =   495
               Left            =   60
               TabIndex        =   33
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label3 
               Caption         =   "Excepci�n:"
               Height          =   315
               Left            =   60
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   90
               TabIndex        =   31
               Top             =   3000
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
               Caption         =   "Condiciones de Aprobaci�n:"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   690
            TabIndex        =   51
            Top             =   60
            Width           =   7185
            _Version        =   65536
            _ExtentX        =   12674
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cr�ditos Hipotecarios - Levantamiento de Aprobaci�n Condicionada"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   690
            TabIndex        =   52
            Top             =   330
            Width           =   4635
            _Version        =   65536
            _ExtentX        =   8176
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Evaluaci�n Crediticia"
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
            Picture         =   "EvaCre_frm_051.frx":0074
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "EvaCre_frm_051.frx":037E
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_LevCon 
            Height          =   585
            Left            =   30
            Picture         =   "EvaCre_frm_051.frx":07C0
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Levantar Condici�n de Aprobaci�n"
            Top             =   30
            Width           =   585
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   4710
            Top             =   60
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
            Left            =   4140
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   42
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
            TabIndex        =   43
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   44
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
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   9450
            TabIndex        =   45
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8070
            TabIndex        =   46
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1875
         Left            =   30
         TabIndex        =   49
         Top             =   8130
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3307
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
            Height          =   1755
            Left            =   60
            TabIndex        =   50
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3096
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   2325
         Left            =   30
         TabIndex        =   53
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
            TabIndex        =   54
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
            TabPicture(0)   =   "EvaCre_frm_051.frx":0C02
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "C�nyuge"
            TabPicture(1)   =   "EvaCre_frm_051.frx":0C1E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "EvaCre_frm_051.frx":0C3A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(7)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Patrimonio"
            TabPicture(3)   =   "EvaCre_frm_051.frx":0C56
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(4)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Referencias Personales"
            TabPicture(4)   =   "EvaCre_frm_051.frx":0C72
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Inmueble"
            TabPicture(5)   =   "EvaCre_frm_051.frx":0C8E
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(2)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos del Cr�dito"
            TabPicture(6)   =   "EvaCre_frm_051.frx":0CAA
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "Label5"
            Tab(6).Control(1)=   "grd_Listad(5)"
            Tab(6).Control(2)=   "txt_ObsSol"
            Tab(6).ControlCount=   3
            TabCaption(7)   =   "Docum. Recibidos"
            TabPicture(7)   =   "EvaCre_frm_051.frx":0CC6
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   55
               Text            =   "EvaCre_frm_051.frx":0CE2
               Top             =   1470
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   0
               Left            =   60
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
               Index           =   1
               Left            =   -74940
               TabIndex        =   57
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
               TabIndex        =   58
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
               TabIndex        =   59
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   3
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
               Index           =   7
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
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74910
               TabIndex        =   64
               Top             =   1470
               Width           =   1155
            End
         End
      End
   End
End
Attribute VB_Name = "frm_LevCon_14"
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

Private Sub cmd_LevCon_Click()
   moddat_g_int_CodIns = 21
   moddat_g_str_Observ = ""
   moddat_g_int_FlgAct_1 = 1
   
   frm_RecSol_17.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      
      If Not moddat_gf_Inserta_LevCon(moddat_g_str_NumSol, 21, moddat_g_str_Observ) Then
         Exit Sub
      End If
      
      Screen.MousePointer = 0
      
      'Enviando Correo Electr�nico
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
      
      MsgBox "Se levanto las condiciones de la Aprobaci�n.", vbInformation, modgen_g_str_NomPlt
      moddat_g_int_FlgAct = 2
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   moddat_g_int_CodIns = 21
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   
   'Buscar Informaci�n de Solicitud de Cr�dito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Informaci�n del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Informaci�n del C�nyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(7))         'Buscar Informaci�n del Apoderado
   
   Call modmip_gs_DatCre(grd_Listad(5), r_arr_Mtz)                                       'Buscar Informaci�n del Cr�dito
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   moddat_g_str_FecIng = r_arr_Mtz(0).DatCom_FecSol
   l_str_EmpSeg = r_arr_Mtz(0).DatCom_EsgDes
  'Call fs_DatCre          'Datos del Cr�dito
  
   Call modmip_gs_DatInm(grd_Listad(2), False)                                            'Buscar Informaci�n del Inmueble
   Call fs_DatPat          'Datos del Patrimonio
   Call fs_DatRef          'Referencias Personales
   Call fs_SolDoc          'Documentos Recibidos
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If
   
   'Cargando Datos de Evaluaci�n Crediticia
   Call fs_Buscar_DatEva
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
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
      
      'N�mero de Observaci�n
      'grd_LisOcu.Col = 0
      'grd_LisOcu.Text = Format(g_rst_Princi!SEGDET_NUMOBS, "000")
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripci�n Ocurrencia
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
      
      'Fecha de Excepci�n
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepci�n
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripci�n Excepci�n
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorizaci�n
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepci�n
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
      
      'Descripci�n Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situaci�n
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripci�n Levantamiento Condiciones
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
      
      pnl_DesExc.Caption = "D�a: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
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
   
   Call grd_LisExc_Click
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
      
      pnl_DesOcu.Caption = "D�a: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
   
   Call grd_LisOcu_Click
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
   l_str_FecCal = ""

   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
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
   l_str_FecCal = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_FECCAL))
   
   'Llenando Grid
   If g_rst_Princi!EVACRE_INGTOT > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Total Ingreso L�quido"
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
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota M�xima Aprob."
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Cuota M�ximo Aprob. (M. Prest.)"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2) & " (Tipo de Cambio: S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 14, 4)
      End If
   End If
   
   If g_rst_Princi!EVACRE_FECCAL > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Monto Pr�stamo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Plazo Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " A�os "
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Per�odo de Gracia Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Cuota Extraordinaria Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!EVACRE_CUODBL_CAL))
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                       grd_LisEva.Text = "Tipo de Seguro Aprobado"
      grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1: grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                    grd_LisEva.Text = "Tipo Cambio de Aprobaci�n"
         grd_LisEva.Col = 1:                    grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:           grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 14, 4)
      End If
   End If
   
   'Verificaci�n Domiciliaria
   If Len(Trim(g_rst_Princi!EVACRE_TIPVDM & "")) > 0 Then
      If g_rst_Princi!EVACRE_TIPVDM > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificaci�n Domiciliaria"
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
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificaci�n"
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
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN1, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN1)) & ")"
         
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN2 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN2, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN2)) & ")"
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN3 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN3, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN3)) & ")"
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN4 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN4, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN4)) & ")"
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN5 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN5, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN5)) & ")"
         End If
      
         If Len(Trim(g_rst_Princi!EVACRE_TIT_CODEN6 & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
            grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_TIT_DEUEN6, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_TIT_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_TIT_CLAEN6)) & ")"
         End If
      End If
      
      'Central de Riesgo (C�nyuge)
      If g_rst_Princi!EVACRE_CYG_CRIFLG > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Central de Riesgo - C�nyuge Reportado"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_CYG_CRIFLG))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Fecha Reporte"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACRE_CYG_CRIFEC))
         
         If g_rst_Princi!EVACRE_CYG_CRIFLG = 1 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Nro. Entidades:"
            grd_LisEva.Col = 1:                          grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_CYG_CRIENT)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                          grd_LisEva.Text = "Clasificaci�n"
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
            grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN1, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN1) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN1)) & ")"
            
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN2 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (2)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN2, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN2) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN2)) & ")"
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN3 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (3)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN3, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN3) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN3)) & ")"
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN4 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (4)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN4, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN4) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN4)) & ")"
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN5 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (5)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN5, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN5) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN5)) & ")"
            End If
         
            If Len(Trim(g_rst_Princi!EVACRE_CYG_CODEN6 & "")) > 0 Then
               grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
               grd_LisEva.Col = 0:                          grd_LisEva.Text = "Deuda por Entidad (6)"
               grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
               grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_CYG_DEUEN6, 12, 2) & " - ENTIDAD: " & moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!EVACRE_CYG_CODEN6) & "") & " (CLASIF.: " & moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!EVACRE_CYG_CLAEN6)) & ")"
            End If
         End If
      End If
   End If
   
   'Referencias Personales
   If Len(Trim(g_rst_Princi!EVACRE_REFPER & "")) > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Verificaci�n Referencias"
      grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_REFPER & "")
   End If
   
   'Actividades Econ�micas
   If Not IsNull(g_rst_Princi!EVACRE_TIT_TIPVE1) Then
      If g_rst_Princi!EVACRE_TIT_TIPVE1 > 0 Then
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
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "C�nyuge - Verif. Lab. (Act. Princ.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE1)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE1 & "")
      End If
   
      If g_rst_Princi!EVACRE_CYG_TIPVE2 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "C�nyuge - Verif. Lab. (Act. Secund.)"
         grd_LisEva.Col = 1:                          grd_LisEva.Text = moddat_gf_Consulta_ParDes("068", g_rst_Princi!EVACRE_CYG_TIPVE2)
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 1:                          grd_LisEva.Text = Trim(g_rst_Princi!EVACRE_CYG_LABVE2 & "")
      End If
   End If
   
   'Ingresos
   If g_rst_Princi!EVACRE_INGDP1 > 0 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso L�quido Titular"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP2, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD1, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP1 + g_rst_Princi!EVACRE_INGDP2 + g_rst_Princi!EVACRE_INGAD1, 12, 2)
   
      If Not IsNull(g_rst_Princi!EVACRE_OMETIT) Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Obligaciones Mensuales Titular"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMETIT, 12, 2)
      End If
   
      If Not IsNull(g_rst_Princi!EVACRE_INTTIT) Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ingreso Neto Titular"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTTIT, 12, 2)
      End If
   
      If g_rst_Princi!EVACRE_INGDP3 > 0 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso L�quido C�nyuge"
         grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP4, 12, 2) & " + " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGAD2, 12, 2) & " = " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGDP3 + g_rst_Princi!EVACRE_INGDP4 + g_rst_Princi!EVACRE_INGAD2, 12, 2)
      End If
   
      If Not IsNull(g_rst_Princi!EVACRE_OMECYG) Then
         If g_rst_Princi!EVACRE_OMECYG > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                       grd_LisEva.Text = "Obligaciones Mensuales C�nyuge"
            grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_OMECYG, 12, 2)
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVACRE_INTCYG) > 0 Then
         If g_rst_Princi!EVACRE_INTCYG > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1:    grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0:                       grd_LisEva.Text = "Ingreso Neto C�nyuge"
            grd_LisEva.Col = 1:                       grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8:              grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INTCYG, 12, 2)
         End If
      End If
   
      grd_LisEva.Rows = grd_LisEva.Rows + 2:       grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0:                          grd_LisEva.Text = "Total Deuda"
      grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_MTODEU, 12, 2)
   
      If Not IsNull(g_rst_Princi!EVACRE_RINGDE) Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Ingreso / Deuda"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINGDE, 12, 2)
      End If
   
      If Not IsNull(g_rst_Princi!EVACRE_RINIDE) Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1:       grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0:                          grd_LisEva.Text = "Ratio Inicial / Deuda"
         grd_LisEva.Col = 1:                          grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8:                 grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_RINIDE, 12, 2) & "%"
      End If
   End If
   
   Call gs_UbiIniGrid(grd_LisEva)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


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
               grd_Listad(3).Col = 0:                          grd_Listad(3).Text = "Tel�fono"
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
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Fecha de Adquisici�n (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                             grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                    grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:       grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                             grd_Listad(4).Text = "Direcci�n (" & Format(r_int_Contad, "00") & ")"
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
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "N�mero de Tarjeta (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "L�nea Cr�dito (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8:                 grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Pago M�nimo (" & Format(r_int_Contad, "00") & ")"
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
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Instituci�n Financiera (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "N�mero de Operaci�n (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            grd_Listad(4).Col = 1:                          grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1:    grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0:                          grd_Listad(4).Text = "Monto del Pr�stamo (" & Format(r_int_Contad, "00") & ")"
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
         'Buscar en Par�metros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Par�metros por Actividad Econ�mica
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

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de C�nyuge
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

   'Lista de Datos de Evaluaci�n
   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940
   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

