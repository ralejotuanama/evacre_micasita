VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_EvaCre_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10575
   ClientLeft      =   465
   ClientTop       =   405
   ClientWidth     =   14340
   Icon            =   "EvaCre_frm_015.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14340
      _Version        =   65536
      _ExtentX        =   25294
      _ExtentY        =   18653
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1335
         Left            =   11010
         TabIndex        =   214
         Top             =   8370
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   2355
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
         Begin VB.TextBox txt_Eva_Observ 
            Height          =   945
            Left            =   90
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   216
            Top             =   330
            Width           =   3105
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Observaciones Finales:"
            Height          =   225
            Index           =   82
            Left            =   90
            TabIndex        =   215
            Top             =   60
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   4785
         Left            =   30
         TabIndex        =   63
         Top             =   2670
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
         _ExtentY        =   8440
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
         Begin TabDlg.SSTab tab_Princi 
            Height          =   4695
            Left            =   60
            TabIndex        =   64
            Top             =   60
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   8281
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Inform. Financ."
            TabPicture(0)   =   "EvaCre_frm_015.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "tab_InfFin"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Refer. Personales"
            TabPicture(1)   =   "EvaCre_frm_015.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_Fam_TipPar"
            Tab(1).Control(1)=   "pnl_Fam_Nombre"
            Tab(1).Control(2)=   "pnl_Fam_Telefo"
            Tab(1).Control(3)=   "pnl_Fam_Celula"
            Tab(1).Control(4)=   "pnl_NFa_TipPar"
            Tab(1).Control(5)=   "pnl_NFa_Nombre"
            Tab(1).Control(6)=   "pnl_NFa_Telefo"
            Tab(1).Control(7)=   "pnl_NFa_Celula"
            Tab(1).Control(8)=   "lbl_Etique(23)"
            Tab(1).Control(9)=   "lbl_Etique(22)"
            Tab(1).Control(10)=   "lbl_Etique(21)"
            Tab(1).Control(11)=   "lbl_Etique(20)"
            Tab(1).Control(12)=   "lbl_Etique(19)"
            Tab(1).Control(13)=   "lbl_Etique(18)"
            Tab(1).Control(14)=   "lbl_Etique(17)"
            Tab(1).Control(15)=   "lbl_Etique(16)"
            Tab(1).Control(16)=   "lbl_Etique(15)"
            Tab(1).Control(17)=   "lbl_Etique(14)"
            Tab(1).ControlCount=   18
            TabCaption(2)   =   "Datos Inmueb."
            TabPicture(2)   =   "EvaCre_frm_015.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "pnl_TipPro"
            Tab(2).Control(1)=   "pnl_NatTit"
            Tab(2).Control(2)=   "pnl_NatCyg"
            Tab(2).Control(3)=   "pnl_JurEmp"
            Tab(2).Control(4)=   "pnl_JurDir"
            Tab(2).Control(5)=   "pnl_JurRep"
            Tab(2).Control(6)=   "pnl_Direcc"
            Tab(2).Control(7)=   "lbl_Etique(32)"
            Tab(2).Control(8)=   "lbl_Etique(31)"
            Tab(2).Control(9)=   "lbl_Etique(30)"
            Tab(2).Control(10)=   "lbl_Etique(29)"
            Tab(2).Control(11)=   "lbl_Etique(28)"
            Tab(2).Control(12)=   "lbl_Etique(27)"
            Tab(2).Control(13)=   "lbl_Etique(26)"
            Tab(2).Control(14)=   "lbl_Etique(25)"
            Tab(2).Control(15)=   "lbl_Etique(24)"
            Tab(2).ControlCount=   16
            TabCaption(3)   =   "Datos Solicitud"
            TabPicture(3)   =   "EvaCre_frm_015.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_Sol_Observ"
            Tab(3).Control(1)=   "pnl_Sol_FecIng"
            Tab(3).Control(2)=   "pnl_Sol_NumCuo"
            Tab(3).Control(3)=   "pnl_Sol_PerGra"
            Tab(3).Control(4)=   "pnl_Sol_CuoExt"
            Tab(3).Control(5)=   "pnl_Sol_ComVta_Dol"
            Tab(3).Control(6)=   "pnl_Sol_ApoPro_Dol"
            Tab(3).Control(7)=   "pnl_Sol_MtoSol_Dol"
            Tab(3).Control(8)=   "pnl_Sol_ComVta_MPr"
            Tab(3).Control(9)=   "pnl_Sol_ApoPro_MPr"
            Tab(3).Control(10)=   "pnl_Sol_MtoSol_MPr"
            Tab(3).Control(11)=   "pnl_Sol_ComVta_Sol"
            Tab(3).Control(12)=   "pnl_Sol_ApoPro_Sol"
            Tab(3).Control(13)=   "pnl_Sol_MtoSol_Sol"
            Tab(3).Control(14)=   "pnl_Sol_TcaDol"
            Tab(3).Control(15)=   "pnl_Sol_TcaMPr"
            Tab(3).Control(16)=   "pnl_Sol_EjeVta"
            Tab(3).Control(17)=   "pnl_Sol_CuoSeg"
            Tab(3).Control(18)=   "pnl_Sol_VctSug"
            Tab(3).Control(19)=   "pnl_Sol_TipSeg"
            Tab(3).Control(20)=   "lbl_Etique(52)"
            Tab(3).Control(21)=   "lbl_Etique(51)"
            Tab(3).Control(22)=   "lbl_Etique(50)"
            Tab(3).Control(23)=   "lbl_Etique(49)"
            Tab(3).Control(24)=   "lbl_Etique(48)"
            Tab(3).Control(25)=   "lbl_Etique(47)"
            Tab(3).Control(26)=   "lbl_Etique(46)"
            Tab(3).Control(27)=   "lbl_Etique(45)"
            Tab(3).Control(28)=   "lbl_Etique(44)"
            Tab(3).Control(29)=   "lbl_Etique(43)"
            Tab(3).Control(30)=   "lbl_Etique(42)"
            Tab(3).Control(31)=   "lbl_Etique(41)"
            Tab(3).Control(32)=   "lbl_Etique(40)"
            Tab(3).Control(33)=   "lbl_Etique(39)"
            Tab(3).Control(34)=   "lbl_Etique(38)"
            Tab(3).Control(35)=   "lbl_Etique(37)"
            Tab(3).Control(36)=   "lbl_Etique(36)"
            Tab(3).Control(37)=   "lbl_Etique(35)"
            Tab(3).Control(38)=   "lbl_Etique(34)"
            Tab(3).Control(39)=   "lbl_Etique(33)"
            Tab(3).ControlCount=   40
            TabCaption(4)   =   "Docum. Recepc."
            TabPicture(4)   =   "EvaCre_frm_015.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Doc_Listad"
            Tab(4).Control(1)=   "SSPanel24"
            Tab(4).ControlCount=   2
            Begin TabDlg.SSTab tab_InfFin 
               Height          =   4185
               Left            =   120
               TabIndex        =   139
               Top             =   420
               Width           =   10485
               _ExtentX        =   18494
               _ExtentY        =   7382
               _Version        =   393216
               Style           =   1
               Tabs            =   4
               TabsPerRow      =   4
               TabHeight       =   520
               TabCaption(0)   =   "Inmuebles"
               TabPicture(0)   =   "EvaCre_frm_015.frx":0098
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lbl_Etique(56)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "lbl_Etique(54)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "lbl_Etique(55)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "lbl_Etique(53)"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "pnl_Inm_Import"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "pnl_Inm_Direcc"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "pnl_Inm_FecAdq"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "pnl_Inm_TipInm"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "SSPanel20"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "SSPanel18"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "SSPanel17"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "grd_Inm_Listad"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "SSPanel8"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).ControlCount=   13
               TabCaption(1)   =   "Tarjetas de Crédito"
               TabPicture(1)   =   "EvaCre_frm_015.frx":00B4
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "lbl_Etique(63)"
               Tab(1).Control(1)=   "lbl_Etique(62)"
               Tab(1).Control(2)=   "lbl_Etique(61)"
               Tab(1).Control(3)=   "lbl_Etique(60)"
               Tab(1).Control(4)=   "lbl_Etique(58)"
               Tab(1).Control(5)=   "lbl_Etique(59)"
               Tab(1).Control(6)=   "lbl_Etique(57)"
               Tab(1).Control(7)=   "SSPanel68"
               Tab(1).Control(8)=   "SSPanel67"
               Tab(1).Control(9)=   "pnl_Tar_MtoMin"
               Tab(1).Control(10)=   "pnl_Tar_LinCre"
               Tab(1).Control(11)=   "pnl_Tar_TipMon"
               Tab(1).Control(12)=   "pnl_Tar_NumTar"
               Tab(1).Control(13)=   "pnl_Tar_SalPag"
               Tab(1).Control(14)=   "pnl_Tar_NomIns"
               Tab(1).Control(15)=   "pnl_Tar_TipTar"
               Tab(1).Control(16)=   "SSPanel15"
               Tab(1).Control(17)=   "SSPanel14"
               Tab(1).Control(18)=   "SSPanel13"
               Tab(1).Control(19)=   "grd_Tar_Listad"
               Tab(1).ControlCount=   20
               TabCaption(2)   =   "Deudas Financieras"
               TabPicture(2)   =   "EvaCre_frm_015.frx":00D0
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "lbl_Etique(68)"
               Tab(2).Control(1)=   "lbl_Etique(69)"
               Tab(2).Control(2)=   "lbl_Etique(70)"
               Tab(2).Control(3)=   "lbl_Etique(67)"
               Tab(2).Control(4)=   "lbl_Etique(65)"
               Tab(2).Control(5)=   "lbl_Etique(66)"
               Tab(2).Control(6)=   "lbl_Etique(64)"
               Tab(2).Control(7)=   "SSPanel69"
               Tab(2).Control(8)=   "pnl_Fin_MesPag"
               Tab(2).Control(9)=   "pnl_Fin_CuoMen"
               Tab(2).Control(10)=   "pnl_Fin_MtoOto"
               Tab(2).Control(11)=   "pnl_Fin_TipMon"
               Tab(2).Control(12)=   "pnl_Fin_SalPag"
               Tab(2).Control(13)=   "pnl_Fin_NomIns"
               Tab(2).Control(14)=   "pnl_Fin_NumOpe"
               Tab(2).Control(15)=   "SSPanel28"
               Tab(2).Control(16)=   "SSPanel27"
               Tab(2).Control(17)=   "SSPanel21"
               Tab(2).Control(18)=   "SSPanel16"
               Tab(2).Control(19)=   "grd_Fin_Listad"
               Tab(2).ControlCount=   20
               TabCaption(3)   =   "Gastos Mensuales"
               TabPicture(3)   =   "EvaCre_frm_015.frx":00EC
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Label2"
               Tab(3).Control(1)=   "pnl_Gas_TotGas"
               Tab(3).Control(2)=   "SSPanel33"
               Tab(3).Control(3)=   "SSPanel32"
               Tab(3).Control(4)=   "grd_Gas_Listad"
               Tab(3).ControlCount=   5
               Begin Threed.SSPanel SSPanel8 
                  Height          =   60
                  Left            =   60
                  TabIndex        =   140
                  Top             =   2010
                  Width           =   10335
                  _Version        =   65536
                  _ExtentX        =   18230
                  _ExtentY        =   106
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
               Begin MSFlexGridLib.MSFlexGrid grd_Inm_Listad 
                  Height          =   1245
                  Left            =   90
                  TabIndex        =   141
                  Top             =   720
                  Width           =   10275
                  _ExtentX        =   18124
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   12
                  Cols            =   4
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel17 
                  Height          =   285
                  Left            =   6990
                  TabIndex        =   142
                  Top             =   420
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Fecha Adquiisición"
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
               Begin Threed.SSPanel SSPanel18 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   143
                  Top             =   420
                  Width           =   6885
                  _Version        =   65536
                  _ExtentX        =   12144
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo de Inmueble"
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
               Begin Threed.SSPanel SSPanel20 
                  Height          =   285
                  Left            =   8490
                  TabIndex        =   144
                  Top             =   420
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Importe Valorizado"
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
               Begin Threed.SSPanel pnl_Inm_TipInm 
                  Height          =   315
                  Left            =   1650
                  TabIndex        =   145
                  Top             =   2160
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Inm_FecAdq 
                  Height          =   315
                  Left            =   1650
                  TabIndex        =   146
                  Top             =   2820
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Inm_Direcc 
                  Height          =   315
                  Left            =   1650
                  TabIndex        =   147
                  Top             =   2490
                  Width           =   8655
                  _Version        =   65536
                  _ExtentX        =   15266
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Inm_Import 
                  Height          =   315
                  Left            =   1650
                  TabIndex        =   148
                  Top             =   3150
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Tar_Listad 
                  Height          =   1245
                  Left            =   -74910
                  TabIndex        =   153
                  Top             =   720
                  Width           =   10215
                  _ExtentX        =   18018
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   12
                  Cols            =   7
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel13 
                  Height          =   285
                  Left            =   -70920
                  TabIndex        =   154
                  Top             =   420
                  Width           =   2505
                  _Version        =   65536
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Nro. Tarjeta"
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
               Begin Threed.SSPanel SSPanel14 
                  Height          =   285
                  Left            =   -74880
                  TabIndex        =   155
                  Top             =   420
                  Width           =   3975
                  _Version        =   65536
                  _ExtentX        =   7011
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Institución Financiera"
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
               Begin Threed.SSPanel SSPanel15 
                  Height          =   285
                  Left            =   -68430
                  TabIndex        =   156
                  Top             =   420
                  Width           =   2055
                  _Version        =   65536
                  _ExtentX        =   3625
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Moneda"
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
               Begin Threed.SSPanel pnl_Tar_TipTar 
                  Height          =   315
                  Left            =   -73380
                  TabIndex        =   157
                  Top             =   2460
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Tar_NomIns 
                  Height          =   315
                  Left            =   -73380
                  TabIndex        =   158
                  Top             =   2130
                  Width           =   8655
                  _Version        =   65536
                  _ExtentX        =   15266
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Tar_SalPag 
                  Height          =   315
                  Left            =   -73380
                  TabIndex        =   159
                  Top             =   3120
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_Tar_NumTar 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   160
                  Top             =   2490
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Tar_TipMon 
                  Height          =   315
                  Left            =   -73380
                  TabIndex        =   161
                  Top             =   2790
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Tar_LinCre 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   162
                  Top             =   2820
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_Tar_MtoMin 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   163
                  Top             =   3150
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel SSPanel67 
                  Height          =   285
                  Left            =   -66390
                  TabIndex        =   164
                  Top             =   420
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Saldo x Pagar"
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
               Begin Threed.SSPanel SSPanel68 
                  Height          =   60
                  Left            =   -74970
                  TabIndex        =   165
                  Top             =   1980
                  Width           =   10335
                  _Version        =   65536
                  _ExtentX        =   18230
                  _ExtentY        =   106
                  _StockProps     =   15
                  BackColor       =   12632256
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
               Begin MSFlexGridLib.MSFlexGrid grd_Fin_Listad 
                  Height          =   1245
                  Left            =   -74910
                  TabIndex        =   173
                  Top             =   720
                  Width           =   10215
                  _ExtentX        =   18018
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   12
                  Cols            =   7
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel16 
                  Height          =   285
                  Left            =   -70920
                  TabIndex        =   174
                  Top             =   420
                  Width           =   2355
                  _Version        =   65536
                  _ExtentX        =   4154
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Nro. Operación"
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
               Begin Threed.SSPanel SSPanel21 
                  Height          =   285
                  Left            =   -74880
                  TabIndex        =   175
                  Top             =   420
                  Width           =   3975
                  _Version        =   65536
                  _ExtentX        =   7011
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Institución Financiera"
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
               Begin Threed.SSPanel SSPanel27 
                  Height          =   285
                  Left            =   -68580
                  TabIndex        =   176
                  Top             =   420
                  Width           =   2055
                  _Version        =   65536
                  _ExtentX        =   3625
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Moneda"
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
               Begin Threed.SSPanel SSPanel28 
                  Height          =   285
                  Left            =   -66540
                  TabIndex        =   177
                  Top             =   420
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Saldo x Pagar"
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
               Begin Threed.SSPanel pnl_Fin_NumOpe 
                  Height          =   315
                  Left            =   -73350
                  TabIndex        =   178
                  Top             =   2490
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Fin_NomIns 
                  Height          =   315
                  Left            =   -73350
                  TabIndex        =   179
                  Top             =   2160
                  Width           =   8655
                  _Version        =   65536
                  _ExtentX        =   15266
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Fin_SalPag 
                  Height          =   315
                  Left            =   -73350
                  TabIndex        =   180
                  Top             =   3150
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_Fin_TipMon 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   181
                  Top             =   2490
                  Width           =   3345
                  _Version        =   65536
                  _ExtentX        =   5900
                  _ExtentY        =   556
                  _StockProps     =   15
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
               Begin Threed.SSPanel pnl_Fin_MtoOto 
                  Height          =   315
                  Left            =   -73350
                  TabIndex        =   182
                  Top             =   2820
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_Fin_CuoMen 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   183
                  Top             =   2820
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel pnl_Fin_MesPag 
                  Height          =   315
                  Left            =   -68040
                  TabIndex        =   184
                  Top             =   3150
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin Threed.SSPanel SSPanel69 
                  Height          =   60
                  Left            =   -74940
                  TabIndex        =   185
                  Top             =   2010
                  Width           =   10335
                  _Version        =   65536
                  _ExtentX        =   18230
                  _ExtentY        =   106
                  _StockProps     =   15
                  BackColor       =   12632256
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
               Begin MSFlexGridLib.MSFlexGrid grd_Gas_Listad 
                  Height          =   2925
                  Left            =   -74910
                  TabIndex        =   193
                  Top             =   720
                  Width           =   10215
                  _ExtentX        =   18018
                  _ExtentY        =   5159
                  _Version        =   393216
                  Rows            =   12
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel32 
                  Height          =   285
                  Left            =   -74880
                  TabIndex        =   194
                  Top             =   420
                  Width           =   7785
                  _Version        =   65536
                  _ExtentX        =   13732
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo de Gasto"
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
               Begin Threed.SSPanel SSPanel33 
                  Height          =   285
                  Left            =   -67110
                  TabIndex        =   195
                  Top             =   420
                  Width           =   2025
                  _Version        =   65536
                  _ExtentX        =   3572
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Importe Valorizado S/."
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
               Begin Threed.SSPanel pnl_Gas_TotGas 
                  Height          =   315
                  Left            =   -67110
                  TabIndex        =   196
                  Top             =   3690
                  Width           =   2025
                  _Version        =   65536
                  _ExtentX        =   3572
                  _ExtentY        =   556
                  _StockProps     =   15
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
                  Alignment       =   4
               End
               Begin VB.Label Label2 
                  Caption         =   "Total ===>"
                  Height          =   315
                  Left            =   -68100
                  TabIndex        =   197
                  Top             =   3690
                  Width           =   855
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Instit. Financiera:"
                  Height          =   315
                  Index           =   64
                  Left            =   -74880
                  TabIndex        =   192
                  Top             =   2160
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Monto Otorgado:"
                  Height          =   315
                  Index           =   66
                  Left            =   -74880
                  TabIndex        =   191
                  Top             =   2820
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Nro. Operación:"
                  Height          =   315
                  Index           =   65
                  Left            =   -74880
                  TabIndex        =   190
                  Top             =   2490
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Saldo x Pagar:"
                  Height          =   315
                  Index           =   67
                  Left            =   -74880
                  TabIndex        =   189
                  Top             =   3150
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Moneda:"
                  Height          =   315
                  Index           =   70
                  Left            =   -69570
                  TabIndex        =   188
                  Top             =   2490
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Cuota Mensual:"
                  Height          =   315
                  Index           =   69
                  Left            =   -69570
                  TabIndex        =   187
                  Top             =   2820
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Meses x Pagar:"
                  Height          =   315
                  Index           =   68
                  Left            =   -69570
                  TabIndex        =   186
                  Top             =   3150
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Instit. Financiera:"
                  Height          =   315
                  Index           =   57
                  Left            =   -74880
                  TabIndex        =   172
                  Top             =   2160
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Moneda:"
                  Height          =   315
                  Index           =   59
                  Left            =   -74880
                  TabIndex        =   171
                  Top             =   2820
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Tarjeta:"
                  Height          =   315
                  Index           =   58
                  Left            =   -74880
                  TabIndex        =   170
                  Top             =   2490
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Saldo x Pagar:"
                  Height          =   315
                  Index           =   60
                  Left            =   -74880
                  TabIndex        =   169
                  Top             =   3150
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Número Tarjeta:"
                  Height          =   315
                  Index           =   61
                  Left            =   -69570
                  TabIndex        =   168
                  Top             =   2490
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Línea Crédito:"
                  Height          =   315
                  Index           =   62
                  Left            =   -69570
                  TabIndex        =   167
                  Top             =   2820
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Monto Mínimo:"
                  Height          =   315
                  Index           =   63
                  Left            =   -69570
                  TabIndex        =   166
                  Top             =   3150
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Tipo Inmueble:"
                  Height          =   315
                  Index           =   53
                  Left            =   120
                  TabIndex        =   152
                  Top             =   2220
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Fecha Adquisic.:"
                  Height          =   315
                  Index           =   55
                  Left            =   120
                  TabIndex        =   151
                  Top             =   2880
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Dirección:"
                  Height          =   315
                  Index           =   54
                  Left            =   120
                  TabIndex        =   150
                  Top             =   2550
                  Width           =   1395
               End
               Begin VB.Label lbl_Etique 
                  Caption         =   "Importe:"
                  Height          =   315
                  Index           =   56
                  Left            =   120
                  TabIndex        =   149
                  Top             =   3210
                  Width           =   1395
               End
            End
            Begin VB.TextBox txt_Sol_Observ 
               Height          =   885
               Left            =   -73350
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   99
               Text            =   "EvaCre_frm_015.frx":0108
               Top             =   3720
               Width           =   8925
            End
            Begin Threed.SSPanel pnl_Fam_TipPar 
               Height          =   315
               Left            =   -73350
               TabIndex        =   65
               Top             =   720
               Width           =   3345
               _Version        =   65536
               _ExtentX        =   5900
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Fam_Nombre 
               Height          =   315
               Left            =   -73350
               TabIndex        =   66
               Top             =   1050
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Fam_Telefo 
               Height          =   315
               Left            =   -73350
               TabIndex        =   67
               Top             =   1380
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Fam_Celula 
               Height          =   315
               Left            =   -73350
               TabIndex        =   68
               Top             =   1710
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NFa_TipPar 
               Height          =   315
               Left            =   -73350
               TabIndex        =   69
               Top             =   2430
               Width           =   3345
               _Version        =   65536
               _ExtentX        =   5900
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NFa_Nombre 
               Height          =   315
               Left            =   -73350
               TabIndex        =   70
               Top             =   2760
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NFa_Telefo 
               Height          =   315
               Left            =   -73350
               TabIndex        =   71
               Top             =   3090
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NFa_Celula 
               Height          =   315
               Left            =   -73350
               TabIndex        =   72
               Top             =   3420
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_TipPro 
               Height          =   315
               Left            =   -73320
               TabIndex        =   83
               Top             =   990
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NatTit 
               Height          =   315
               Left            =   -73320
               TabIndex        =   84
               Top             =   1770
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_NatCyg 
               Height          =   315
               Left            =   -73320
               TabIndex        =   85
               Top             =   2100
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_JurEmp 
               Height          =   315
               Left            =   -73320
               TabIndex        =   86
               Top             =   2850
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_JurDir 
               Height          =   615
               Left            =   -73320
               TabIndex        =   87
               Top             =   3180
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   1085
               _StockProps     =   15
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_JurRep 
               Height          =   315
               Left            =   -73320
               TabIndex        =   88
               Top             =   3810
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Direcc 
               Height          =   555
               Left            =   -73320
               TabIndex        =   97
               Top             =   420
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Sol_FecIng 
               Height          =   315
               Left            =   -73350
               TabIndex        =   100
               Top             =   1080
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Sol_NumCuo 
               Height          =   315
               Left            =   -73350
               TabIndex        =   101
               Top             =   1410
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_PerGra 
               Height          =   315
               Left            =   -73350
               TabIndex        =   102
               Top             =   1740
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_CuoExt 
               Height          =   315
               Left            =   -73350
               TabIndex        =   103
               Top             =   2070
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Sol_ComVta_Dol 
               Height          =   315
               Left            =   -73350
               TabIndex        =   104
               Top             =   2400
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "30,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_ApoPro_Dol 
               Height          =   315
               Left            =   -73350
               TabIndex        =   105
               Top             =   2730
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_MtoSol_Dol 
               Height          =   315
               Left            =   -73350
               TabIndex        =   106
               Top             =   3060
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_ComVta_MPr 
               Height          =   315
               Left            =   -69480
               TabIndex        =   107
               Top             =   2400
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "30,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_ApoPro_MPr 
               Height          =   315
               Left            =   -69480
               TabIndex        =   108
               Top             =   2730
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_MtoSol_MPr 
               Height          =   315
               Left            =   -69480
               TabIndex        =   109
               Top             =   3060
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_ComVta_Sol 
               Height          =   315
               Left            =   -65580
               TabIndex        =   110
               Top             =   2400
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "30,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_ApoPro_Sol 
               Height          =   315
               Left            =   -65580
               TabIndex        =   111
               Top             =   2730
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_MtoSol_Sol 
               Height          =   315
               Left            =   -65580
               TabIndex        =   112
               Top             =   3060
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_TcaDol 
               Height          =   315
               Left            =   -65580
               TabIndex        =   113
               Top             =   1080
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_TcaMPr 
               Height          =   315
               Left            =   -65580
               TabIndex        =   114
               Top             =   1410
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_EjeVta 
               Height          =   315
               Left            =   -73350
               TabIndex        =   115
               Top             =   420
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin Threed.SSPanel pnl_Sol_CuoSeg 
               Height          =   315
               Left            =   -73350
               TabIndex        =   116
               Top             =   3390
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_VctSug 
               Height          =   315
               Left            =   -69480
               TabIndex        =   117
               Top             =   3390
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Sol_TipSeg 
               Height          =   315
               Left            =   -73350
               TabIndex        =   118
               Top             =   750
               Width           =   8925
               _Version        =   65536
               _ExtentX        =   15743
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin MSFlexGridLib.MSFlexGrid grd_Doc_Listad 
               Height          =   3885
               Left            =   -74910
               TabIndex        =   199
               Top             =   720
               Width           =   10455
               _ExtentX        =   18441
               _ExtentY        =   6853
               _Version        =   393216
               Rows            =   12
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel24 
               Height          =   285
               Left            =   -74880
               TabIndex        =   200
               Top             =   420
               Width           =   10185
               _Version        =   65536
               _ExtentX        =   17965
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Documento"
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
            Begin VB.Label lbl_Etique 
               Caption         =   "F. Ingreso Solic.:"
               Height          =   315
               Index           =   52
               Left            =   -74880
               TabIndex        =   138
               Top             =   1080
               Width           =   1185
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Período de Gracia:"
               Height          =   315
               Index           =   51
               Left            =   -74880
               TabIndex        =   137
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Plazo (Meses):"
               Height          =   315
               Index           =   50
               Left            =   -74880
               TabIndex        =   136
               Top             =   1410
               Width           =   1275
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Cuotas Extraord.:"
               Height          =   315
               Index           =   49
               Left            =   -74880
               TabIndex        =   135
               Top             =   2070
               Width           =   1305
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Monto Solicit. S/.:"
               Height          =   315
               Index           =   48
               Left            =   -67170
               TabIndex        =   134
               Top             =   3060
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Aporte Propio S/.:"
               Height          =   315
               Index           =   47
               Left            =   -67170
               TabIndex        =   133
               Top             =   2730
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Valor Inmueb. S/.:"
               Height          =   315
               Index           =   46
               Left            =   -67170
               TabIndex        =   132
               Top             =   2400
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Monto Solicit. MPr:"
               Height          =   315
               Index           =   45
               Left            =   -71040
               TabIndex        =   131
               Top             =   3060
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Aporte Propio MPr:"
               Height          =   315
               Index           =   44
               Left            =   -71040
               TabIndex        =   130
               Top             =   2730
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Valor Inmueb. MPr.:"
               Height          =   315
               Index           =   43
               Left            =   -71040
               TabIndex        =   129
               Top             =   2400
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Monto Solicit. US$:"
               Height          =   315
               Index           =   42
               Left            =   -74880
               TabIndex        =   128
               Top             =   3060
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Aporte Propio US$:"
               Height          =   315
               Index           =   41
               Left            =   -74880
               TabIndex        =   127
               Top             =   2730
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Valor Inmueb. US$:"
               Height          =   315
               Index           =   40
               Left            =   -74880
               TabIndex        =   126
               Top             =   2400
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "T. Cambio US$:"
               Height          =   315
               Index           =   39
               Left            =   -67170
               TabIndex        =   125
               Top             =   1080
               Width           =   1305
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "T. Cambio M. Prest.:"
               Height          =   315
               Index           =   38
               Left            =   -67170
               TabIndex        =   124
               Top             =   1410
               Width           =   1485
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Ejec. Ventas:"
               Height          =   315
               Index           =   37
               Left            =   -74880
               TabIndex        =   123
               Top             =   420
               Width           =   945
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Cuota Mens. Sug.:"
               Height          =   315
               Index           =   36
               Left            =   -74880
               TabIndex        =   122
               Top             =   3390
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Observaciones:"
               Height          =   315
               Index           =   35
               Left            =   -74880
               TabIndex        =   121
               Top             =   3720
               Width           =   1335
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Día Vcto. Sug.:"
               Height          =   315
               Index           =   34
               Left            =   -71040
               TabIndex        =   120
               Top             =   3390
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Tipo de Seguro:"
               Height          =   315
               Index           =   33
               Left            =   -74880
               TabIndex        =   119
               Top             =   750
               Width           =   1365
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Dirección Inmueble:"
               Height          =   405
               Index           =   32
               Left            =   -74880
               TabIndex        =   98
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Tipo Propietario:"
               Height          =   315
               Index           =   31
               Left            =   -74880
               TabIndex        =   96
               Top             =   990
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Persona Natural"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   30
               Left            =   -74880
               TabIndex        =   95
               Top             =   1470
               Width           =   1815
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Cónyuge:"
               Height          =   315
               Index           =   29
               Left            =   -74880
               TabIndex        =   94
               Top             =   2100
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Titular:"
               Height          =   315
               Index           =   28
               Left            =   -74880
               TabIndex        =   93
               Top             =   1770
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Rep. Legal:"
               Height          =   315
               Index           =   27
               Left            =   -74880
               TabIndex        =   92
               Top             =   3810
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Dirección Empresa:"
               Height          =   285
               Index           =   26
               Left            =   -74880
               TabIndex        =   91
               Top             =   3180
               Width           =   1485
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Empresa:"
               Height          =   315
               Index           =   25
               Left            =   -74880
               TabIndex        =   90
               Top             =   2850
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Persona Jurídica"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   24
               Left            =   -74880
               TabIndex        =   89
               Top             =   2550
               Width           =   1815
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Tipo Parentesco:"
               Height          =   315
               Index           =   23
               Left            =   -74880
               TabIndex        =   82
               Top             =   720
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Nombre Referencia:"
               Height          =   315
               Index           =   22
               Left            =   -74880
               TabIndex        =   81
               Top             =   1050
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Referencia Familiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   21
               Left            =   -74880
               TabIndex        =   80
               Top             =   420
               Width           =   1815
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Teléfono:"
               Height          =   315
               Index           =   20
               Left            =   -74880
               TabIndex        =   79
               Top             =   1380
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Celular:"
               Height          =   315
               Index           =   19
               Left            =   -74880
               TabIndex        =   78
               Top             =   1710
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Tipo Parentesco:"
               Height          =   315
               Index           =   18
               Left            =   -74880
               TabIndex        =   77
               Top             =   2430
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Nombre Referencia:"
               Height          =   315
               Index           =   17
               Left            =   -74880
               TabIndex        =   76
               Top             =   2760
               Width           =   1425
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Referencia No Familiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   16
               Left            =   -74880
               TabIndex        =   75
               Top             =   2130
               Width           =   2235
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Teléfono:"
               Height          =   315
               Index           =   15
               Left            =   -74880
               TabIndex        =   74
               Top             =   3090
               Width           =   1395
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Celular:"
               Height          =   315
               Index           =   14
               Left            =   -74880
               TabIndex        =   73
               Top             =   3420
               Width           =   1395
            End
         End
      End
      Begin Threed.SSPanel SSPanel42 
         Height          =   765
         Left            =   11010
         TabIndex        =   56
         Top             =   9750
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   2550
            Picture         =   "EvaCre_frm_015.frx":010C
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5655
         Left            =   11010
         TabIndex        =   28
         Top             =   2670
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   9975
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
         Begin VB.CommandButton cmd_Calcul 
            Caption         =   "Calcular"
            Height          =   285
            Left            =   2010
            TabIndex        =   53
            Top             =   3060
            Width           =   1155
         End
         Begin Threed.SSPanel pnl_Eva_TCaDol 
            Height          =   315
            Left            =   2010
            TabIndex        =   29
            Top             =   90
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_TCaMPr 
            Height          =   315
            Left            =   2010
            TabIndex        =   30
            Top             =   420
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_SegDes 
            Height          =   315
            Left            =   2010
            TabIndex        =   31
            Top             =   900
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_SegViv 
            Height          =   315
            Left            =   2010
            TabIndex        =   32
            Top             =   1230
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_TasInt 
            Height          =   315
            Left            =   2010
            TabIndex        =   37
            Top             =   1560
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin EditLib.fpLongInteger ipp_Eva_NumCuo 
            Height          =   315
            Left            =   2010
            TabIndex        =   39
            Top             =   2370
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
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
            MaxValue        =   "240"
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
         Begin EditLib.fpDoubleSingle ipp_Eva_MtoCre 
            Height          =   315
            Left            =   2010
            TabIndex        =   40
            Top             =   2700
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
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
            ButtonStyle     =   0
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin Threed.SSPanel pnl_Eva_CuoMen_Dol 
            Height          =   315
            Left            =   1980
            TabIndex        =   44
            Top             =   3480
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_CuoMen_Sol 
            Height          =   315
            Left            =   1980
            TabIndex        =   46
            Top             =   3810
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_CuoMen_MPr 
            Height          =   315
            Left            =   1980
            TabIndex        =   48
            Top             =   4140
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel34 
            Height          =   60
            Left            =   60
            TabIndex        =   50
            Top             =   780
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   106
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
         Begin Threed.SSPanel SSPanel35 
            Height          =   60
            Left            =   60
            TabIndex        =   51
            Top             =   1920
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   106
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
         Begin Threed.SSPanel SSPanel39 
            Height          =   60
            Left            =   60
            TabIndex        =   52
            Top             =   3390
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   106
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
         Begin Threed.SSPanel pnl_Eva_CuoRta 
            Height          =   315
            Left            =   1980
            TabIndex        =   54
            Top             =   5280
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel43 
            Height          =   60
            Left            =   60
            TabIndex        =   58
            Top             =   5160
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   106
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
         Begin Threed.SSPanel pnl_Eva_PriCuo 
            Height          =   315
            Left            =   1980
            TabIndex        =   59
            Top             =   4470
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Eva_UltCuo 
            Height          =   315
            Left            =   1980
            TabIndex        =   61
            Top             =   4800
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   4
         End
         Begin EditLib.fpDoubleSingle ipp_Eva_TotILD 
            Height          =   315
            Left            =   2010
            TabIndex        =   198
            Top             =   2040
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
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
            ButtonStyle     =   0
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin VB.Label lbl_Etique 
            Caption         =   "Ult. Cuota US$:"
            Height          =   315
            Index           =   13
            Left            =   90
            TabIndex        =   62
            Top             =   4800
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "1ra Cuota US$:"
            Height          =   315
            Index           =   12
            Left            =   90
            TabIndex        =   60
            Top             =   4470
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Relación Cuota/Renta:"
            Height          =   315
            Index           =   11
            Left            =   90
            TabIndex        =   55
            Top             =   5280
            Width           =   1695
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cuota Mensual MPr.:"
            Height          =   315
            Index           =   10
            Left            =   90
            TabIndex        =   49
            Top             =   4140
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cuota Mensual S/.:"
            Height          =   315
            Index           =   9
            Left            =   90
            TabIndex        =   47
            Top             =   3810
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cuota Mensual US$:"
            Height          =   315
            Index           =   8
            Left            =   90
            TabIndex        =   45
            Top             =   3480
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Total ILD S/.:"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Plazo (Meses):"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Top             =   2370
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Monto Crédito US$:"
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   41
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "T. Interes:"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   1560
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "T. Cambio US$:"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   90
            Width           =   1305
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "T. Cambio M. Prest.:"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "T. Seguro Desg.:"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   900
            Width           =   1845
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "T. Seguro Vivienda:"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   1230
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3015
         Left            =   30
         TabIndex        =   3
         Top             =   7500
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
         BevelOuter      =   1
         Begin TabDlg.SSTab tab_ActEco 
            Height          =   2895
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   10755
            _ExtentX        =   18971
            _ExtentY        =   5106
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Tit. (Act. Econ. Princ.)"
            TabPicture(0)   =   "EvaCre_frm_015.frx":054E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label12"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbl_Etique(78)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "grd_Tit_Listad(0)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_Tit_Ocupac(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Tit. (Act. Econ. Sec.)"
            TabPicture(1)   =   "EvaCre_frm_015.frx":056A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_Tit_Ocupac(1)"
            Tab(1).Control(1)=   "grd_Tit_Listad(1)"
            Tab(1).Control(2)=   "Label3"
            Tab(1).Control(3)=   "lbl_Etique(79)"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Cyg. (Act. Econ. Princ.)"
            TabPicture(2)   =   "EvaCre_frm_015.frx":0586
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "pnl_Cyg_Ocupac(0)"
            Tab(2).Control(1)=   "grd_Cyg_Listad(0)"
            Tab(2).Control(2)=   "Label5"
            Tab(2).Control(3)=   "lbl_Etique(80)"
            Tab(2).ControlCount=   4
            TabCaption(3)   =   "Cyg. (Act. Econ. Sec.)"
            TabPicture(3)   =   "EvaCre_frm_015.frx":05A2
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "pnl_Cyg_Ocupac(1)"
            Tab(3).Control(1)=   "grd_Cyg_Listad(1)"
            Tab(3).Control(2)=   "Label7"
            Tab(3).Control(3)=   "lbl_Etique(81)"
            Tab(3).ControlCount=   4
            Begin Threed.SSPanel pnl_Tit_Ocupac 
               Height          =   315
               Index           =   0
               Left            =   1560
               TabIndex        =   18
               Top             =   360
               Width           =   4095
               _Version        =   65536
               _ExtentX        =   7223
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin MSFlexGridLib.MSFlexGrid grd_Tit_Listad 
               Height          =   2145
               Index           =   0
               Left            =   1530
               TabIndex        =   19
               Top             =   690
               Width           =   9165
               _ExtentX        =   16166
               _ExtentY        =   3784
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Tit_Ocupac 
               Height          =   315
               Index           =   1
               Left            =   -73440
               TabIndex        =   202
               Top             =   360
               Width           =   4095
               _Version        =   65536
               _ExtentX        =   7223
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin MSFlexGridLib.MSFlexGrid grd_Tit_Listad 
               Height          =   2145
               Index           =   1
               Left            =   -73470
               TabIndex        =   203
               Top             =   690
               Width           =   9165
               _ExtentX        =   16166
               _ExtentY        =   3784
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Cyg_Ocupac 
               Height          =   315
               Index           =   0
               Left            =   -73440
               TabIndex        =   206
               Top             =   360
               Width           =   4095
               _Version        =   65536
               _ExtentX        =   7223
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin MSFlexGridLib.MSFlexGrid grd_Cyg_Listad 
               Height          =   2145
               Index           =   0
               Left            =   -73470
               TabIndex        =   207
               Top             =   690
               Width           =   9165
               _ExtentX        =   16166
               _ExtentY        =   3784
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Cyg_Ocupac 
               Height          =   315
               Index           =   1
               Left            =   -73440
               TabIndex        =   210
               Top             =   360
               Width           =   4095
               _Version        =   65536
               _ExtentX        =   7223
               _ExtentY        =   556
               _StockProps     =   15
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
            Begin MSFlexGridLib.MSFlexGrid grd_Cyg_Listad 
               Height          =   2145
               Index           =   1
               Left            =   -73470
               TabIndex        =   211
               Top             =   690
               Width           =   9165
               _ExtentX        =   16166
               _ExtentY        =   3784
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label7 
               Caption         =   "Datos Actividad:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   213
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Ocupación:"
               Height          =   285
               Index           =   81
               Left            =   -74910
               TabIndex        =   212
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label5 
               Caption         =   "Datos Actividad:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   209
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Ocupación:"
               Height          =   285
               Index           =   80
               Left            =   -74910
               TabIndex        =   208
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label3 
               Caption         =   "Datos Actividad:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   205
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Ocupación:"
               Height          =   285
               Index           =   79
               Left            =   -74910
               TabIndex        =   204
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label lbl_Etique 
               Caption         =   "Ocupación:"
               Height          =   285
               Index           =   78
               Left            =   90
               TabIndex        =   21
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label12 
               Caption         =   "Datos Actividad:"
               Height          =   285
               Left            =   90
               TabIndex        =   20
               Top             =   690
               Width           =   1275
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14235
         _Version        =   65536
         _ExtentX        =   25109
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
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Créditos"
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
            Picture         =   "EvaCre_frm_015.frx":05BE
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   735
         Left            =   30
         TabIndex        =   4
         Top             =   750
         Width           =   14235
         _Version        =   65536
         _ExtentX        =   25109
         _ExtentY        =   1296
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
            Height          =   675
            Left            =   13530
            Picture         =   "EvaCre_frm_015.frx":08C8
            Style           =   1  'Graphical
            TabIndex        =   201
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Observ 
            Height          =   675
            Left            =   2430
            Picture         =   "EvaCre_frm_015.frx":0D0A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Observaciones"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   12150
            Picture         =   "EvaCre_frm_015.frx":114C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12840
            Picture         =   "EvaCre_frm_015.frx":1456
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   8
         Top             =   1530
         Width           =   14235
         _Version        =   65536
         _ExtentX        =   25109
         _ExtentY        =   1931
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
            Left            =   1620
            TabIndex        =   9
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   10080
            TabIndex        =   10
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   4740
            TabIndex        =   11
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   10080
            TabIndex        =   12
            Top             =   720
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_FecEnv 
            Height          =   315
            Left            =   10080
            TabIndex        =   22
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_CliTit 
            Height          =   315
            Left            =   1620
            TabIndex        =   24
            Top             =   390
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_CliCyg 
            Height          =   315
            Left            =   1620
            TabIndex        =   26
            Top             =   720
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin VB.Label lbl_Etique 
            Caption         =   "Cliente Cónyuge:"
            Height          =   315
            Index           =   73
            Left            =   60
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cliente Titular:"
            Height          =   315
            Index           =   74
            Left            =   60
            TabIndex        =   25
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "F. Envío:"
            Height          =   315
            Index           =   76
            Left            =   9210
            TabIndex        =   23
            Top             =   390
            Width           =   795
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Moneda:"
            Height          =   315
            Index           =   77
            Left            =   9210
            TabIndex        =   16
            Top             =   720
            Width           =   765
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Producto:"
            Height          =   315
            Index           =   72
            Left            =   3900
            TabIndex        =   15
            Top             =   60
            Width           =   825
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Modalidad:"
            Height          =   315
            Index           =   71
            Left            =   9210
            TabIndex        =   14
            Top             =   60
            Width           =   855
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Index           =   75
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_EvaCre_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CliNCo()   As modcal_g_est_CuoCli
Dim l_int_CuoExt     As Integer
Dim l_int_PerGra     As Integer
Dim l_int_AplDes     As Integer
Dim l_int_AplViv     As Integer
Dim l_int_FlgCal     As Integer
Dim l_str_IniEva     As String

Private Sub cmd_Aprueb_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   If l_int_FlgCal = 0 Then
      MsgBox "Debe efectuar un cálculo con los datos del día de hoy para proceder a la Aprobación de la Evaluación.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(cmd_Calcul)
      Exit Sub
   End If
   
   If ff_Buscar_LisObs() Then
      MsgBox "La Solicitud presenta Observaciones pendientes de Subsanación.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(cmd_Observ)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_AteCli) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaCre, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaCre, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando Datos de Aprobación en Solicitud de Crédito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_SOLMAE_APRUEBA ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Format(CDbl(ipp_Eva_MtoCre.Value * CDbl(pnl_Eva_TCaDol.Caption) / CDbl(pnl_Eva_TCaMPr.Caption)), "#########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(CDbl(ipp_Eva_MtoCre.Value * CDbl(pnl_Eva_TCaDol.Caption)), "#########0.00") & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Eva_MtoCre.Value)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_TCaDol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_TCaMPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_Eva_NumCuo.Value)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_CuoMen_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Eva_TotILD.Value)) & ", "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_CuoRta.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_PriCuo.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_UltCuo.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra) & ", "
      
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
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_AceIni) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_AceIni, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, modatecli_g_con_AceIni) Then
      Exit Sub
   End If
   
   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "APROBACION DE EVALUACION CREDITICIA INICIAL (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
   
   moddat_g_int_FlgAct = 1
   
   Unload Me
End Sub

Private Sub cmd_Calcul_Click()
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_SegPre     As Double
   Dim r_dbl_SegViv     As Double
   
   If ipp_Eva_TotILD.Value = 0 Then
      MsgBox "Debe ingresar el Total de ILD.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(ipp_Eva_TotILD)
      Exit Sub
   End If
   
   If ipp_Eva_NumCuo.Value = 0 Then
      MsgBox "Debe ingresar el Número de Cuotas.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(ipp_Eva_NumCuo)
      Exit Sub
   End If
   
   If ipp_Eva_MtoCre.Value = 0 Then
      MsgBox "Debe ingresar el Monto de Crédito.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(ipp_Eva_MtoCre)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Calculando Seguro de Préstamo
   If l_int_AplDes = 1 Then           'Factor
      r_dbl_SegPre = CDbl(Format(CDbl(pnl_Eva_SegDes.Caption) * CDbl(ipp_Eva_MtoCre.Text), "###,###,##0.00"))
   Else                                   'Importe
      r_dbl_SegPre = CDbl(Format(CDbl(pnl_Eva_SegDes.Caption), "###,###,##0.00"))
   End If
   
   'Calculando Seguro de Vivienda
   If l_int_AplViv = 1 Then           'Factor
      r_dbl_SegViv = CDbl(Format(CDbl(pnl_Eva_SegViv.Caption) * CDbl(pnl_Sol_ComVta_Dol), "###,###,##0.00"))
   Else                                   'Importe
      r_dbl_SegViv = CDbl(Format(CDbl(pnl_Eva_SegViv.Caption), "###,###,##0.00"))
   End If
   
   'Obteniendo Valor de Otros Cargos
   Call moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "104", Format(moddat_g_int_TipMon, "000"))
   r_dbl_OtrCar = moddat_g_arr_Genera(1).Genera_Cantid
   
   'Generando Cronograma Cliente No Concesional
   Call gs_Calcul_Cliente(l_arr_CliNCo(), CDbl(pnl_Eva_TasInt.Caption), l_int_CuoExt, Format(Date, "dd/mm/yyyy"), ipp_Eva_NumCuo.Value, CDbl(ipp_Eva_MtoCre.Text), 1, l_int_PerGra, "")
   
   pnl_Eva_CuoMen_Dol.Caption = Format(l_arr_CliNCo(2).CuoCli_ValCuo + r_dbl_SegPre + r_dbl_SegViv + r_dbl_OtrCar, "###,###,0.00") & " "
   pnl_Eva_PriCuo.Caption = Format(l_arr_CliNCo(1).CuoCli_ValCuo + r_dbl_SegPre + r_dbl_SegViv + r_dbl_OtrCar, "###,###,0.00") & " "
   pnl_Eva_UltCuo.Caption = Format(l_arr_CliNCo(ipp_Eva_NumCuo.Value).CuoCli_ValCuo + r_dbl_SegPre + r_dbl_SegViv + r_dbl_OtrCar, "###,###,0.00") & " "
   
   pnl_Eva_CuoMen_Sol.Caption = Format(pnl_Eva_CuoMen_Dol.Caption * CDbl(pnl_Eva_TCaDol.Caption), "###,###,0.00") & " "
   pnl_Eva_CuoMen_MPr.Caption = Format(pnl_Eva_CuoMen_Dol.Caption * CDbl(pnl_Eva_TCaDol.Caption) / CDbl(pnl_Eva_TCaMPr.Caption), "###,###,0.00") & " "
   pnl_Eva_CuoRta.Caption = Format(CDbl(pnl_Eva_CuoMen_Sol.Caption) / CDbl(ipp_Eva_TotILD.Text) * 100, "##0.00") & " "
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_RcrMax     As Double

   If CDbl(pnl_Eva_CuoMen_Dol.Caption) = 0 Then
      MsgBox "Debe calcular los valores de Cuota.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(cmd_Calcul)
      Exit Sub
   End If
   
   'Buscando Parámetro de RCR Máximo
   If moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "001", "001") Then
      r_dbl_RcrMax = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   If CDbl(pnl_Eva_CuoRta.Caption) > r_dbl_RcrMax Then
      MsgBox "Porcentaje de Relación Cuota/Renta excede el Parámetro en Políticas.", vbExclamation, modgen_g_con_EvaCre
      Call gs_SetFocus(cmd_Calcul)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de registrar la Evaluación?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVACRE ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & Format(CDbl(ipp_Eva_MtoCre.Value * CDbl(pnl_Eva_TCaDol.Caption) / CDbl(pnl_Eva_TCaMPr.Caption)), "#########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(CDbl(ipp_Eva_MtoCre.Value * CDbl(pnl_Eva_TCaDol.Caption)), "#########0.00") & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Eva_MtoCre.Value)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_TCaDol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_TCaMPr.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(ipp_Eva_NumCuo.Value)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_CuoMen_Dol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Eva_TotILD.Value)) & ", "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_CuoRta.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_PriCuo.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Eva_UltCuo.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_int_PerGra) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Eva_Observ.Text & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                        'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
         
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
   Screen.MousePointer = 0
   
   MsgBox "Evaluación registrada correctamente.", vbInformation, modgen_g_str_NomPlt
   
   l_int_FlgCal = 1
End Sub

Private Sub cmd_Observ_Click()
   frm_EvaCre_04.Show 1
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = modatecli_g_con_EvaCre
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaCre, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaCre, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
   
      r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      r_str_Cadena = r_str_Cadena & Chr(13)
   
   
      modgen_g_str_Mail_Asunto = "RECHAZO DE EVALUACION CREDITICIA INICIAL (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      
      frm_EnvMai_01.Show 1
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
      
      moddat_g_int_FlgAct = 1
      
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Command5_Click()
   
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call fs_Inicia
   Call fs_Limpia
   
   Me.Caption = modgen_g_con_EvaCre
   
   Call fs_Carga_SolInb
   Call fs_Carga_SolTrj
   Call fs_Carga_SolDeu
   Call fs_Carga_SolGas
   
   Call fs_Carga_Refere
   Call fs_Carga_DatInm
   Call fs_Carga_DatCre
   Call fs_Carga_SolDoc
   
   tab_Princi.Tab = 0
   tab_InfFin.Tab = 0
   tab_ActEco.Tab = 0
   
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1, grd_Tit_Listad(0), pnl_Tit_Ocupac(0))
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2, grd_Tit_Listad(1), pnl_Tit_Ocupac(1))
   
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1, grd_Cyg_Listad(0), pnl_Cyg_Ocupac(0))
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2, grd_Cyg_Listad(1), pnl_Cyg_Ocupac(1))
   
   'Buscar Información de Cálculos Anteriores
   Call fs_Carga_EvaCre
   
   Call fs_Buscar_SegDet
   
   l_int_FlgCal = 0
   
   Call gs_CentraForm(Me)
   
   If CDbl(pnl_Eva_TCaDol.Caption) = 0 Or CDbl(pnl_Eva_TCaMPr.Caption) = 0 Then
      cmd_Observ.Enabled = False
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
      cmd_Calcul.Enabled = False
      cmd_Grabar.Enabled = False
      
      ipp_Eva_TotILD.Enabled = False
      ipp_Eva_NumCuo.Enabled = False
      ipp_Eva_MtoCre.Enabled = False
      txt_Eva_Observ.Enabled = False
      
      MsgBox "No se encuentra registrado el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   pnl_Eva_TCaDol.Caption = "0.00 "
   pnl_Eva_TCaMPr.Caption = "0.00 "
   
   pnl_Eva_SegDes.Caption = "0.00 "
   pnl_Eva_SegViv.Caption = "0.00 "
   pnl_Eva_TasInt.Caption = "0.00 "
   
   ipp_Eva_TotILD.Value = 0
   ipp_Eva_NumCuo.Value = 0
   ipp_Eva_MtoCre.Value = 0
   
   pnl_Eva_CuoMen_Dol.Caption = "0.00 "
   pnl_Eva_CuoMen_Sol.Caption = "0.00 "
   pnl_Eva_CuoMen_MPr.Caption = "0.00 "
   
   pnl_Eva_CuoRta.Caption = "0.00 "
   pnl_Eva_PriCuo.Caption = "0.00 "
   pnl_Eva_UltCuo.Caption = "0.00 "

   txt_Eva_Observ.Text = ""
End Sub

Private Sub fs_Carga_DatInm()
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON))

   pnl_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMERO)
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_INTDPT))) > 0 Then
      pnl_Direcc.Caption = pnl_Direcc.Caption & " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_NOMZON))) > 0 Then
      pnl_Direcc.Caption = pnl_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_NOMZON) & Chr(13) & Chr(10)
   Else
      pnl_Direcc.Caption = pnl_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
   
   pnl_Direcc.Caption = pnl_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   pnl_TipPro.Caption = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!SOLINM_TIPPER))
   
   If g_rst_Princi!SOLINM_TIPPER = 2 Then
      'Persona Jurídica
      
      pnl_JurEmp.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PRORZS)
      pnl_JurRep.Caption = Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_PROTVI))
      r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_PROTZO))
   
      pnl_JurDir.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_PRONVI) & " " & Trim(g_rst_Princi!SOLINM_PRONUM)
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PROINT))) > 0 Then
         pnl_JurDir.Caption = pnl_JurDir.Caption & " (" & Trim(g_rst_Princi!SOLINM_PROINT) & ")"
      End If
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PRONZO))) > 0 Then
         pnl_JurDir.Caption = pnl_JurDir.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_PRONZO) & Chr(13) & Chr(10)
      Else
         pnl_JurDir.Caption = pnl_JurDir.Caption & Chr(13) & Chr(10)
      End If
      
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 2) & "0000")
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 4) & "00")
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_PROUBI))
      
      pnl_JurDir.Caption = pnl_JurDir.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   Else
      'Persona Natural
      
      pnl_NatTit.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      If g_rst_Princi!SOLINM_CYGTDO > 0 Then
         pnl_NatCyg.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_CYGTDO)) & "-" & Trim(g_rst_Princi!SOLINM_CYGNDO) & " / " & Trim(g_rst_Princi!SOLINM_CYGAPP) & " " & Trim(g_rst_Princi!SOLINM_CYGAPM) & " " & Trim(g_rst_Princi!SOLINM_CYGNOM)
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_DatCre()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   
   'Obteniendo Nombre de Cliente
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNom)
   Else
      moddat_g_int_CygTDo = 0
      moddat_g_str_CygNDo = ""
      moddat_g_str_CygNom = ""
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
      
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)
      
   
   pnl_CliTit.Caption = moddat_g_str_NomCli
   pnl_CliCyg.Caption = moddat_g_str_CygNom
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecEnv.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECENV))
   
   pnl_Sol_EjeVta.Caption = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJEVTA))
   pnl_Sol_TipSeg.Caption = Trim(moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG))

   pnl_Sol_FecIng.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   pnl_Sol_NumCuo.Caption = Format(g_rst_Princi!SOLMAE_PLAPRE, "##0") & " "
   pnl_Sol_PerGra.Caption = Format(g_rst_Princi!SOLMAE_PERGRA, "##0") & " "
   pnl_Sol_CuoExt.Caption = moddat_gf_Consulta_ParDes("223", g_rst_Princi!SOLMAE_CUOANO) & " "
   
   pnl_Sol_TcaDol.Caption = Format(g_rst_Princi!SOLMAE_TCADOL, "###,##0.000000") & " "
   pnl_Sol_TcaMPr.Caption = Format(g_rst_Princi!SOLMAE_TCAMPR, "###,##0.000000") & " "
   
   pnl_Sol_ComVta_Dol.Caption = Format(g_rst_Princi!SOLMAE_COMVTA, "###,###,##0.00") & " "
   pnl_Sol_ApoPro_Dol.Caption = Format(g_rst_Princi!SOLMAE_APOPRO, "###,###,##0.00") & " "
   pnl_Sol_MtoSol_Dol.Caption = Format(g_rst_Princi!SOLMAE_MTOSOL, "###,###,##0.00") & " "
   
   pnl_Sol_ComVta_Sol.Caption = Format(g_rst_Princi!SOLMAE_COMVTA * g_rst_Princi!SOLMAE_TCADOL, "###,###,##0.00" & " ")
   pnl_Sol_ApoPro_Sol.Caption = Format(g_rst_Princi!SOLMAE_APOPRO * g_rst_Princi!SOLMAE_TCADOL, "###,###,##0.00" & " ")
   pnl_Sol_MtoSol_Sol.Caption = Format(g_rst_Princi!SOLMAE_MTOSOL * g_rst_Princi!SOLMAE_TCADOL, "###,###,##0.00" & " ")
   
   pnl_Sol_ComVta_MPr.Caption = Format(g_rst_Princi!SOLMAE_COMVTA * g_rst_Princi!SOLMAE_TCADOL / g_rst_Princi!SOLMAE_TCAMPR, "###,###,##0.00" & " ")
   pnl_Sol_ApoPro_MPr.Caption = Format(g_rst_Princi!SOLMAE_APOPRO * g_rst_Princi!SOLMAE_TCADOL / g_rst_Princi!SOLMAE_TCAMPR, "###,###,##0.00" & " ")
   pnl_Sol_MtoSol_MPr.Caption = Format(g_rst_Princi!SOLMAE_MTOSOL * g_rst_Princi!SOLMAE_TCADOL / g_rst_Princi!SOLMAE_TCAMPR, "###,###,##0.00" & " ")
  
   pnl_Sol_CuoSeg.Caption = Format(g_rst_Princi!SOLMAE_CUOMEN, "###,###,##0.00") & " "
   pnl_Sol_VctSug.Caption = Format(g_rst_Princi!SOLMAE_DIAVCT, "###,###,##0.00") & " "
   
   txt_Sol_Observ.Locked = True
   txt_Sol_Observ.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   pnl_Eva_SegDes.Caption = Format(moddat_gf_Consulta_AplSeg_Factor(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG, moddat_g_int_TipMon, g_rst_Princi!SOLMAE_PREMPR, l_int_AplDes), "###,##0.000000000") & " "
   pnl_Eva_SegViv.Caption = Format(moddat_gf_Consulta_AplSeg_Factor(g_rst_Princi!SOLMAE_ESGVIV, 0, moddat_g_int_TipMon, g_rst_Princi!SOLMAE_PREMPR, l_int_AplViv), "###,##0.000000000") & " "
   
   lbl_Etique(2).Caption = "Seg. Desg. (" & moddat_gf_Consulta_ParDes("227", CStr(l_int_AplDes)) & "): "
   lbl_Etique(3).Caption = "Seg. Viv. (" & moddat_gf_Consulta_ParDes("227", CStr(l_int_AplViv)) & "): "
   
   pnl_Eva_TasInt.Caption = Format(g_rst_Princi!SOLMAE_TASINT, "###,##0.00") & " "
   
   pnl_Eva_TCaDol.Caption = Format(moddat_gf_Obtiene_TipCam(1, 2), "###,##0.000000") & " "
   pnl_Eva_TCaMPr.Caption = Format(moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon), "###,##0.000000") & " "
   
   l_int_CuoExt = g_rst_Princi!SOLMAE_CUOANO
   l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_Refere()
   'Referencia Familiar
   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' AND SOLREF_TIPREF = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_Fam_TipPar.Caption = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPREF))
   pnl_Fam_Nombre.Caption = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
   pnl_Fam_Telefo.Caption = Trim(g_rst_Princi!SOLREF_TELEFO & "")
   pnl_Fam_Celula.Caption = Trim(g_rst_Princi!SOLREF_TELEFO & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Referencia No Familiar
   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' AND SOLREF_TIPREF = 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_NFa_TipPar.Caption = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPREF))
   pnl_NFa_Nombre.Caption = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
   pnl_NFa_Telefo.Caption = Trim(g_rst_Princi!SOLREF_TELEFO & "")
   pnl_NFa_Celula.Caption = Trim(g_rst_Princi!SOLREF_TELEFO & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
   grd_Doc_Listad.ColWidth(0) = 10080
   grd_Doc_Listad.ColAlignment(0) = flexAlignLeftCenter
   
   grd_Inm_Listad.ColWidth(0) = 9870
   grd_Inm_Listad.ColWidth(1) = 1500
   grd_Inm_Listad.ColWidth(2) = 1560
   grd_Inm_Listad.ColWidth(3) = 0
   
   grd_Inm_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Inm_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Inm_Listad.ColAlignment(2) = flexAlignRightCenter

   grd_Tar_Listad.ColWidth(0) = 3960
   grd_Tar_Listad.ColWidth(1) = 2490
   grd_Tar_Listad.ColWidth(2) = 2040
   grd_Tar_Listad.ColWidth(3) = 1410
   grd_Tar_Listad.ColWidth(4) = 0
   grd_Tar_Listad.ColWidth(5) = 0
   grd_Tar_Listad.ColWidth(6) = 0
   
   grd_Tar_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Tar_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Tar_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Tar_Listad.ColAlignment(3) = flexAlignRightCenter
   
   grd_Fin_Listad.ColWidth(0) = 3960
   grd_Fin_Listad.ColWidth(1) = 2340
   grd_Fin_Listad.ColWidth(2) = 2040
   grd_Fin_Listad.ColWidth(3) = 1560
   grd_Fin_Listad.ColWidth(4) = 0
   grd_Fin_Listad.ColWidth(5) = 0
   grd_Fin_Listad.ColWidth(6) = 0
   
   grd_Fin_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Fin_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Fin_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Fin_Listad.ColAlignment(3) = flexAlignRightCenter
   
   grd_Gas_Listad.ColWidth(0) = 7770
   grd_Gas_Listad.ColWidth(1) = 2010
   
   grd_Gas_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Gas_Listad.ColAlignment(1) = flexAlignRightCenter

   'Actividades Económicas
   grd_Tit_Listad(0).ColWidth(0) = 2000
   grd_Tit_Listad(0).ColWidth(1) = 6850
   grd_Tit_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   
   grd_Tit_Listad(1).ColWidth(0) = 2000
   grd_Tit_Listad(1).ColWidth(1) = 6850
   grd_Tit_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_Listad(1).ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_Listad(0).ColWidth(0) = 2000
   grd_Cyg_Listad(0).ColWidth(1) = 6850
   grd_Cyg_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_Listad(1).ColWidth(0) = 2000
   grd_Cyg_Listad(1).ColWidth(1) = 6850
   grd_Cyg_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_Listad(1).ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Carga_SolInb()
   Call gs_LimpiaGrid(grd_Inm_Listad)
   
   g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
   g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Inm_Listad.Rows = grd_Inm_Listad.Rows + 1
      grd_Inm_Listad.Row = grd_Inm_Listad.Rows - 1
   
      grd_Inm_Listad.Col = 0
      grd_Inm_Listad.Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Princi!SOLINB_TIPINM))
      
      grd_Inm_Listad.Col = 1
      grd_Inm_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLINB_FECADQ))
   
      grd_Inm_Listad.Col = 2
      grd_Inm_Listad.Text = Format(g_rst_Princi!SOLINB_IMPVAL, "###,###,##0.00")
      
      grd_Inm_Listad.Col = 3
      grd_Inm_Listad.Text = Trim(g_rst_Princi!SOLINB_DIRECC & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_Inm_Listad)
   Call grd_Inm_Listad_Click
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_SolTrj()
   Call gs_LimpiaGrid(grd_Tar_Listad)
   
   g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
   g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Tar_Listad.Rows = grd_Tar_Listad.Rows + 1
      grd_Tar_Listad.Row = grd_Tar_Listad.Rows - 1
   
      grd_Tar_Listad.Col = 0
      grd_Tar_Listad.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLTRJ_CODINS)
      
      grd_Tar_Listad.Col = 1
      grd_Tar_Listad.Text = Trim(g_rst_Princi!SOLTRJ_NUMTRJ & "")
   
      grd_Tar_Listad.Col = 2
      grd_Tar_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLTRJ_TIPMON))
      
      grd_Tar_Listad.Col = 3
      grd_Tar_Listad.Text = Format(g_rst_Princi!SOLTRJ_SALACT, "###,###,##0.00")
      
      grd_Tar_Listad.Col = 4
      grd_Tar_Listad.Text = moddat_gf_Consulta_ParDes("506", g_rst_Princi!SOLTRJ_TIPTRJ)
      
      grd_Tar_Listad.Col = 5
      grd_Tar_Listad.Text = Format(g_rst_Princi!SOLTRJ_LIMCRD, "###,###,##0.00")
      
      grd_Tar_Listad.Col = 6
      grd_Tar_Listad.Text = Format(g_rst_Princi!SOLTRJ_PAGMIN, "###,###,##0.00")
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_Tar_Listad)
   Call grd_Tar_Listad_Click
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_SolDeu()
   Call gs_LimpiaGrid(grd_Fin_Listad)
   
   g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
   g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Fin_Listad.Rows = grd_Fin_Listad.Rows + 1
      grd_Fin_Listad.Row = grd_Fin_Listad.Rows - 1
   
      grd_Fin_Listad.Col = 0
      grd_Fin_Listad.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLDEU_CODINS)
      
      grd_Fin_Listad.Col = 1
      grd_Fin_Listad.Text = Trim(g_rst_Princi!SOLDEU_NUMOPE & "")
   
      grd_Fin_Listad.Col = 2
      grd_Fin_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLDEU_TIPMON))
      
      grd_Fin_Listad.Col = 3
      grd_Fin_Listad.Text = Format(g_rst_Princi!SOLDEU_SALPAG, "###,###,##0.00")
      
      grd_Fin_Listad.Col = 4
      grd_Fin_Listad.Text = Format(g_rst_Princi!SOLDEU_MTOOTO, "###,###,##0.00")
      
      grd_Fin_Listad.Col = 5
      grd_Fin_Listad.Text = Format(g_rst_Princi!SOLDEU_CUOMEN, "###,###,##0.00")
      
      grd_Fin_Listad.Col = 6
      grd_Fin_Listad.Text = Format(g_rst_Princi!SOLDEU_PLAMEN, "###,##0")
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_Fin_Listad)
   Call grd_Fin_Listad_Click
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_SolGas()
   Dim r_dbl_TotGas     As Double
   
   pnl_Gas_TotGas.Caption = "0.00 "
   r_dbl_TotGas = 0
   
   Call gs_LimpiaGrid(grd_Gas_Listad)
   
   g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
   g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Gas_Listad.Rows = grd_Gas_Listad.Rows + 1
      grd_Gas_Listad.Row = grd_Gas_Listad.Rows - 1
   
      grd_Gas_Listad.Col = 0
      grd_Gas_Listad.Text = moddat_gf_Consulta_ParDes("220", g_rst_Princi!SOLEYM_CODEYM)
      
      grd_Gas_Listad.Col = 1
      grd_Gas_Listad.Text = Format(g_rst_Princi!SOLEYM_IMPORT, "###,###,##0.00")
      
      r_dbl_TotGas = r_dbl_TotGas + g_rst_Princi!SOLEYM_IMPORT
      
      g_rst_Princi.MoveNext
   Loop
   
   pnl_Gas_TotGas.Caption = Format(r_dbl_TotGas, "###,##0.00") & " "
   
   Call gs_UbiIniGrid(grd_Gas_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Carga_SolDoc()
   Call gs_LimpiaGrid(grd_Doc_Listad)
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Doc_Listad.Rows = grd_Doc_Listad.Rows + 1
      grd_Doc_Listad.Row = grd_Doc_Listad.Rows - 1
   
      grd_Doc_Listad.Col = 0
      
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Doc_Listad.Text = Mid(moddat_g_arr_Genera(1).Genera_Nombre, 5)
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Doc_Listad.Text = Mid(moddat_g_arr_Genera(1).Genera_Nombre, 5)
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_SorteaGrid(grd_Doc_Listad, 0, "C")
   Call gs_UbiIniGrid(grd_Doc_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Fin_Listad_Click()
   If grd_Fin_Listad.Rows > 0 Then
      grd_Fin_Listad.Col = 0
      pnl_Fin_NomIns.Caption = grd_Fin_Listad.Text
      
      grd_Fin_Listad.Col = 1
      pnl_Fin_NumOpe.Caption = grd_Fin_Listad.Text
      
      grd_Fin_Listad.Col = 2
      pnl_Fin_TipMon.Caption = grd_Fin_Listad.Text
      
      grd_Fin_Listad.Col = 3
      pnl_Fin_SalPag.Caption = grd_Fin_Listad.Text & " "
      
      grd_Fin_Listad.Col = 4
      pnl_Fin_MtoOto.Caption = grd_Fin_Listad.Text & " "
      
      grd_Fin_Listad.Col = 5
      pnl_Fin_CuoMen.Caption = grd_Fin_Listad.Text & " "
      
      grd_Fin_Listad.Col = 6
      pnl_Fin_MesPag.Caption = grd_Fin_Listad.Text & " "
      
      Call gs_RefrescaGrid(grd_Fin_Listad)
   End If
End Sub

Private Sub grd_Inm_Listad_Click()
   If grd_Inm_Listad.Rows > 0 Then
      grd_Inm_Listad.Col = 0
      pnl_Inm_TipInm.Caption = grd_Inm_Listad.Text
      
      grd_Inm_Listad.Col = 1
      pnl_Inm_FecAdq.Caption = grd_Inm_Listad.Text
      
      grd_Inm_Listad.Col = 2
      pnl_Inm_Import.Caption = grd_Inm_Listad.Text & " "
      
      grd_Inm_Listad.Col = 3
      pnl_Inm_Direcc.Caption = grd_Inm_Listad.Text
      
      Call gs_RefrescaGrid(grd_Inm_Listad)
   End If
End Sub

Private Sub grd_Inm_Listad_SelChange()
   If grd_Inm_Listad.Rows > 2 Then
      grd_Inm_Listad.RowSel = grd_Inm_Listad.Row
   End If
   
   Call grd_Inm_Listad_Click
End Sub

Private Sub grd_Tar_Listad_Click()
   If grd_Tar_Listad.Rows > 0 Then
      grd_Tar_Listad.Col = 0
      pnl_Tar_NomIns.Caption = grd_Tar_Listad.Text
      
      grd_Tar_Listad.Col = 1
      pnl_Tar_NumTar.Caption = grd_Tar_Listad.Text
      
      grd_Tar_Listad.Col = 2
      pnl_Tar_TipMon.Caption = grd_Tar_Listad.Text
      
      grd_Tar_Listad.Col = 3
      pnl_Tar_SalPag.Caption = grd_Tar_Listad.Text & " "
      
      grd_Tar_Listad.Col = 4
      pnl_Tar_TipTar.Caption = grd_Tar_Listad.Text
      
      grd_Tar_Listad.Col = 5
      pnl_Tar_LinCre.Caption = grd_Tar_Listad.Text & " "
      
      grd_Tar_Listad.Col = 6
      pnl_Tar_MtoMin.Caption = grd_Tar_Listad.Text & " "
      
      Call gs_RefrescaGrid(grd_Tar_Listad)
   End If
End Sub

Private Sub grd_Tar_Listad_SelChange()
   If grd_Tar_Listad.Rows > 2 Then
      grd_Tar_Listad.RowSel = grd_Tar_Listad.Row
   End If
End Sub

Private Sub grd_Fin_Listad_SelChange()
   If grd_Fin_Listad.Rows > 2 Then
      grd_Fin_Listad.RowSel = grd_Fin_Listad.Row
   End If
End Sub

Private Sub grd_Gas_Listad_SelChange()
   If grd_Gas_Listad.Rows > 2 Then
      grd_Gas_Listad.RowSel = grd_Gas_Listad.Row
   End If
End Sub

Private Sub grd_Doc_Listad_SelChange()
   If grd_Doc_Listad.Rows > 2 Then
      grd_Doc_Listad.RowSel = grd_Doc_Listad.Row
   End If
End Sub

Private Sub fs_Buscar_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, p_Grid As MSFlexGrid, p_Panel As SSPanel)
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String
   Dim r_str_TipDoc     As String
   Dim l_rst_Genera     As ADODB.Recordset
   
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Ocupación
   p_Panel.Caption = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ActEco_CodAct))
   
   Select Case g_rst_Princi!ActEco_CodAct
      Case 11, 31, 41
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
      
         If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
            Exit Sub
         End If
   
         l_rst_Genera.MoveFirst
         
         'Documento de Identidad
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Documento de Identidad"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
      
         'Razón Social
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Razón Social"
      
         p_Grid.Col = 1
         p_Grid.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
      
         'Nombre Comercial
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Nombre Comercial"
      
         p_Grid.Col = 1
         p_Grid.Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
      
         'Giro Comercial
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Giro Comercial"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
      
         If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
            p_Grid.Text = p_Grid.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
         End If
      
         'Dirección
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Dirección Empresa"
      
         p_Grid.Col = 1
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))

         p_Grid.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")

         If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
            p_Grid.Text = p_Grid.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
         End If

         If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
            p_Grid.Text = p_Grid.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
         Else
            p_Grid.Text = p_Grid.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
   
         p_Grid.Text = p_Grid.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
         'Teléfono
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Teléfono(s) Empresa"
      
         p_Grid.Col = 1
         p_Grid.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
         If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
            p_Grid.Text = p_Grid.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         'Sucursal
         If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Sucursal"
         
            p_Grid.Col = 1
            p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_SUCURS & "")
            
            'Dirección Sucursal
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Dirección Sucursal"
         
            p_Grid.Col = 1
            
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

            p_Grid.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
               p_Grid.Text = p_Grid.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
               p_Grid.Text = p_Grid.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
            Else
               p_Grid.Text = p_Grid.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
      
            p_Grid.Text = p_Grid.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfono Sucursal
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Teléfono(s) Sucursal"
         
            p_Grid.Col = 1
            p_Grid.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
            If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
               p_Grid.Text = p_Grid.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
            End If
         End If
         
         
         If g_rst_Princi!ActEco_CodAct = 11 Or g_rst_Princi!ActEco_CodAct = 12 Then
            'Teléfono y Anexo RR.HH
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Teléfono RR.HH"
         
            p_Grid.Col = 1
            
            If Len(Trim(l_rst_Genera!DATGEN_TELERH & "")) = 0 Then
               p_Grid.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            Else
               p_Grid.Text = Trim(l_rst_Genera!DATGEN_TELERH & "")
            End If
            
            If Len(Trim(l_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
               p_Grid.Text = p_Grid.Text & " - " & Trim(l_rst_Genera!DATGEN_ANEXRH & "")
            End If
         
            'Cargo
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Cargo"
         
            p_Grid.Col = 1
            If Len(Trim(g_rst_Princi!ActEco_Dep_CargoN & "")) > 0 Then
               p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_CargoN)
            Else
               p_Grid.Text = moddat_gf_Consulta_ParDes("503", Trim(g_rst_Princi!ActEco_Dep_CargoC))
            End If
         
            'Area
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Area"
         
            p_Grid.Col = 1
            p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NomAre)
            
            'Número Anexo
            If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1
               p_Grid.Row = p_Grid.Rows - 1
               
               p_Grid.Col = 0
               p_Grid.Text = "Anexo"
            
               p_Grid.Col = 1
               p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx)
            End If
            
            'Teléfono Directo
            If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1
               p_Grid.Row = p_Grid.Rows - 1
               
               p_Grid.Col = 0
               p_Grid.Text = "Teléfono Directo"
            
               p_Grid.Col = 1
               p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir)
            End If
         
            'Celular Laboral
            If Len(Trim(g_rst_Princi!ActEco_Dep_Celula)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1
               p_Grid.Row = p_Grid.Rows - 1
               
               p_Grid.Col = 0
               p_Grid.Text = "Celular Laboral"
            
               p_Grid.Col = 1
               p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Celula)
            End If
         
            'E-mail
            If Len(Trim(g_rst_Princi!ActEco_Dep_DirEle)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1
               p_Grid.Row = p_Grid.Rows - 1
               
               p_Grid.Col = 0
               p_Grid.Text = "E-mail"
            
               p_Grid.Col = 1
               p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_DirEle)
            End If
         End If
         
         l_rst_Genera.Close
         Set l_rst_Genera = Nothing
         
         Call gs_UbiIniGrid(p_Grid)
         
      Case 21
         'Documento de Identidad
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Documento de Identidad"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Dirección Tributaria
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Dirección Tributaria"
      
         p_Grid.Col = 1
         
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

         p_Grid.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")

         If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
            p_Grid.Text = p_Grid.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
         End If

         If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
            p_Grid.Text = p_Grid.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
         Else
            p_Grid.Text = p_Grid.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
   
         p_Grid.Text = p_Grid.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
      
         'Teléfono
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Teléfono(s) "
      
         p_Grid.Col = 1
         p_Grid.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
         
         If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
            p_Grid.Text = p_Grid.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
         End If
         
         'Giro Comercial
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Giro Comercial"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Grid.Text = p_Grid.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         'Contrato de Locación de Servicios
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Contrato Locación "
         
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
         
         If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
      
            If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
               Exit Sub
            End If
   
            l_rst_Genera.MoveFirst
         
            'Documento de Identidad
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
         
            p_Grid.Col = 0
            p_Grid.Text = "Documento de Identidad"
      
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(l_rst_Genera!DatGen_EMPTDO)) & " - " & Trim(l_rst_Genera!DatGen_EMPNDO)
      
            'Razón Social
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
         
            p_Grid.Col = 0
            p_Grid.Text = "Razón Social"
      
            p_Grid.Col = 1
            p_Grid.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
         
            'Nombre Comercial
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Nombre Comercial"
         
            p_Grid.Col = 1
            p_Grid.Text = Trim(l_rst_Genera!DATGEN_NOMCOM & "")
         
            'Giro Comercial
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Giro Comercial"
         
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
         
            If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
               p_Grid.Text = p_Grid.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
            End If
         
            'Dirección
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
            
            p_Grid.Col = 0
            p_Grid.Text = "Dirección Empresa"
         
            p_Grid.Col = 1
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))
   
            p_Grid.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
               p_Grid.Text = p_Grid.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
               p_Grid.Text = p_Grid.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
            Else
               p_Grid.Text = p_Grid.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
      
            p_Grid.Text = p_Grid.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfonos
            p_Grid.Rows = p_Grid.Rows + 1
            p_Grid.Row = p_Grid.Rows - 1
         
            p_Grid.Col = 0
            p_Grid.Text = "Teléfonos"
            
            p_Grid.Col = 1
            p_Grid.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
            If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
               p_Grid.Text = p_Grid.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
            End If
         End If
         
         Call gs_UbiIniGrid(p_Grid)
         
      Case 51
         'Documento de Identidad
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Documento de Identidad"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Giro Comercial
         p_Grid.Rows = p_Grid.Rows + 1
         p_Grid.Row = p_Grid.Rows - 1
         
         p_Grid.Col = 0
         p_Grid.Text = "Giro Comercial"
      
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Grid.Text = p_Grid.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         Call gs_UbiIniGrid(p_Grid)
   End Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_Eva_MtoCre_Change()
   Call ipp_Eva_TotILD_Change
End Sub

Private Sub ipp_Eva_MtoCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Calcul)
   End If
End Sub

Private Sub ipp_Eva_NumCuo_Change()
   Call ipp_Eva_TotILD_Change
End Sub

Private Sub ipp_Eva_NumCuo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Eva_MtoCre)
   End If
End Sub

Private Sub ipp_Eva_TotILD_Change()
   pnl_Eva_CuoMen_Dol.Caption = "0.00 "
   pnl_Eva_CuoMen_Sol.Caption = "0.00 "
   pnl_Eva_CuoMen_MPr.Caption = "0.00 "
   
   pnl_Eva_CuoRta.Caption = "0.00 "
   pnl_Eva_PriCuo.Caption = "0.00 "
   pnl_Eva_UltCuo.Caption = "0.00 "
   
   l_int_FlgCal = 0
End Sub

Private Sub ipp_Eva_TotILD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Eva_NumCuo)
   End If
End Sub

Private Sub txt_Eva_Observ_GotFocus()
   Call gs_SelecTodo(txt_Eva_Observ)
End Sub

Private Sub txt_Eva_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Carga_EvaCre()
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "EVACRE_TIPEVA = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      moddat_g_int_FlgGrb = 1
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   ipp_Eva_TotILD.Value = g_rst_Princi!EVACRE_APRIN1
   ipp_Eva_NumCuo.Value = g_rst_Princi!EVACRE_APRPLA
   ipp_Eva_MtoCre.Value = g_rst_Princi!EVACRE_APRDOL
   
   pnl_Eva_CuoMen_Dol.Caption = Format(g_rst_Princi!EVACRE_APRCUO, "###,###,##0.00") & " "
   pnl_Eva_PriCuo.Caption = Format(g_rst_Princi!EVACRE_APRCIN, "###,###,##0.00") & " "
   pnl_Eva_UltCuo.Caption = Format(g_rst_Princi!EVACRE_APRCFN, "###,###,##0.00") & " "
   
   'Calculando Valores Referenciales
   pnl_Eva_CuoMen_Sol.Caption = Format(pnl_Eva_CuoMen_Dol.Caption * CDbl(pnl_Eva_TCaDol.Caption), "###,###,0.00") & " "
   pnl_Eva_CuoMen_MPr.Caption = Format(pnl_Eva_CuoMen_Dol.Caption * CDbl(pnl_Eva_TCaDol.Caption) / CDbl(pnl_Eva_TCaMPr.Caption), "###,###,0.00") & " "
   pnl_Eva_CuoRta.Caption = Format(CDbl(pnl_Eva_CuoMen_Sol.Caption) / CDbl(ipp_Eva_TotILD.Text) * 100, "##0.00") & " "
   
   txt_Eva_Observ.Text = Trim(g_rst_Princi!EVACRE_OBSERV & "")

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   moddat_g_int_FlgGrb = 2
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaCre) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function ff_Buscar_LisObs() As Integer
   ff_Buscar_LisObs = False
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaCre) & " AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!SEGFECACT = 0 Then
         ff_Buscar_LisObs = True
         
         Exit Do
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

