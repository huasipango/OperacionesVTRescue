VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGeneraSyC 
   Caption         =   "Validación de Productos Electrónicos (S & C)"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frBotones 
      Height          =   855
      Left            =   2880
      TabIndex        =   6
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   840
         Picture         =   "frmGeneraSyC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   1440
         Picture         =   "frmGeneraSyC.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdImprime 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   240
         Picture         =   "frmGeneraSyC.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Subir desde archivo"
         Top             =   240
         Width           =   400
      End
   End
   Begin VB.Frame frFecha 
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   1800
         Picture         =   "frmGeneraSyC.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Subir desde archivo"
         Top             =   240
         Width           =   400
      End
      Begin MSMask.MaskEdBox mskFechaProc 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAño1 
         Caption         =   "Fecha de Proceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkOtroDir 
      Caption         =   "Envia a carpeta de Prueba"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   9000
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin TabDlg.SSTab ssTab1 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Mart"
      TabPicture(0)   =   "frmGeneraSyC.frx":0408
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "spdDomAlta1"
      Tab(0).Control(1)=   "spdDomCambio1"
      Tab(0).Control(2)=   "spdDomBaja1"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(8)=   "Frame1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Fuel Magna"
      TabPicture(1)   =   "frmGeneraSyC.frx":0424
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame14"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame10"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "spdDomAlta2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "spdDomCambio2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "spdDomBaja2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Mart Abierto"
      TabPicture(2)   =   "frmGeneraSyC.frx":0440
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "spdDomAlta3"
      Tab(2).Control(1)=   "spdDomCambio3"
      Tab(2).Control(2)=   "spdDomBaja3"
      Tab(2).Control(3)=   "Frame16"
      Tab(2).Control(4)=   "Frame21"
      Tab(2).Control(5)=   "Frame20"
      Tab(2).Control(6)=   "Frame19"
      Tab(2).Control(7)=   "Frame18"
      Tab(2).Control(8)=   "Frame15"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Fuel Premium"
      TabPicture(3)   =   "frmGeneraSyC.frx":045C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(2)=   "Frame17"
      Tab(3).Control(3)=   "Frame22"
      Tab(3).Control(4)=   "Frame23"
      Tab(3).Control(5)=   "Frame24"
      Tab(3).Control(6)=   "spdDomAlta4"
      Tab(3).Control(7)=   "spdDomCambio4"
      Tab(3).Control(8)=   "spdDomBaja4"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Fuel Diesel"
      TabPicture(4)   =   "frmGeneraSyC.frx":0478
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame26"
      Tab(4).Control(1)=   "Frame25"
      Tab(4).Control(2)=   "Frame27"
      Tab(4).Control(3)=   "Frame28"
      Tab(4).Control(4)=   "Frame29"
      Tab(4).Control(5)=   "Frame30"
      Tab(4).Control(6)=   "spdDomAlta5"
      Tab(4).Control(7)=   "spdDomCambio5"
      Tab(4).Control(8)=   "spdDomBaja5"
      Tab(4).ControlCount=   9
      Begin FPSpread.vaSpread spdDomBaja5 
         Height          =   735
         Left            =   -68640
         OleObjectBlob   =   "frmGeneraSyC.frx":0494
         TabIndex        =   137
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomCambio5 
         Height          =   735
         Left            =   -68640
         OleObjectBlob   =   "frmGeneraSyC.frx":0791
         TabIndex        =   136
         Top             =   6120
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomAlta5 
         Height          =   735
         Left            =   -68640
         OleObjectBlob   =   "frmGeneraSyC.frx":0A8E
         TabIndex        =   135
         Top             =   5400
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomBaja4 
         Height          =   735
         Left            =   -69240
         OleObjectBlob   =   "frmGeneraSyC.frx":0D8B
         TabIndex        =   106
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomCambio4 
         Height          =   735
         Left            =   -69240
         OleObjectBlob   =   "frmGeneraSyC.frx":1088
         TabIndex        =   105
         Top             =   6120
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomAlta4 
         Height          =   735
         Left            =   -69240
         OleObjectBlob   =   "frmGeneraSyC.frx":1385
         TabIndex        =   104
         Top             =   5400
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomBaja3 
         Height          =   735
         Left            =   -69120
         OleObjectBlob   =   "frmGeneraSyC.frx":1682
         TabIndex        =   102
         Top             =   6900
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomCambio3 
         Height          =   735
         Left            =   -69120
         OleObjectBlob   =   "frmGeneraSyC.frx":197F
         TabIndex        =   101
         Top             =   6180
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomAlta3 
         Height          =   735
         Left            =   -69120
         OleObjectBlob   =   "frmGeneraSyC.frx":1C7C
         TabIndex        =   100
         Top             =   5460
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomBaja2 
         Height          =   735
         Left            =   5760
         OleObjectBlob   =   "frmGeneraSyC.frx":1F79
         TabIndex        =   99
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomCambio2 
         Height          =   735
         Left            =   5760
         OleObjectBlob   =   "frmGeneraSyC.frx":2276
         TabIndex        =   98
         Top             =   5880
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomAlta2 
         Height          =   735
         Left            =   5520
         OleObjectBlob   =   "frmGeneraSyC.frx":2573
         TabIndex        =   97
         Top             =   5280
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomBaja1 
         Height          =   735
         Left            =   -68880
         OleObjectBlob   =   "frmGeneraSyC.frx":2870
         TabIndex        =   96
         Top             =   6780
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomCambio1 
         Height          =   735
         Left            =   -68880
         OleObjectBlob   =   "frmGeneraSyC.frx":2B6D
         TabIndex        =   95
         Top             =   6060
         Visible         =   0   'False
         Width           =   4455
      End
      Begin FPSpread.vaSpread spdDomAlta1 
         Height          =   735
         Left            =   -68880
         OleObjectBlob   =   "frmGeneraSyC.frx":2E6A
         TabIndex        =   94
         Top             =   5340
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Frame Frame30 
         Caption         =   "Alta de Empleadoras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   164
         Top             =   1680
         Width           =   4695
         Begin FPSpread.vaSpread spdClientes5 
            Height          =   735
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":3167
            TabIndex        =   165
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -69840
         TabIndex        =   161
         Top             =   360
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaTit5 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":3464
            TabIndex        =   163
            Top             =   360
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaTit5 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":370B
            TabIndex        =   162
            Top             =   1440
            Width           =   4455
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Stock de Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -69840
         TabIndex        =   158
         Top             =   3000
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaStockTit5 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":39B2
            TabIndex        =   160
            Top             =   240
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaStockTit5 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":3C59
            TabIndex        =   159
            Top             =   1200
            Width           =   4455
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Ajustes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   153
         Top             =   3120
         Width           =   4695
         Begin FPSpread.vaSpread spdAjustes5 
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":3F00
            TabIndex        =   154
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtTotEmpAju5 
            Height          =   285
            Left            =   2520
            TabIndex        =   156
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtAjustes5 
            Height          =   285
            Left            =   3360
            TabIndex        =   155
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   157
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Tipos de Archivos revisados para generar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   138
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox chkdie_Empleadoras 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkdie_Tarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   145
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkdie_Cancel 
            Caption         =   "Baja de Tit. y Adic."
            Height          =   255
            Left            =   2520
            TabIndex        =   144
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkdie_Ajustes 
            Caption         =   "Ajus"
            Height          =   255
            Left            =   960
            TabIndex        =   143
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkdie_Disp 
            Caption         =   "Disp"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkdie_StockCancel 
            Caption         =   "Baja de Stock Tit y Adic"
            Height          =   255
            Left            =   2520
            TabIndex        =   141
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkdie_StockTarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   140
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkDie_Domicilios 
            Caption         =   "Domic."
            Height          =   255
            Left            =   960
            TabIndex        =   139
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Alta de Empleadoras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   133
         Top             =   1680
         Width           =   4695
         Begin FPSpread.vaSpread spdClientes4 
            Height          =   735
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":4246
            TabIndex        =   134
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -69840
         TabIndex        =   130
         Top             =   360
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaTit4 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":4543
            TabIndex        =   132
            Top             =   360
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaTit4 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":47EA
            TabIndex        =   131
            Top             =   1440
            Width           =   4455
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Stock de Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -69840
         TabIndex        =   127
         Top             =   3000
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaStockTit4 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":4A91
            TabIndex        =   129
            Top             =   240
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaStockTit4 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":4D38
            TabIndex        =   128
            Top             =   1200
            Width           =   4455
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Ajustes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   122
         Top             =   3120
         Width           =   4695
         Begin FPSpread.vaSpread spdAjustes4 
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":4FDF
            TabIndex        =   123
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtTotEmpAju4 
            Height          =   285
            Left            =   2520
            TabIndex        =   125
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtAjustes4 
            Height          =   285
            Left            =   3360
            TabIndex        =   124
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   126
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Tipos de Archivos revisados para generar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   107
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox chkpre_Empleadoras 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkpre_Tarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   114
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkpre_Cancel 
            Caption         =   "Baja de Tit. y Adic."
            Height          =   255
            Left            =   2520
            TabIndex        =   113
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkpre_Ajustes 
            Caption         =   "Ajus"
            Height          =   255
            Left            =   960
            TabIndex        =   112
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkpre_Disp 
            Caption         =   "Disp"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkpre_StockCancel 
            Caption         =   "Baja de Stock Tit y Adic"
            Height          =   255
            Left            =   2520
            TabIndex        =   110
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkpre_StockTarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   109
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkPre_Domicilios 
            Caption         =   "Domic."
            Height          =   255
            Left            =   960
            TabIndex        =   108
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Stock de Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -69480
         TabIndex        =   88
         Top             =   2940
         Width           =   4695
         Begin FPSpread.vaSpread spdBajaStockTit 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":5325
            TabIndex        =   90
            Top             =   1200
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdAltaStockTit 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":55CC
            TabIndex        =   89
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -69480
         TabIndex        =   85
         Top             =   420
         Width           =   4695
         Begin FPSpread.vaSpread spdBajaTit 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":5873
            TabIndex        =   87
            Top             =   1440
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdAltaTit 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":5B1A
            TabIndex        =   86
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Stock de Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   5160
         TabIndex        =   82
         Top             =   3060
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaStockTit2 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":5DC1
            TabIndex        =   84
            Top             =   240
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaStockTit2 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":6068
            TabIndex        =   83
            Top             =   1200
            Width           =   4455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5160
         TabIndex        =   79
         Top             =   540
         Width           =   4695
         Begin FPSpread.vaSpread spdAltaTit2 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":630F
            TabIndex        =   81
            Top             =   360
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdBajaTit2 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":65B6
            TabIndex        =   80
            Top             =   1440
            Width           =   4455
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Tipos de Archivos revisados para generar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   62
         Top             =   420
         Width           =   4695
         Begin VB.CheckBox chkReg_Domicilios 
            Caption         =   "Domic."
            Height          =   255
            Left            =   960
            TabIndex        =   93
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkreg_StockTarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   75
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkreg_StockCancel 
            Caption         =   "Baja de Stock Tit y Adic"
            Height          =   255
            Left            =   2520
            TabIndex        =   74
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkreg_Disp 
            Caption         =   "Disp"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkreg_Ajustes 
            Caption         =   "Ajus"
            Height          =   255
            Left            =   960
            TabIndex        =   66
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkreg_Cancel 
            Caption         =   "Baja de Tit. y Adic."
            Height          =   255
            Left            =   2520
            TabIndex        =   65
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkreg_Tarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   64
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkreg_Empleadoras 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tipos de Archivos revisados para generar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   56
         Top             =   540
         Width           =   4695
         Begin VB.CheckBox chkAli_Domicilios 
            Caption         =   "Domic."
            Height          =   255
            Left            =   960
            TabIndex        =   92
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkali_StockCancel 
            Caption         =   "Baja de Stock Tit y Adic"
            Height          =   255
            Left            =   2400
            TabIndex        =   73
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkali_StockTarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   72
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkali_Empleadoras 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkali_Tarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   60
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkali_Cancel 
            Caption         =   "Baja de Tit. y Adic."
            Height          =   255
            Left            =   2400
            TabIndex        =   59
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkali_Ajustes 
            Caption         =   "Ajus"
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkali_Disp 
            Caption         =   "Disp"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Dispersiones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         TabIndex        =   50
         Top             =   5340
         Width           =   9735
         Begin FPSpread.vaSpread spdDisp3 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":685D
            TabIndex        =   54
            Top             =   360
            Width           =   9255
         End
         Begin VB.TextBox txtTotImp3 
            Height          =   285
            Left            =   7920
            TabIndex        =   53
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtTotEmp3 
            Height          =   285
            Left            =   6840
            TabIndex        =   52
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox chkTodos3 
            Caption         =   "Seleccionar Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   255
            Left            =   5880
            TabIndex        =   55
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Ajustes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   45
         Top             =   3180
         Width           =   4695
         Begin FPSpread.vaSpread spdAjustes3 
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":7A3A
            TabIndex        =   48
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtAjustes3 
            Height          =   285
            Left            =   3360
            TabIndex        =   47
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtTotEmpAju3 
            Height          =   285
            Left            =   2520
            TabIndex        =   46
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   49
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Stock de Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -69720
         TabIndex        =   44
         Top             =   3060
         Width           =   4695
         Begin FPSpread.vaSpread spdBajaStockTit3 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":7D80
            TabIndex        =   78
            Top             =   1200
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdAltaStockTit3 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":8027
            TabIndex        =   77
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Titulares/Adicionales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -69720
         TabIndex        =   42
         Top             =   480
         Width           =   4695
         Begin FPSpread.vaSpread spdBajaTit3 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":82CE
            TabIndex        =   76
            Top             =   1440
            Width           =   4455
         End
         Begin FPSpread.vaSpread spdAltaTit3 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":8575
            TabIndex        =   43
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Alta de Empleadoras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   41
         Top             =   1740
         Width           =   4695
         Begin FPSpread.vaSpread spdClientes3 
            Height          =   735
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":881C
            TabIndex        =   69
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Dispersiones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   35
         Top             =   5220
         Width           =   9735
         Begin FPSpread.vaSpread spdDisp2 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":8B19
            TabIndex        =   36
            Top             =   360
            Width           =   9255
         End
         Begin VB.CheckBox chkCombustibles 
            Caption         =   "Tipos Combustible Tarjetas"
            Height          =   255
            Left            =   2640
            TabIndex        =   103
            Top             =   2160
            Width           =   2535
         End
         Begin VB.CheckBox chkTodos2 
            Caption         =   "Seleccionar Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtTotEmp2 
            Height          =   285
            Left            =   6840
            TabIndex        =   38
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtTotImp2 
            Height          =   285
            Left            =   7920
            TabIndex        =   37
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   255
            Left            =   5880
            TabIndex        =   40
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Ajustes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   30
         Top             =   3300
         Width           =   4695
         Begin FPSpread.vaSpread spdAjustes2 
            Height          =   975
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":9CF6
            TabIndex        =   31
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtTotEmpAju2 
            Height          =   285
            Left            =   2520
            TabIndex        =   33
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAjustes2 
            Height          =   285
            Left            =   3360
            TabIndex        =   32
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   34
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Alta de Empleadoras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   29
         Top             =   1980
         Width           =   4695
         Begin FPSpread.vaSpread spdClientes2 
            Height          =   735
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":A03C
            TabIndex        =   68
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tipos de Archivos revisados para generar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74640
         TabIndex        =   23
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox chkUni_Domicilios 
            Caption         =   "Domic."
            Height          =   255
            Left            =   960
            TabIndex        =   91
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkUni_StockCancel 
            Caption         =   "Baja de Stock Tit y Adic"
            Height          =   255
            Left            =   2400
            TabIndex        =   71
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkUni_StockTarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   70
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkUni_Disp 
            Caption         =   "Disp"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   735
         End
         Begin VB.CheckBox chkUni_Ajustes 
            Caption         =   "Ajus"
            Height          =   255
            Left            =   960
            TabIndex        =   27
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkUni_Cancel 
            Caption         =   "Baja de Tit. y Adic."
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkUni_Tarjetas 
            Caption         =   "Alta"
            Height          =   255
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkUni_Empleadoras 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dispersiones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   17
         Top             =   5160
         Width           =   9735
         Begin FPSpread.vaSpread spdDisp 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":A339
            TabIndex        =   18
            Top             =   360
            Width           =   9255
         End
         Begin VB.TextBox txtTotImp 
            Height          =   285
            Left            =   7920
            TabIndex        =   21
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtTotEmp 
            Height          =   285
            Left            =   6840
            TabIndex        =   20
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Seleccionar Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblTotDisp 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   255
            Left            =   5880
            TabIndex        =   22
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ajustes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   12
         Top             =   3180
         Width           =   4695
         Begin FPSpread.vaSpread spdAjustes 
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":B516
            TabIndex        =   13
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtAjustes 
            Height          =   285
            Left            =   3360
            TabIndex        =   15
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtTotEmpAju 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Total:"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   16
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Alta de Empleadoras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   10
         Top             =   1740
         Width           =   4935
         Begin FPSpread.vaSpread spdClientes 
            Height          =   735
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":B85C
            TabIndex        =   11
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Dispersiones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   116
         Top             =   5280
         Width           =   9735
         Begin FPSpread.vaSpread spdDisp4 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":BB59
            TabIndex        =   117
            Top             =   360
            Width           =   9255
         End
         Begin VB.CheckBox chkTodos4 
            Caption         =   "Seleccionar Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtTotEmp4 
            Height          =   285
            Left            =   6840
            TabIndex        =   119
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtTotImp4 
            Height          =   285
            Left            =   7920
            TabIndex        =   118
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   255
            Left            =   5880
            TabIndex        =   121
            Top             =   2160
            Width           =   855
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Dispersiones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   147
         Top             =   5280
         Width           =   9735
         Begin FPSpread.vaSpread spdDisp5 
            Height          =   1695
            Left            =   120
            OleObjectBlob   =   "frmGeneraSyC.frx":CD36
            TabIndex        =   148
            Top             =   360
            Width           =   9255
         End
         Begin VB.CheckBox chkTodos5 
            Caption         =   "Seleccionar Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtTotEmp5 
            Height          =   285
            Left            =   6840
            TabIndex        =   150
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtTotImp5 
            Height          =   285
            Left            =   7920
            TabIndex        =   149
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            Height          =   255
            Left            =   5880
            TabIndex        =   152
            Top             =   2160
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmGeneraSyC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim dir_prueba As String, cadenafija As String, Ventana As String
Dim nemp_uni As Integer, nemp_ali As Integer, nemp_reg As Integer

Public Function CreaDir(ValDir As String)
Dim AttrDev%
  On Error Resume Next
  ValDir = dir_prueba
  AttrDev = GetAttr(ValDir)
  
  If ERR.Number Then
     ERR.Clear
     MkDir ValDir
  End If
End Function





Private Sub chkTodos_Click()
Dim importe As Double
If chkTodos.value = 1 Then
        spdDisp.Col = 7
        spdDisp.Row = -1
        spdDisp.value = 1
        Numepl = 0
        importe = 0
        With spdDisp
         For i = 1 To .MaxRows
             .Row = i
             .Col = 7
             If .value = 1 Then
                 .Col = 5  'empleados
                     NumEmpl = NumEmpl + Val(.Text)
                 .Col = 6 ' importe
                     importe = importe + CDbl(.Text)
             End If
         Next i
       
        txtTotEmp = NumEmpl
        txtTotImp = Format(importe, "##,###,###.00")
        End With
                
Else
        spdDisp.Col = 7
        spdDisp.Row = -1
        spdDisp.value = 0
        txtTotEmp = 0
        txtTotImp = 0
End If
End Sub

Private Sub chkTodos2_Click()
Dim importe As Double
If chkTodos2.value = 1 Then
        spdDisp2.Col = 7
        spdDisp2.Row = -1
        spdDisp2.value = 1
        Numepl = 0
        importe = 0
        With spdDisp2
         For i = 1 To .MaxRows
             .Row = i
             .Col = 7
             If .value = 1 Then
                 .Col = 5  'empleados
                     NumEmpl = NumEmpl + Val(.Text)
                 .Col = 6 ' importe
                     importe = importe + CDbl(.Text)
             End If
         Next i
       
        txtTotEmp2 = NumEmpl
        txtTotImp2 = Format(importe, "##,###,###.00")
        End With
                
Else
        spdDisp2.Col = 7
        spdDisp2.Row = -1
        spdDisp2.value = 0
        txtTotEmp2 = 0
        txtTotImp2 = 0
End If
End Sub

Private Sub chkTodos3_Click()
Dim importe As Double
If chkTodos3.value = 1 Then
        spdDisp3.Col = 7
        spdDisp3.Row = -1
        spdDisp3.value = 1
        Numepl = 0
        importe = 0
        With spdDisp3
         For i = 1 To .MaxRows
             .Row = i
             .Col = 7
             If .value = 1 Then
                 .Col = 5  'empleados
                     NumEmpl = NumEmpl + Val(.Text)
                 .Col = 6 ' importe
                     importe = importe + CDbl(.Text)
             End If
         Next i
       
        txtTotEmp3 = NumEmpl
        txtTotImp3 = Format(importe, "##,###,###.00")
        End With
                
Else
        spdDisp3.Col = 7
        spdDisp3.Row = -1
        spdDisp3.value = 0
        txtTotEmp3 = 0
        txtTotImp3 = 0
End If
End Sub
Private Sub chkTodos4_Click()
Dim importe As Double
If chkTodos4.value = 1 Then
        spdDisp4.Col = 7
        spdDisp4.Row = -1
        spdDisp4.value = 1
        Numepl = 0
        importe = 0
        With spdDisp4
         For i = 1 To .MaxRows
             .Row = i
             .Col = 7
             If .value = 1 Then
                 .Col = 5  'empleados
                     NumEmpl = NumEmpl + Val(.Text)
                 .Col = 6 ' importe
                     importe = importe + CDbl(.Text)
             End If
         Next i
       
        txtTotEmp4 = NumEmpl
        txtTotImp4 = Format(importe, "##,###,###.00")
        End With
                
Else
        spdDisp4.Col = 7
        spdDisp4.Row = -1
        spdDisp4.value = 0
        txtTotEmp4 = 0
        txtTotImp4 = 0
End If
End Sub
Private Sub chkTodos5_Click()
Dim importe As Double
If chkTodos5.value = 1 Then
        spdDisp5.Col = 7
        spdDisp5.Row = -1
        spdDisp5.value = 1
        Numepl = 0
        importe = 0
        With spdDisp5
         For i = 1 To .MaxRows
             .Row = i
             .Col = 7
             If .value = 1 Then
                 .Col = 5  'empleados
                     NumEmpl = NumEmpl + Val(.Text)
                 .Col = 6 ' importe
                     importe = importe + CDbl(.Text)
             End If
         Next i
       
        txtTotEmp5 = NumEmpl
        txtTotImp5 = Format(importe, "##,###,###.00")
        End With
                
Else
        spdDisp5.Col = 7
        spdDisp5.Row = -1
        spdDisp5.value = 0
        txtTotEmp5 = 0
        txtTotImp5 = 0
End If
End Sub

Private Sub cmdAbrir_Click()
  cmdGrabar.Enabled = False
  CargaDatos
End Sub

Private Sub cmdGrabar_Click()
Dim resp, okas As Boolean
On Error GoTo ERR:
okas = False
If chkOtroDir.value = 1 Then
   dir_prueba = gstrPath & "\Prueba"
Else
   dir_prueba = gstrPath
End If

resp = MsgBox("¿Desea generar los archivos para S&C ? Una vez procesado el día, ya no se podra generar informacion de tarjetas y clientes, solo dispersiones", vbYesNo + vbQuestion + vbDefaultButton2, "Generacion de archivos electrónicos")
If resp = vbYes Then
       If chkUni_Empleadoras.value = 1 Or chkali_Empleadoras.value = 1 Or chkreg_Empleadoras.value = 1 Or chkpre_Empleadoras.value = 1 Or chkdie_Empleadoras.value = 1 Then
          If spdClientes.MaxRows > 0 Or spdClientes2.MaxRows > 0 Or spdClientes3.MaxRows > 0 Or spdClientes4.MaxRows > 0 Or spdClientes5.MaxRows > 0 Then
             okas = True
             GeneraArchivoAltaEmpleadora
         End If
       End If
       If chkUni_Tarjetas.value = 1 Or chkali_Tarjetas.value = 1 Or chkreg_Tarjetas.value = 1 Or chkpre_Tarjetas.value = 1 Or chkdie_Tarjetas.value = 1 Then
          If spdAltaTit.MaxRows > 0 Or spdAltaTit2.MaxRows > 0 Or spdAltaTit3.MaxRows > 0 Or spdAltaTit4.MaxRows Or spdAltaTit5.MaxRows > 0 > 0 Then
             okas = True
             GeneraArchivoTarjetas
         End If
       End If
       If chkUni_StockTarjetas.value = 1 Or chkali_StockTarjetas.value = 1 Or chkreg_StockTarjetas.value = 1 Or chkpre_StockTarjetas.value = 1 Or chkdie_StockTarjetas.value = 1 Then
          If spdAltaStockTit.MaxRows > 0 Or spdAltaStockTit2.MaxRows > 0 Or spdAltaStockTit3.MaxRows > 0 Or spdAltaStockTit4.MaxRows > 0 Or spdAltaStockTit5.MaxRows > 0 Then
             okas = True
             GeneraArchivoStockTarjetas
         End If
       End If
       If chkUni_Cancel.value = 1 Or chkali_Cancel.value = 1 Or chkreg_Tarjetas.value = 1 Or chkpre_Tarjetas.value = 1 Or chkdie_Tarjetas.value = 1 Then
          If spdBajaTit.MaxRows > 0 Or spdBajaTit2.MaxRows > 0 Or spdBajaTit3.MaxRows > 0 Or spdBajaTit4.MaxRows > 0 Or spdBajaTit5.MaxRows > 0 Then
             okas = True
             GeneraArchivoBajaTit
         End If
       End If
       If chkUni_Ajustes.value = 1 Or chkali_Ajustes.value = 1 Or chkreg_Ajustes.value = 1 Or chkpre_Ajustes.value = 1 Or chkdie_Ajustes.value = 1 Then
          If spdAjustes.MaxRows > 0 Or spdAjustes2.MaxRows > 0 Or spdAjustes3.MaxRows > 0 Or spdAjustes4.MaxRows > 0 Or spdAjustes5.MaxRows > 0 Then
             okas = True
             GeneraArchivoAjustes
         End If
       End If
       If chkUni_Disp.value = 1 Or chkali_Disp.value = 1 Or chkreg_Disp.value = 1 Or chkpre_Disp.value = 1 Or chkdie_Disp.value = 1 Then
          If spdDisp.MaxRows > 0 Or spdDisp2.MaxRows > 0 Or spdDisp3.MaxRows > 0 Or spdDisp4.MaxRows > 0 Or spdDisp5.MaxRows > 0 Then
             okas = True
             GeneraArchivoDispersiones
         End If
       End If
'       If chkUni_Domicilios.value = 1 Or chkAli_Domicilios.value = 1 Or chkReg_Domicilios.value = 1 Then
'          If spdDomAlta1.MaxRows > 0 Or spdDomAlta2.MaxRows > 0 Or spdDomAlta3.MaxRows > 0 _
'          Or spdDomBaja1.MaxRows > 0 Or spdDomBaja2.MaxRows > 0 Or spdDomBaja3.MaxRows > 0 _
'          Or spdDomCambio1.MaxRows > 0 Or spdDomCambio2.MaxRows > 0 Or spdDomCambio3.MaxRows > 0 Then
 '            okas = True
 '            GeneraArchivoDomicilios
 '        End If
 '      End If
       If chkCombustibles.value = 1 Then
             okas = True
             GeneraArchivoDomicilios
             GeneraArchivoCombustibles
       End If
       If okas = True Then
          MsgBox "Archivos generados " & dir_prueba, vbInformation, "Sistema Bono Electronico"
       Else
          MsgBox "No se generaron archivos", vbExclamation, "Sistema Bono Electronico"
       End If
End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores genereados"
  Exit Sub
End Sub

Private Sub cmdImprime_Click()
On Error GoTo ERR:
'--Uniformes
If chkreg_Empleadoras.value = 0 And chkreg_Cancel.value = 0 And chkreg_Tarjetas.value = 0 And chkreg_StockCancel.value = 0 And chkreg_StockTarjetas.value = 0 And chkreg_Ajustes.value = 0 And chkreg_Disp.value = 0 And chkReg_Domicilios.value = 0 _
   And chkUni_Empleadoras.value = 0 And chkUni_Cancel.value = 0 And chkUni_Tarjetas.value = 0 And chkUni_StockCancel.value = 0 And chkUni_StockTarjetas.value = 0 And chkUni_Ajustes.value = 0 And chkUni_Disp.value = 0 And chkUni_Domicilios.value = 0 _
   And chkali_Empleadoras.value = 0 And chkali_Cancel.value = 0 And chkali_Tarjetas.value = 0 And chkali_StockCancel.value = 0 And chkali_StockTarjetas.value = 0 And chkali_Ajustes.value = 0 And chkali_Disp.value = 0 And chkAli_Domicilios.value = 0 _
   And chkpre_Empleadoras.value = 0 And chkpre_Cancel.value = 0 And chkpre_Tarjetas.value = 0 And chkpre_StockCancel.value = 0 And chkpre_StockTarjetas.value = 0 And chkpre_Ajustes.value = 0 And chkpre_Disp.value = 0 And chkPre_Domicilios.value = 0 _
   And chkdie_Empleadoras.value = 0 And chkdie_Cancel.value = 0 And chkdie_Tarjetas.value = 0 And chkdie_StockCancel.value = 0 And chkdie_StockTarjetas.value = 0 And chkdie_Ajustes.value = 0 And chkdie_Disp.value = 0 And chkDie_Domicilios.value = 0 _
   And chkCombustibles = 0 Then
   MsgBox "No tiene seleccionado ningun tipo de archivo para generar." & vbCrLf & "[Seleccione los tipos de archivos que necesita para generarlos]", vbInformation, "No ha seleccionado ningun tipo de archivo"
   Exit Sub
End If

If chkUni_Tarjetas.value = 1 Or chkali_Tarjetas.value = 1 Or chkreg_Tarjetas.value = 1 Or chkpre_Tarjetas.value = 1 Or chkdie_Tarjetas.value = 1 Or _
    chkUni_StockTarjetas.value = 1 Or chkali_StockTarjetas.value = 1 Or chkreg_StockTarjetas.value = 1 Or chkpre_StockTarjetas.value = 1 Or chkdie_StockTarjetas.value = 1 Then
   Call GrabaTarjetas
End If

If chkUni_Ajustes.value = 1 Or chkali_Ajustes.value = 1 Or chkreg_Ajustes.value = 1 Or chkpre_Ajustes.value = 1 Or chkdie_Ajustes.value = 1 Then
   Call GrabaAjustes
End If

If chkUni_Disp.value = 1 Or chkali_Disp.value = 1 Or chkreg_Disp.value = 1 Or chkpre_Disp.value = 1 Or chkdie_Disp.value = 1 Then
   Call GrabaDispersiones
End If

If chkCombustibles.value = 1 Then
    Call GrabaCombustibles
End If

Imprime (crptToWindow)
cmdGrabar.Enabled = True
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub Imprime(Destino)
Dim Result As Integer
'***UNIFORMES
If chkUni_Empleadoras.value = 1 Or chkUni_Domicilios.value = 1 Or chkUni_Cancel.value = 1 Or chkUni_Tarjetas.value = 1 Or chkUni_StockTarjetas.value = 1 Or chkUni_StockCancel.value = 1 Or chkUni_Ajustes.value = 1 Or chkUni_Disp.value = 1 Then
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaProc.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(1) = "1"
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMultipleTrans.rpt" 'ojo con este reporte
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
   
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End If
'***ALIMENTACION
If chkali_Empleadoras.value = 1 Or chkAli_Domicilios.value = 1 Or chkali_Cancel.value = 1 Or chkali_Tarjetas.value = 1 Or chkali_StockCancel.value = 1 Or chkali_StockTarjetas.value = 1 Or chkali_Ajustes.value = 1 Or chkali_Disp.value = 1 Or chkCombustibles.value = 1 Then
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaProc.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(1) = "2"
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMultipleTrans.rpt" 'ojo con este reporte
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino

    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End If
'***REGALO
If chkreg_Empleadoras.value = 1 Or chkReg_Domicilios.value = 1 Or chkreg_Cancel.value = 1 Or chkreg_Tarjetas.value = 1 Or chkreg_StockCancel.value = 1 Or chkreg_StockTarjetas.value = 1 Or chkreg_Ajustes.value = 1 Or chkreg_Disp.value = 1 Or chkCombustibles.value = 1 Then
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaProc.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(1) = "3"
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMultipleTrans.rpt" 'ojo con este reporte
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino

    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End If
'***PREMIUM
If chkpre_Empleadoras.value = 1 Or chkPre_Domicilios.value = 1 Or chkpre_Cancel.value = 1 Or chkpre_Tarjetas.value = 1 Or chkpre_StockCancel.value = 1 Or chkpre_StockTarjetas.value = 1 Or chkpre_Ajustes.value = 1 Or chkpre_Disp.value = 1 Or chkCombustibles.value = 1 Then
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaProc.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(1) = "10"
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMultipleTrans.rpt" 'ojo con este reporte
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino

    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End If
'***DIESEL
If chkdie_Empleadoras.value = 1 Or chkDie_Domicilios.value = 1 Or chkdie_Cancel.value = 1 Or chkdie_Tarjetas.value = 1 Or chkdie_StockCancel.value = 1 Or chkdie_StockTarjetas.value = 1 Or chkdie_Ajustes.value = 1 Or chkdie_Disp.value = 1 Or chkCombustibles.value = 1 Then
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaProc.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(1) = "11"
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMultipleTrans.rpt" 'ojo con este reporte
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino

    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End If

End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Sub CargaDatos()
    Call pon_enblanco_spreads
    Call omite_ceros
    Call CargaDatos_Empleadoras
'    Call CargaDatos_Domicilios
    Call CargaDatosTarjetas
    Call CargaDatosStockTarjetas
    Call CargaBajaTitulares
    Call CargaDatosAjustes
    Call CargaDatosDispersion
    Call marca_opciones
    chkTodos.value = 1
    chkTodos2.value = 1
    chkTodos3.value = 1
    chkTodos4.value = 1
    chkTodos5.value = 1
    cmdGrabar.Enabled = False
End Sub

Private Sub Form_Load()
On Error GoTo ERR:
    Set mclsAniform = New clsAnimated
    SSTab1.Tab = 0
    dir_prueba = gstrPath & "\Prueba"
    cadenafija = "+00000+00000+00000+00000+00000"
    mskFechaProc.Text = Date '+ 1
    CargaDatos
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Sub omite_ceros()
    sqls = "sp_Omiteceros_BE '" & Format(Date, "mm/dd/yyyy") & "',1"
    cnxBD.Execute sqls
End Sub

Sub pon_enblanco_spreads()
    '---Pago Uniformes
    spdClientes.Col = -1
    spdClientes.Row = -1
    spdClientes.Action = 12
    spdClientes.MaxRows = 0
    
    spdAltaTit.Col = -1
    spdAltaTit.Row = -1
    spdAltaTit.Action = 12
    spdAltaTit.MaxRows = 0
    
    spdBajaTit.Col = -1
    spdBajaTit.Row = -1
    spdBajaTit.Action = 12
    spdBajaTit.MaxRows = 0
    
    spdAltaStockTit.Col = -1
    spdAltaStockTit.Row = -1
    spdAltaStockTit.Action = 12
    spdAltaStockTit.MaxRows = 0
    
    spdBajaStockTit.Col = -1
    spdBajaStockTit.Row = -1
    spdBajaStockTit.Action = 12
    spdBajaStockTit.MaxRows = 0
    
    spdAjustes.Col = -1
    spdAjustes.Row = -1
    spdAjustes.Action = 12
    spdAjustes.MaxRows = 0
    
    spdDisp.Col = -1
    spdDisp.Row = -1
    spdDisp.Action = 12
    spdDisp.MaxRows = 0
    '---Pago Alimentos
    spdClientes2.Col = -1
    spdClientes2.Row = -1
    spdClientes2.Action = 12
    spdClientes2.MaxRows = 0
    
    spdAltaTit2.Col = -1
    spdAltaTit2.Row = -1
    spdAltaTit2.Action = 12
    spdAltaTit2.MaxRows = 0
    
    spdBajaTit2.Col = -1
    spdBajaTit2.Row = -1
    spdBajaTit2.Action = 12
    spdBajaTit2.MaxRows = 0
    
    spdAltaStockTit2.Col = -1
    spdAltaStockTit2.Row = -1
    spdAltaStockTit2.Action = 12
    spdAltaStockTit2.MaxRows = 0
    
    spdBajaStockTit2.Col = -1
    spdBajaStockTit2.Row = -1
    spdBajaStockTit2.Action = 12
    spdBajaStockTit2.MaxRows = 0
    
    spdAjustes2.Col = -1
    spdAjustes2.Row = -1
    spdAjustes2.Action = 12
    spdAjustes2.MaxRows = 0
    
    spdDisp2.Col = -1
    spdDisp2.Row = -1
    spdDisp2.Action = 12
    spdDisp2.MaxRows = 0
    
    '---Pago Regalo
    spdClientes3.Col = -1
    spdClientes3.Row = -1
    spdClientes3.Action = 12
    spdClientes3.MaxRows = 0
    
    spdAltaTit3.Col = -1
    spdAltaTit3.Row = -1
    spdAltaTit3.Action = 12
    spdAltaTit3.MaxRows = 0
    
    spdBajaTit3.Col = -1
    spdBajaTit3.Row = -1
    spdBajaTit3.Action = 12
    spdBajaTit3.MaxRows = 0
    
    spdAltaStockTit3.Col = -1
    spdAltaStockTit3.Row = -1
    spdAltaStockTit3.Action = 12
    spdAltaStockTit3.MaxRows = 0
    
    spdBajaStockTit3.Col = -1
    spdBajaStockTit3.Row = -1
    spdBajaStockTit3.Action = 12
    spdBajaStockTit3.MaxRows = 0
    
    spdAjustes3.Col = -1
    spdAjustes3.Row = -1
    spdAjustes3.Action = 12
    spdAjustes3.MaxRows = 0
    
    spdDisp3.Col = -1
    spdDisp3.Row = -1
    spdDisp3.Action = 12
    spdDisp3.MaxRows = 0
    
    
    '---Premium
    spdClientes4.Col = -1
    spdClientes4.Row = -1
    spdClientes4.Action = 12
    spdClientes4.MaxRows = 0
    
    spdAltaTit4.Col = -1
    spdAltaTit4.Row = -1
    spdAltaTit4.Action = 12
    spdAltaTit4.MaxRows = 0
    
    spdBajaTit4.Col = -1
    spdBajaTit4.Row = -1
    spdBajaTit4.Action = 12
    spdBajaTit4.MaxRows = 0
    
    spdAltaStockTit4.Col = -1
    spdAltaStockTit4.Row = -1
    spdAltaStockTit4.Action = 12
    spdAltaStockTit4.MaxRows = 0
    
    spdBajaStockTit4.Col = -1
    spdBajaStockTit4.Row = -1
    spdBajaStockTit4.Action = 12
    spdBajaStockTit4.MaxRows = 0
    
    spdAjustes4.Col = -1
    spdAjustes4.Row = -1
    spdAjustes4.Action = 12
    spdAjustes4.MaxRows = 0
    
    spdDisp4.Col = -1
    spdDisp4.Row = -1
    spdDisp4.Action = 12
    spdDisp4.MaxRows = 0
    
    '---Diesel
    spdClientes5.Col = -1
    spdClientes5.Row = -1
    spdClientes5.Action = 12
    spdClientes5.MaxRows = 0
    
    spdAltaTit5.Col = -1
    spdAltaTit5.Row = -1
    spdAltaTit5.Action = 12
    spdAltaTit5.MaxRows = 0
    
    spdBajaTit5.Col = -1
    spdBajaTit5.Row = -1
    spdBajaTit5.Action = 12
    spdBajaTit5.MaxRows = 0
    
    spdAltaStockTit5.Col = -1
    spdAltaStockTit5.Row = -1
    spdAltaStockTit5.Action = 12
    spdAltaStockTit5.MaxRows = 0
    
    spdBajaStockTit5.Col = -1
    spdBajaStockTit5.Row = -1
    spdBajaStockTit5.Action = 12
    spdBajaStockTit5.MaxRows = 0
    
    spdAjustes5.Col = -1
    spdAjustes5.Row = -1
    spdAjustes5.Action = 12
    spdAjustes5.MaxRows = 0
    
    spdDisp5.Col = -1
    spdDisp5.Row = -1
    spdDisp5.Action = 12
    spdDisp5.MaxRows = 0
    
End Sub

Sub CargaDatos_Empleadoras()
Dim sql2 As String
On Error GoTo ERRO:
'Totales
   Dim total As Double, i As Integer 'contador
   Dim TotEmp As Long
   '----Uniformes
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerCliente',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
    
    '----Alimentacion
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerCliente',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
    
    '----Regalo
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerCliente',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
    
    '----Premium
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerCliente',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
    
    '----Diesel
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerCliente',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing

Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub
Sub CargaDatos_Domicilios()
Dim sql2 As String
On Error GoTo ERRO:
'Solamente cargo domicilios de productos de gasolina para el estado de cuenta
'Totales
   Dim total As Double, i As Integer 'contador
   Dim TotEmp As Long
   
   spdDomAlta1.MaxRows = 0
   
   '----Alta Gas
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerDomiciliosAlta',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdDomAlta2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
       
       '----Cambios Gas
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerDomiciliosCambio',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdDomCambio2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
       
       '----Bajas Gas
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'VerDomiciliosBaja',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdDomBaja2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
   
   '----Alta Viaticos
   spdDomAlta3.MaxRows = 0

Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub

Sub CargaDatosTarjetas()
Dim sql2 As String
Dim total As Double
Dim TotEmp As Long
On Error GoTo ERRO:
    '-Uniformes
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',NULL,'TityAdi',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaTit
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit.Text = TotEmp
    
    '****************************
    '-Alimentacion
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',NULL,'TityAdi',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaTit2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit2.Text = TotEmp
    
    '****************************
    '-Regalo
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',NULL,'TityAdi',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaTit3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit3.Text = TotEmp
    
    '****************************
    '-Premium
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',NULL,'TityAdi',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaTit4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit3.Text = TotEmp
    '****************************
    '-Diesel
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',NULL,'TityAdi',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaTit5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit3.Text = TotEmp
    
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub
Sub CargaDatosStockTarjetas()
Dim sql2 As String
Dim total As Double
Dim TotEmp As Long
On Error GoTo ERRO:
    '-Uniformes
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',NULL,'TityAdiStock',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaStockTit
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit.Text = TotEmp
    
    '****************************
    '-Alimentacion
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',NULL,'TityAdiStock',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaStockTit2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit2.Text = TotEmp
    
    '****************************
    '-Regalo
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',NULL,'TityAdiStock',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaStockTit3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit3.Text = TotEmp
    
    '****************************
    '-Premium
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',NULL,'TityAdiStock',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaStockTit4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
    
'    txtAltaTit3.Text = TotEmp
    '****************************
    '-Diesel
    '-Titulares
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',NULL,'TityAdiStock',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAltaStockTit5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtAltaTit3.Text = TotEmp
    
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub

Sub marca_opciones()
    '-Uniformes
    If spdClientes.MaxRows > 0 Then
       chkUni_Empleadoras.value = 1
    Else
       chkUni_Empleadoras.value = 0
    End If
    If spdDomAlta1.MaxRows > 0 And spdDomAlta1.MaxRows <> 500 Then
       chkUni_Domicilios.value = 1
    Else
      chkUni_Domicilios.value = 0
    End If
    If spdAltaTit.MaxRows > 0 Then
       chkUni_Tarjetas.value = 1
    Else
       chkUni_Tarjetas.value = 0
    End If
    If spdBajaTit.MaxRows > 0 Then
       chkUni_Cancel.value = 1
    Else
       chkUni_Cancel.value = 0
    End If
    If spdAltaStockTit.MaxRows > 0 Then
       chkUni_StockTarjetas.value = 1
    Else
       chkUni_StockTarjetas.value = 0
    End If
    If spdBajaStockTit.MaxRows > 0 Then
       chkUni_StockCancel.value = 1
    Else
       chkUni_StockCancel.value = 0
    End If
    If spdDisp.MaxRows > 0 Then
       chkUni_Disp.value = 1
    Else
       chkUni_Disp.value = 0
    End If
    If spdAjustes.MaxRows > 0 Then
       chkUni_Ajustes.value = 1
    Else
       chkUni_Ajustes.value = 0
    End If
    '-Alimentacion
    If spdClientes2.MaxRows > 0 Then
       chkali_Empleadoras.value = 1
    Else
       chkali_Empleadoras.value = 0
    End If
    If spdDomAlta2.MaxRows > 0 And spdDomAlta2.MaxRows <> 500 Then
       chkAli_Domicilios.value = 1
    Else
       chkAli_Domicilios.value = 0
    End If
    If spdAltaTit2.MaxRows > 0 Then
       chkali_Tarjetas.value = 1
    Else
       chkali_Tarjetas.value = 0
    End If
    If spdBajaTit2.MaxRows > 0 Then
       chkali_Cancel.value = 1
    Else
       chkali_Cancel.value = 0
    End If
    If spdAltaStockTit2.MaxRows > 0 Then
       chkali_StockTarjetas.value = 1
    Else
       chkali_StockTarjetas.value = 0
    End If
    If spdBajaStockTit2.MaxRows > 0 Then
       chkali_StockCancel.value = 1
    Else
       chkali_StockCancel.value = 0
    End If
    If spdDisp2.MaxRows > 0 Then
       chkali_Disp.value = 1
    Else
       chkali_Disp.value = 0
    End If
    If spdAjustes2.MaxRows > 0 Then
       chkali_Ajustes.value = 1
    Else
       chkali_Ajustes.value = 0
    End If
    '-Regalo
    If spdClientes3.MaxRows > 0 Then
       chkreg_Empleadoras.value = 1
    Else
       chkreg_Empleadoras.value = 0
    End If
    If spdDomAlta3.MaxRows > 0 And spdDomAlta3.MaxRows <> 500 Then
       chkReg_Domicilios.value = 1
    Else
       chkReg_Domicilios.value = 0
    End If
    If spdAltaTit3.MaxRows > 0 Then
       chkreg_Tarjetas.value = 1
    Else
       chkreg_Tarjetas.value = 0
    End If
    If spdBajaTit3.MaxRows > 0 Then
       chkreg_Cancel.value = 1
    Else
       chkreg_Cancel.value = 0
    End If
    If spdAltaStockTit3.MaxRows > 0 Then
       chkreg_StockTarjetas.value = 1
    Else
       chkreg_StockTarjetas.value = 0
    End If
    If spdBajaStockTit3.MaxRows > 0 Then
       chkreg_StockCancel.value = 1
    Else
       chkreg_StockCancel.value = 0
    End If
    If spdDisp3.MaxRows > 0 Then
       chkreg_Disp.value = 1
    Else
       chkreg_Disp.value = 0
    End If
    If spdAjustes3.MaxRows > 0 Then
       chkreg_Ajustes.value = 1
    Else
       chkreg_Ajustes.value = 0
    End If
    '-Premium
    If spdClientes4.MaxRows > 0 Then
       chkpre_Empleadoras.value = 1
    Else
       chkpre_Empleadoras.value = 0
    End If
    If spdDomAlta4.MaxRows > 0 And spdDomAlta4.MaxRows <> 500 Then
       chkPre_Domicilios.value = 1
    Else
       chkPre_Domicilios.value = 0
    End If
    If spdAltaTit4.MaxRows > 0 Then
       chkpre_Tarjetas.value = 1
    Else
       chkpre_Tarjetas.value = 0
    End If
    If spdBajaTit4.MaxRows > 0 Then
       chkpre_Cancel.value = 1
    Else
       chkpre_Cancel.value = 0
    End If
    If spdAltaStockTit4.MaxRows > 0 Then
       chkpre_StockTarjetas.value = 1
    Else
       chkpre_StockTarjetas.value = 0
    End If
    If spdBajaStockTit4.MaxRows > 0 Then
       chkpre_StockCancel.value = 1
    Else
       chkpre_StockCancel.value = 0
    End If
    If spdDisp4.MaxRows > 0 Then
       chkpre_Disp.value = 1
    Else
       chkpre_Disp.value = 0
    End If
    If spdAjustes4.MaxRows > 0 Then
       chkpre_Ajustes.value = 1
    Else
       chkpre_Ajustes.value = 0
    End If
    '-Diesel
    If spdClientes5.MaxRows > 0 Then
       chkdie_Empleadoras.value = 1
    Else
       chkdie_Empleadoras.value = 0
    End If
    If spdDomAlta5.MaxRows > 0 And spdDomAlta5.MaxRows <> 500 Then
       chkDie_Domicilios.value = 1
    Else
       chkDie_Domicilios.value = 0
    End If
    If spdAltaTit5.MaxRows > 0 Then
       chkdie_Tarjetas.value = 1
    Else
       chkdie_Tarjetas.value = 0
    End If
    If spdBajaTit5.MaxRows > 0 Then
       chkdie_Cancel.value = 1
    Else
       chkdie_Cancel.value = 0
    End If
    If spdAltaStockTit5.MaxRows > 0 Then
       chkdie_StockTarjetas.value = 1
    Else
       chkdie_StockTarjetas.value = 0
    End If
    If spdBajaStockTit5.MaxRows > 0 Then
       chkdie_StockCancel.value = 1
    Else
       chkdie_StockCancel.value = 0
    End If
    If spdDisp5.MaxRows > 0 Then
       chkdie_Disp.value = 1
    Else
       chkdie_Disp.value = 0
    End If
    If spdAjustes5.MaxRows > 0 Then
       chkdie_Ajustes.value = 1
    Else
       chkdie_Ajustes.value = 0
    End If
End Sub

Sub GeneraArchivoAltaEmpleadora()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal, reg90 As Integer, reg91 As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
Dim CP

'nfile = FreeFile
If chkUni_Empleadoras.value = 1 Or chkali_Empleadoras.value = 1 Or chkreg_Empleadoras.value = 1 Or chkpre_Empleadoras.value = 1 Or chkdie_Empleadoras.value = 1 Then
Open dir_prueba & "AEMP" & Format(CDate(mskFechaProc.Text), "mmdd") & "_SC.vlt" For Output As #1
'Header A
clinea = "10COMPANY04" & Format(CDate(mskFechaProc.Text), "YYYYMMDD") & rellena("0", 251, "")
Print #1, clinea
'Header B
reg91 = 0
With spdClientes  'UNIFORMES
 reg90 = 0
 For i = 1 To .MaxRows
       clinea = "20"
       .Row = i
       .Col = 3
       If .TypeComboBoxCurSel = 0 Then
           ntransa = 1
       End If
       If .TypeComboBoxCurSel = 1 Then
           ntransa = 2
       End If
       If .TypeComboBoxCurSel = 2 Then
           ntransa = 3
       End If
       If .TypeComboBoxCurSel = -1 Then
           ntransa = 1
       End If
       transa = rellena(CStr(ntransa), 2, "0", "D") ' 01 alta, 02 Baja, 03 modificacion
       
       If transa = "" Or transa = Null Or transa = "-1" Then
          transa = "01"
       End If
       clinea = clinea & transa
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       .Col = 2
       elgrupo = Mid(UCase(.Text), 1, 100)
       clinea = clinea & rellena(CStr(cliente), 10, " ", "I") & rellena(elgrupo, 100, " ", "I")
       clinea = clinea & "01VTDESPENSA" & cadenafija & rellena("0", 114, "")   'porke es pagouniformes
       Print #1, clinea
       '--Detail Uniformes

       sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ", 'DetDomEnt',1"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
        'Detalle
        reg90 = 0
    
         Do While Not rsBD.EOF
            reg90 = reg90 + 1
            clinea = "21" & rsBD!Clave
            clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
            clinea = clinea & rellena(CStr(rsBD!descripcion), 60, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Direccion & " " & rsBD!Colonia, 68), 68, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Ciudad, 20), 20, " ", "I")
            clinea = clinea & rellena(Left(rsBD!estado, 3), 3, " ", "I") & Format(rsBD!CodigoPostal, "00000")
            clinea = clinea & rellena(Left(rsBD!Telefono, 12), 12, " ", "I")
            clinea = clinea & rellena(Left(rsBD!contacto, 26), 26, " ", "I")
            clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
            Print #1, clinea
            rsBD.MoveNext
        Loop

 Next i
 reg91 = reg91 + reg90
 clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
Else
  Exit Sub
End If

'HEADER B
If chkali_Empleadoras.value = 1 Then
With spdClientes2     'ALIMENTOS
 reg90 = 0
 For i = 1 To .MaxRows
       clinea = "20"
       .Row = i
       .Col = 3
       If .TypeComboBoxCurSel = 0 Then
           ntransa = 1
       End If
       If .TypeComboBoxCurSel = 1 Then
           ntransa = 2
       End If
       If .TypeComboBoxCurSel = 2 Then
           ntransa = 3
       End If
       If .TypeComboBoxCurSel = -1 Then
           ntransa = 1
       End If
       transa = rellena(CStr(ntransa), 2, "0", "D") ' 01 alta, 02 Baja, 03 modificacion
       
       If transa = "" Or transa = Null Or transa = "-1" Then
          transa = "01"
       End If
       clinea = clinea & transa
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       .Col = 2
       elgrupo = Mid(UCase(.Text), 1, 100)
       clinea = clinea & rellena(CStr(cliente), 10, " ", "I") & rellena(elgrupo, 100, " ", "I")
       clinea = clinea & "02VTCOMBUSTI" & cadenafija & rellena("0", 114, "")  'porke es pagoALIMENTOS
       Print #1, clinea
       '---Detail Alimentos
       sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ", 'DetDomEnt',2"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
        'Detalle
        reg90 = 0
    
         Do While Not rsBD.EOF
            reg90 = reg90 + 1
            clinea = "21" & rsBD!Clave
            clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
            clinea = clinea & rellena(CStr(rsBD!descripcion), 60, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Direccion & " " & rsBD!Colonia, 68), 68, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Ciudad, 20), 20, " ", "I")
            clinea = clinea & rellena(Left(rsBD!estado, 3), 3, " ", "I") & Format(rsBD!CodigoPostal, "00000")
            clinea = clinea & rellena(Left(rsBD!Telefono, 12), 12, " ", "I")
            clinea = clinea & rellena(Left(rsBD!contacto, 26), 26, " ", "I")
            clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
            Print #1, clinea
            rsBD.MoveNext
        Loop
 Next i
 reg91 = reg91 + reg90
 clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'HEADER B
If chkreg_Empleadoras.value = 1 Then
With spdClientes3     'REGALO
 reg90 = 0
 For i = 1 To .MaxRows
       clinea = "20"
       .Row = i
       .Col = 3
       If .TypeComboBoxCurSel = 0 Then
           ntransa = 1
       End If
       If .TypeComboBoxCurSel = 1 Then
           ntransa = 2
       End If
       If .TypeComboBoxCurSel = 2 Then
           ntransa = 3
       End If
       If .TypeComboBoxCurSel = -1 Then
           ntransa = 1
       End If
       transa = rellena(CStr(ntransa), 2, "0", "D") ' 01 alta, 02 Baja, 03 modificacion
       
       If transa = "" Or transa = Null Or transa = "-1" Then
          transa = "01"
       End If
       clinea = clinea & transa
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       .Col = 2
       elgrupo = Mid(UCase(.Text), 1, 100)
       clinea = clinea & rellena(CStr(cliente), 10, " ", "I") & rellena(elgrupo, 100, " ", "I")
       clinea = clinea & "03VTMARTABIE" & cadenafija & rellena("0", 114, "")  'porke es pago REGALO
       Print #1, clinea
       '--Detail Regalo
       sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ", 'DetDomEnt',3"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
        'Detalle
        reg90 = 0
    
         Do While Not rsBD.EOF
            reg90 = reg90 + 1
            clinea = "21" & rsBD!Clave
            clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
            clinea = clinea & rellena(CStr(rsBD!descripcion), 60, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Direccion & " " & rsBD!Colonia, 68), 68, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Ciudad, 20), 20, " ", "I")
            clinea = clinea & rellena(Left(rsBD!estado, 3), 3, " ", "I") & Format(rsBD!CodigoPostal, "00000")
            clinea = clinea & rellena(Left(rsBD!Telefono, 12), 12, " ", "I")
            clinea = clinea & rellena(Left(rsBD!contacto, 26), 26, " ", "I")
            clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
            Print #1, clinea
            rsBD.MoveNext
        Loop
 Next i
 reg91 = reg91 + reg90
 'Trailer B
 clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'HEADER B
If chkpre_Empleadoras.value = 1 Then
With spdClientes4     'PREMIUM
 reg90 = 0
 For i = 1 To .MaxRows
       clinea = "20"
       .Row = i
       .Col = 3
       If .TypeComboBoxCurSel = 0 Then
           ntransa = 1
       End If
       If .TypeComboBoxCurSel = 1 Then
           ntransa = 2
       End If
       If .TypeComboBoxCurSel = 2 Then
           ntransa = 3
       End If
       If .TypeComboBoxCurSel = -1 Then
           ntransa = 1
       End If
       transa = rellena(CStr(ntransa), 2, "0", "D") ' 01 alta, 02 Baja, 03 modificacion
       
       If transa = "" Or transa = Null Or transa = "-1" Then
          transa = "01"
       End If
       clinea = clinea & transa
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       .Col = 2
       elgrupo = Mid(UCase(.Text), 1, 100)
       clinea = clinea & rellena(CStr(cliente), 10, " ", "I") & rellena(elgrupo, 100, " ", "I")
       clinea = clinea & "10VTPREMIUM " & cadenafija & rellena("0", 114, "")  'porke es pago REGALO
       Print #1, clinea
       '--Detail Regalo
       sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ", 'DetDomEnt',10"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
        'Detalle
        reg90 = 0
    
         Do While Not rsBD.EOF
            reg90 = reg90 + 1
            clinea = "21" & rsBD!Clave
            clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
            clinea = clinea & rellena(CStr(rsBD!descripcion), 60, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Direccion & " " & rsBD!Colonia, 68), 68, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Ciudad, 20), 20, " ", "I")
            clinea = clinea & rellena(Left(rsBD!estado, 3), 3, " ", "I") & Format(rsBD!CodigoPostal, "00000")
            clinea = clinea & rellena(Left(rsBD!Telefono, 12), 12, " ", "I")
            clinea = clinea & rellena(Left(rsBD!contacto, 26), 26, " ", "I")
            clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
            Print #1, clinea
            rsBD.MoveNext
        Loop
 Next i
 reg91 = reg91 + reg90
 'Trailer B
 clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'HEADER B
If chkdie_Empleadoras.value = 1 Then
With spdClientes5     'DIESEL
 reg90 = 0
 For i = 1 To .MaxRows
       clinea = "20"
       .Row = i
       .Col = 3
       If .TypeComboBoxCurSel = 0 Then
           ntransa = 1
       End If
       If .TypeComboBoxCurSel = 1 Then
           ntransa = 2
       End If
       If .TypeComboBoxCurSel = 2 Then
           ntransa = 3
       End If
       If .TypeComboBoxCurSel = -1 Then
           ntransa = 1
       End If
       transa = rellena(CStr(ntransa), 2, "0", "D") ' 01 alta, 02 Baja, 03 modificacion
       
       If transa = "" Or transa = Null Or transa = "-1" Then
          transa = "01"
       End If
       clinea = clinea & transa
       .Row = i
       .Col = 1
       cliente = Val(.Text)
       .Col = 2
       elgrupo = Mid(UCase(.Text), 1, 100)
       clinea = clinea & rellena(CStr(cliente), 10, " ", "I") & rellena(elgrupo, 100, " ", "I")
       clinea = clinea & "11VTDIESEL  " & cadenafija & rellena("0", 114, "")  'porke es pago REGALO
       Print #1, clinea
       '--Detail Regalo
       sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ", 'DetDomEnt',11"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
        'Detalle
        reg90 = 0
    
         Do While Not rsBD.EOF
            reg90 = reg90 + 1
            clinea = "21" & rsBD!Clave
            clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
            clinea = clinea & rellena(CStr(rsBD!descripcion), 60, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Direccion & " " & rsBD!Colonia, 68), 68, " ", "I")
            clinea = clinea & rellena(Left(rsBD!Ciudad, 20), 20, " ", "I")
            clinea = clinea & rellena(Left(rsBD!estado, 3), 3, " ", "I") & Format(rsBD!CodigoPostal, "00000")
            clinea = clinea & rellena(Left(rsBD!Telefono, 12), 12, " ", "I")
            clinea = clinea & rellena(Left(rsBD!contacto, 26), 26, " ", "I")
            clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
            Print #1, clinea
            rsBD.MoveNext
        Loop
 Next i
 reg91 = reg91 + reg90
 'Trailer B
 clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If


'Trailer A
clinea = "91" & Format(reg91, "000000") & rellena("0", 262, "")
Print #1, clinea

Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de empleadoras"
   Exit Sub
End Sub
Sub GeneraArchivoDomicilios()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal, reg90 As Integer, reg91 As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
Dim CP

'nfile = FreeFile
'If chkUni_Domicilios.value = 1 Or chkAli_Domicilios.value = 1 Or chkReg_Domicilios.value = 1 Then
    ' Creamos Domicilios de Entrega
    Open dir_prueba & "MDF" & Format(CDate(mskFechaProc.Text), "yymmdd") & "_SC.vlt" For Output As #1
    'Header A
    clinea = "11" & Format(CDate(mskFechaProc.Text), "YYYYMMDD") & "03" & rellena("0", 521, "")
    Print #1, clinea
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "', 0, 'DetDomEnt2',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
    'Detalle
    reg91 = 0
    
     Do While Not rsBD.EOF
         reg91 = reg91 + 1
        clinea = "22" & Format(rsBD!Clave, "00")
        clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
        clinea = clinea & rellena(CStr(rsBD!Nombre), 150, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!Rfc), 13, " ", "I")
        clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
        clinea = clinea & rellena(CStr(rsBD!Direccion), 90, " ", "I") & rellena(CStr(rsBD!Colonia), 90, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!Ciudad), 50, " ", "I") & rellena(CStr(rsBD!Ciudad), 50, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!estado), 50, " ", "I") & Format(rsBD!CodigoPostal, "00000")
        clinea = clinea & "01" & rellena("0", 9, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
           
    'Trailer A
    clinea = "91" & Format(reg91, "00000000") & rellena("0", 523, "")
    Print #1, clinea
    
    Close #1
    rsBD.Close
    Set rsBD = Nothing
'End If
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de Archivo Domicilios"
   Exit Sub
End Sub

Sub GeneraArchivoTarjetas()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre As String, ap As String, am As String, elnombre As String
Dim Empleadora As String
Dim RetVal, reg93 As Integer, reg91 As Integer, reg92 As Integer, nemp As Integer, reg922 As Integer

Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
reg92 = 0
reg922 = 0
nemp = 0
If (chkUni_Tarjetas.value = 1 And spdAltaTit.MaxRows > 0) Or _
   (chkali_Tarjetas.value = 1 And spdAltaTit2.MaxRows > 0) Or _
   (chkreg_Tarjetas.value = 1 And spdAltaTit3.MaxRows > 0) Or _
   (chkpre_Tarjetas.value = 1 And spdAltaTit4.MaxRows > 0) Or _
   (chkdie_Tarjetas.value = 1 And spdAltaTit5.MaxRows > 0) Then
   Open dir_prueba & "TAT" & Format(CDate(mskFechaProc.Text), "yymmdd") & "_SC.vlt" For Output As #1
   'Header A
   clinea = "11EMP-REG04" & Format(CDate(mskFechaProc.Text), "YYYYMMDD") & rellena("0", 53, "")
   Print #1, clinea
Else
   Exit Sub
End If

If chkUni_Tarjetas.value = 1 And spdAltaTit.MaxRows > 0 Then
'Header B Uniformes
   reg93 = 0
   clinea = "1301" & rellena("0", 68, "") 'porke es Uniformes
   Print #1, clinea
   With spdAltaTit
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ",'ArchivoAmbosTipos',1"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 60, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Uniformes Titulares
               clinea = "21"
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "01"
               Else
                  clinea = clinea & "02"
               End If
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               ap = IIf(IsNull(rsBD!apat), "", Trim(rsBD!apat))
               ap = Trim(ap)
               am = IIf(IsNull(rsBD!amat), "", Trim(rsBD!amat))
               am = Trim(am)
               elnombre = Trim(rsBD!Nombre) & " " & ap & " " & am
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I")
               clinea = clinea & rellena("0", 8 - Len(rsBD!Cuenta), "")
               clinea = clinea & rsBD!Cuenta
               clinea = clinea & rellena("0", 24, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C UNIFORMES
            clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B UNIFORMES
   clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If
'++++++Alimentacion
'Header B Alimentacion
If chkali_Tarjetas.value = 1 And spdAltaTit2.MaxRows > 0 Then
   reg93 = 0
   clinea = "1302" & rellena("0", 68, "") 'porke es Alimentacion
   Print #1, clinea
   With spdAltaTit2
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ",'ArchivoAmbosTipos',2"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 60, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
               clinea = "21"
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "01"
               Else
                  clinea = clinea & "02"
               End If
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               ap = IIf(IsNull(rsBD!apat), "", Trim(rsBD!apat))
               ap = Trim(ap)
               am = IIf(IsNull(rsBD!amat), "", Trim(rsBD!amat))
               am = Trim(am)
               elnombre = Trim(rsBD!Nombre) & " " & ap & " " & am
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I")
               clinea = clinea & rellena("0", 8 - Len(rsBD!Cuenta), "")
               clinea = clinea & rsBD!Cuenta
               clinea = clinea & rellena("0", 24, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C Alimentacion
            clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
            Print #1, clinea
       Next
   End With
   'Trailer B Alimentacion
   clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Regalo
'Header B Regalo
If chkreg_Tarjetas.value = 1 And spdAltaTit3.MaxRows > 0 Then
   reg93 = 0
   clinea = "1303" & rellena("0", 68, "") 'porke es Regalo
   Print #1, clinea
   With spdAltaTit3
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ",'ArchivoAmbosTipos',3"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 60, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
               clinea = "21"
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "01"
               Else
                  clinea = clinea & "02"
               End If
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               ap = IIf(IsNull(rsBD!apat), "", Trim(rsBD!apat))
               ap = Trim(ap)
               am = IIf(IsNull(rsBD!amat), "", Trim(rsBD!amat))
               am = Trim(am)
               elnombre = Trim(rsBD!Nombre) & " " & ap & " " & am
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I")
               clinea = clinea & rellena("0", 8 - Len(rsBD!Cuenta), "")
               clinea = clinea & rsBD!Cuenta
               clinea = clinea & rellena("0", 24, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C Regalo
            clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
            Print #1, clinea
       Next
   End With
   'Trailer B Regalo
   clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Premium
'Header B premium
If chkpre_Tarjetas.value = 1 And spdAltaTit4.MaxRows > 0 Then
   reg93 = 0
   clinea = "1310" & rellena("0", 68, "") 'porke es premium
   Print #1, clinea
   With spdAltaTit4
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ",'ArchivoAmbosTipos',10"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 60, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
               clinea = "21"
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "01"
               Else
                  clinea = clinea & "02"
               End If
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               ap = IIf(IsNull(rsBD!apat), "", Trim(rsBD!apat))
               ap = Trim(ap)
               am = IIf(IsNull(rsBD!amat), "", Trim(rsBD!amat))
               am = Trim(am)
               elnombre = Trim(rsBD!Nombre) & " " & ap & " " & am
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I")
               clinea = clinea & rellena("0", 8 - Len(rsBD!Cuenta), "")
               clinea = clinea & rsBD!Cuenta
               clinea = clinea & rellena("0", 24, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C premium
            clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
            Print #1, clinea
       Next
   End With
   'Trailer B premium
   clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++diesel
'Header B diesel
If chkdie_Tarjetas.value = 1 And spdAltaTit5.MaxRows > 0 Then
   reg93 = 0
   clinea = "1311" & rellena("0", 68, "") 'porke es diesel
   Print #1, clinea
   With spdAltaTit5
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'," & cliente & ",'ArchivoAmbosTipos',11"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 60, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
               clinea = "21"
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "01"
               Else
                  clinea = clinea & "02"
               End If
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               ap = IIf(IsNull(rsBD!apat), "", Trim(rsBD!apat))
               ap = Trim(ap)
               am = IIf(IsNull(rsBD!amat), "", Trim(rsBD!amat))
               am = Trim(am)
               elnombre = Trim(rsBD!Nombre) & " " & ap & " " & am
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I")
               clinea = clinea & rellena("0", 8 - Len(rsBD!Cuenta), "")
               clinea = clinea & rsBD!Cuenta
               clinea = clinea & rellena("0", 24, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C diesel
            clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
            Print #1, clinea
       Next
   End With
   'Trailer B diesel
   clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If


'Trailer A
clinea = "92" & Format(reg922, "000000") & Format(reg92, "000000") & rellena("0", 58, "")
Print #1, clinea

Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de titulares"
   Exit Sub
End Sub
Sub GeneraArchivoStockTarjetas()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre As String, ap As String, am As String, elnombre As String
Dim Empleadora As String
Dim RetVal, reg93 As Integer, reg91 As Integer, reg92 As Integer, nemp As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
Dim sArchivo As String
Dim consArch As Integer

EmpleadosClienteGlobal = 0
reg92 = 0
nemp = 0

If chkUni_StockTarjetas.value = 1 And spdAltaStockTit.MaxRows > 0 Then
'Header B Uniformes
    consArch = 1
    sArchivo = Dir(dir_prueba & "CAF50640601" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "*.txt")
    Do While sArchivo <> ""
        consArch = consArch + 1
        sArchivo = Dir
    Loop
   Open dir_prueba & "CAF50640601" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & ".EMBSYC-" & IIf(consArch < 10, "0" & consArch, consArch) & ".txt" For Output As #1

   With spdAltaStockTit
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           'Detail Uniformes Titulares
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoStockAmbosTip',1"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
               clinea = "1"
               clinea = clinea & rellena(rsBD!Bin, 19, " ", "I")
               clinea = clinea & "510000"
               clinea = clinea & rellena(rsBD!Cuenta, 19, " ", "I")
               clinea = clinea & "0450"
               clinea = clinea & rellena(rsBD!Nombre, 35, " ", "I")
               clinea = clinea & rellena(Left(rsBD!Nombre, 24), 24, " ", "I")
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "0"
               Else
                  clinea = clinea & "1"
               End If
               clinea = clinea & rellena(rsBD!Direccion, 90, " ", "I")
               clinea = clinea & rellena(rsBD!Ciudad, 45, " ", "I")
               clinea = clinea & rellena(rsBD!estado, 45, " ", "I")
               clinea = clinea & rellena(rsBD!CP, 5, "0", "D")
               clinea = clinea & rellena(rsBD!Telefono, 20, " ", "I")
               clinea = clinea & rellena(rsBD!Plaza, 3, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & "000000000000"
               clinea = clinea & Left(Replace(rsBD!Rfc, "-", ""), 13)
               clinea = clinea & rellena("", 51, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & rellena("", 32, " ", "I")
               clinea = clinea & rellena(rsBD!nombrecli, 30, " ", "I")
               Print #1, clinea
               rsBD.MoveNext
           Loop
       Next
   End With
    Close #1
    rsBD.Close
    Set rsBD = Nothing
End If
'++++++Alimentacion
'Header B Alimentacion
If chkali_StockTarjetas.value = 1 And spdAltaStockTit2.MaxRows > 0 Then
    consArch = 1
    sArchivo = Dir(dir_prueba & "CAF50640501" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "*.txt")
    Do While sArchivo <> ""
        consArch = consArch + 1
        sArchivo = Dir
    Loop
   Open dir_prueba & "CAF50640501" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & ".EMBSYC-" & IIf(consArch < 10, "0" & consArch, consArch) & ".txt" For Output As #1

   With spdAltaStockTit2
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           'Detail Alimentacion Titulares
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoStockAmbosTip',2"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
               clinea = "1"
               clinea = clinea & rellena(rsBD!Bin, 19, " ", "I")
               clinea = clinea & "510000"
               clinea = clinea & rellena(rsBD!Cuenta, 19, " ", "I")
               clinea = clinea & "0450"
               clinea = clinea & rellena(rsBD!Nombre, 35, " ", "I")
               clinea = clinea & rellena(Left(rsBD!Nombre, 24), 24, " ", "I")
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "0"
               Else
                  clinea = clinea & "1"
               End If
               clinea = clinea & rellena(rsBD!Direccion, 90, " ", "I")
               clinea = clinea & rellena(rsBD!Ciudad, 45, " ", "I")
               clinea = clinea & rellena(rsBD!estado, 45, " ", "I")
               clinea = clinea & rellena(rsBD!CP, 5, "0", "D")
               clinea = clinea & rellena(rsBD!Telefono, 20, " ", "I")
               clinea = clinea & rellena(rsBD!Plaza, 3, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & "000000000000"
               clinea = clinea & Left(Replace(rsBD!Rfc, "-", ""), 13)
               clinea = clinea & rellena("", 51, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & rellena("", 32, " ", "I")
               clinea = clinea & rellena(rsBD!nombrecli, 30, " ", "I")
               Print #1, clinea
               rsBD.MoveNext
           Loop
       Next
   End With
    Close #1
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Regalo
'Header B Regalo
If chkreg_StockTarjetas.value = 1 And spdAltaStockTit3.MaxRows > 0 Then
    consArch = 1
    sArchivo = Dir(dir_prueba & "CAF50640602" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "*.txt")
    Do While sArchivo <> ""
        consArch = consArch + 1
        sArchivo = Dir
    Loop
   Open dir_prueba & "CAF50640602" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & ".EMBSYC-" & IIf(consArch < 10, "0" & consArch, consArch) & ".txt" For Output As #1
   
   
   With spdAltaStockTit3
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           'Detail Regalo Titulares
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoStockAmbosTip',3"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
               clinea = "1"
               clinea = clinea & rellena(rsBD!Bin, 19, " ", "I")
               clinea = clinea & "510000"
               clinea = clinea & rellena(rsBD!Cuenta, 19, " ", "I")
               clinea = clinea & "0450"
               clinea = clinea & rellena(rsBD!Nombre, 35, " ", "I")
               clinea = clinea & rellena(Left(rsBD!Nombre, 24), 24, " ", "I")
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "0"
               Else
                  clinea = clinea & "1"
               End If
               clinea = clinea & rellena(rsBD!Direccion, 90, " ", "I")
               clinea = clinea & rellena(rsBD!Ciudad, 45, " ", "I")
               clinea = clinea & rellena(rsBD!estado, 45, " ", "I")
               clinea = clinea & rellena(rsBD!CP, 5, "0", "D")
               clinea = clinea & rellena(rsBD!Telefono, 20, " ", "I")
               clinea = clinea & rellena(rsBD!Plaza, 3, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & "000000000000"
               clinea = clinea & Left(Replace(rsBD!Rfc, "-", ""), 13)
               clinea = clinea & rellena("", 51, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & rellena("", 32, " ", "I")
               clinea = clinea & rellena(rsBD!nombrecli, 30, " ", "I")
               Print #1, clinea
               rsBD.MoveNext
           Loop
       Next
   End With
    Close #1
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Premium
'Header B Premium
If chkpre_StockTarjetas.value = 1 And spdAltaStockTit4.MaxRows > 0 Then
    consArch = 1
    sArchivo = Dir(dir_prueba & "CAF50640502" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "*.txt")
    Do While sArchivo <> ""
        consArch = consArch + 1
        sArchivo = Dir
    Loop
   Open dir_prueba & "CAF50640502" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & ".EMBSYC-" & IIf(consArch < 10, "0" & consArch, consArch) & ".txt" For Output As #1
   
   
   With spdAltaStockTit4
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           'Detail Premium Titulares
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoStockAmbosTip',10"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
               clinea = "1"
               clinea = clinea & rellena(rsBD!Bin, 19, " ", "I")
               clinea = clinea & "510000"
               clinea = clinea & rellena(rsBD!Cuenta, 19, " ", "I")
               clinea = clinea & "0450"
               clinea = clinea & rellena(rsBD!Nombre, 35, " ", "I")
               clinea = clinea & rellena(Left(rsBD!Nombre, 24), 24, " ", "I")
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "0"
               Else
                  clinea = clinea & "1"
               End If
               clinea = clinea & rellena(rsBD!Direccion, 90, " ", "I")
               clinea = clinea & rellena(rsBD!Ciudad, 45, " ", "I")
               clinea = clinea & rellena(rsBD!estado, 45, " ", "I")
               clinea = clinea & rellena(rsBD!CP, 5, "0", "D")
               clinea = clinea & rellena(rsBD!Telefono, 20, " ", "I")
               clinea = clinea & rellena(rsBD!Plaza, 3, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & "000000000000"
               clinea = clinea & Left(Replace(rsBD!Rfc, "-", ""), 13)
               clinea = clinea & rellena("", 51, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & rellena("", 32, " ", "I")
               clinea = clinea & rellena(rsBD!nombrecli, 30, " ", "I")
               Print #1, clinea
               rsBD.MoveNext
           Loop
       Next
   End With
    Close #1
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Diesel
'Header B Diesel
If chkdie_StockTarjetas.value = 1 And spdAltaStockTit5.MaxRows > 0 Then
    consArch = 1
    sArchivo = Dir(dir_prueba & "CAF50640503" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "*.txt")
    Do While sArchivo <> ""
        consArch = consArch + 1
        sArchivo = Dir
    Loop
   Open dir_prueba & "CAF50640503" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & ".EMBSYC-" & IIf(consArch < 10, "0" & consArch, consArch) & ".txt" For Output As #1
   
   
   With spdAltaStockTit5
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           'Detail Diesel Titulares
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoStockAmbosTip',11"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
               clinea = "1"
               clinea = clinea & rellena(rsBD!Bin, 19, " ", "I")
               clinea = clinea & "510000"
               clinea = clinea & rellena(rsBD!Cuenta, 19, " ", "I")
               clinea = clinea & "0450"
               clinea = clinea & rellena(rsBD!Nombre, 35, " ", "I")
               clinea = clinea & rellena(Left(rsBD!Nombre, 24), 24, " ", "I")
               If Trim(rsBD!tipo) = "T" Then
                  clinea = clinea & "0"
               Else
                  clinea = clinea & "1"
               End If
               clinea = clinea & rellena(rsBD!Direccion, 90, " ", "I")
               clinea = clinea & rellena(rsBD!Ciudad, 45, " ", "I")
               clinea = clinea & rellena(rsBD!estado, 45, " ", "I")
               clinea = clinea & rellena(rsBD!CP, 5, "0", "D")
               clinea = clinea & rellena(rsBD!Telefono, 20, " ", "I")
               clinea = clinea & rellena(rsBD!Plaza, 3, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & "000000000000"
               clinea = clinea & Left(Replace(rsBD!Rfc, "-", ""), 13)
               clinea = clinea & rellena("", 51, " ", "I")
               clinea = clinea & "01"
               clinea = clinea & rellena("", 32, " ", "I")
               clinea = clinea & rellena(rsBD!nombrecli, 30, " ", "I")
               Print #1, clinea
               rsBD.MoveNext
           Loop
       Next
   End With
    Close #1
    rsBD.Close
    Set rsBD = Nothing
End If

Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de titulares"
   Exit Sub
End Sub
Sub GeneraArchivoCombustibles()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal, reg90 As Integer, reg91 As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
Dim CP

    ' Creamos Domicilios de Entrega
    Open dir_prueba & "IEDOCTA" & Format(CDate(mskFechaProc.Text), "yymmdd") & "_SC.vlt" For Output As #1
    'Header A
    clinea = "11" & Format(CDate(mskFechaProc.Text), "YYYYMMDD") & "03" & rellena("0", 50, "")
    Print #1, clinea
    
    reg91 = 0
    
    'Detalle Magna
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "', 0, 'Combustible',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Detalle Premium
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "', 0, 'Combustible',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Detalle Diesel
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "', 0, 'Combustible',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Trailer A
    clinea = "91" & Format(reg91, "00000000") & rellena("0", 52, "")
    Print #1, clinea
    
    Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Archivo Combustibles"
   Exit Sub
End Sub

Sub dame_empleadoras_tarjetas()
Dim i As Integer, j As Integer, resto As Integer
resto = 0
For i = 1 To spdAltaTit.MaxRows
    For j = 1 To spdAltaAd.MaxRows
        spdAltaTit.Col = 1
        spdAltaTit.Row = i
        spdAltaAd.Col = 1
        spdAltaAd.Row = j
        If Val(spdAltaTit.Text) = Val(spdAltaAd.Text) Then
            resto = resto + 1
        End If
    Next
Next
nemp_uni = (spdAltaTit.MaxRows + spdAltaAd.MaxRows) - resto
End Sub

Sub CargaBajaTitulares()
Dim total As Double
Dim TotEmp As Long
On Error GoTo ERRO:
'***Uniformes
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',Null,'BajaTitulares',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdBajaTit
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
'    txtBajaTit.Text = TotEmp
'***Alimentos
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',Null,'BajaTitulares',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdBajaTit2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
'    txtBajaTit2.Text = TotEmp
    rsBD.Close
    Set rsBD = Nothing

'***Regalo
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',Null,'BajaTitulares',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdBajaTit3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
'    txtBajaTit3.Text = TotEmp
    rsBD.Close
    Set rsBD = Nothing

'***Premium
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',Null,'BajaTitulares',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdBajaTit4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
'    txtBajaTit3.Text = TotEmp
    rsBD.Close
    Set rsBD = Nothing

'***Diesel
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "',Null,'BajaTitulares',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdBajaTit5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    TotEmp = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        .Col = 3
        .Text = rsBD!NumEmpl
        TotEmp = TotEmp + rsBD!NumEmpl
        rsBD.MoveNext
    Loop
    End With
'    txtBajaTit3.Text = TotEmp
    rsBD.Close
    Set rsBD = Nothing

Exit Sub
ERRO:
  MsgBox "Error al generar archivo de cancelaciones", vbCritical, "Errores encontrados"
  Exit Sub
End Sub
Sub GeneraArchivoBajaTit()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre As String, ap As String, am As String, elnombre As String
Dim Empleadora As String
Dim RetVal, reg93 As Integer, reg91 As Integer, reg92 As Integer, nemp As Integer, reg922 As Integer

Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
reg92 = 0
reg922 = 0
nemp = 0
If chkUni_Cancel.value = 1 Or chkali_Cancel.value = 1 Or chkreg_Cancel.value = 1 Or chkpre_Cancel.value = 1 Or chkdie_Cancel.value = 1 Then
   Open dir_prueba & "PC" & Format(CDate(mskFechaProc.Text) - 1, "yymmdd") & "_SC.vlt" For Output As #1
  'Header A
   clinea = "11CANCELA04" & Format(CDate(mskFechaProc.Text) - 1, "YYYYMMDD") & rellena("0", 37, "")
   Print #1, clinea
Else
   Exit Sub
End If

If chkUni_Cancel.value = 1 And spdBajaTit.MaxRows > 0 Then
'Header B Uniformes
   reg93 = 0
   clinea = "1301" & rellena("0", 52, "") 'porke es Uniformes
   Print #1, clinea
   With spdBajaTit
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoBajaTitulares',1"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 44, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 44, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Uniformes Titulares
               clinea = "2102"
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               elnombre = Trim(rsBD!Nombre)
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I") & rellena(rsBD!Tarjeta, 16, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C UNIFORMES
            clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B UNIFORMES
   clinea = "93" & Format(reg93, "000000") & rellena("0", 48, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If
'++++++Alimentacion
'Header B Alimentacion
If chkali_Cancel.value = 1 And spdBajaTit2.MaxRows > 0 Then
   reg93 = 0
   clinea = "1301" & rellena("0", 52, "") 'porke es Uniformes
   Print #1, clinea
   With spdBajaTit2
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoBajaTitulares',2"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 44, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 44, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Uniformes Titulares
               clinea = "2102"
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               elnombre = Trim(rsBD!Nombre)
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I") & rellena(rsBD!Tarjeta, 16, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C UNIFORMES
            clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B UNIFORMES
   clinea = "93" & Format(reg93, "000000") & rellena("0", 48, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Regalos
'Header B Regalos
If chkreg_Cancel.value = 1 And spdBajaTit3.MaxRows > 0 Then
   reg93 = 0
   clinea = "1301" & rellena("0", 52, "") 'porke es Uniformes
   Print #1, clinea
   With spdBajaTit3
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoBajaTitulares',3"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Uniformes
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Uniformes
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 44, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 44, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Uniformes Titulares
               clinea = "2102"
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               elnombre = Trim(rsBD!Nombre)
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I") & rellena(rsBD!Tarjeta, 16, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C UNIFORMES
            clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B UNIFORMES
   clinea = "93" & Format(reg93, "000000") & rellena("0", 48, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Premium
'Header B Premium
If chkpre_Cancel.value = 1 And spdBajaTit4.MaxRows > 0 Then
   reg93 = 0
   clinea = "1301" & rellena("0", 52, "") 'porke es Premium
   Print #1, clinea
   With spdBajaTit4
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoBajaTitulares',10"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Premium
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Premium
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 44, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 44, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Premium Titulares
               clinea = "2102"
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               elnombre = Trim(rsBD!Nombre)
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I") & rellena(rsBD!Tarjeta, 16, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C Premium
            clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B Premium
   clinea = "93" & Format(reg93, "000000") & rellena("0", 48, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'++++++Diesel
'Header B Diesel
If chkdie_Cancel.value = 1 And spdBajaTit5.MaxRows > 0 Then
   reg93 = 0
   clinea = "1301" & rellena("0", 52, "") 'porke es Diesel
   Print #1, clinea
   With spdBajaTit5
       For i = 1 To .MaxRows
           .Col = 1
           .Row = i
           cliente = Val(.Text)
           nemp = nemp + 1
           Plaza = 0
           sqls = "sp_Vistas_AltasSBI '" & Format(CDate(mskFechaProc.Text) - 1, "mm/dd/yyyy") & "'," & cliente & ",'ArchivoBajaTitulares',11"
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           reg91 = 0
           Do While Not rsBD.EOF
                If rsBD!Plaza <> Plaza Then
                   If Plaza <> 0 Then
                        'Footer Diesel
                        clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
                        Print #1, clinea
                        reg91 = 0
                   End If
                   'Header C Diesel
                   If cliente < 4 And rsBD!Plaza = 1 Then
                        clinea = "12" & rellena(CStr(cliente), 10, " ", "I") & rellena("0", 44, "")
                   Else
                        clinea = "12" & Format(cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 44, "")
                   End If
                   Print #1, clinea
                   Plaza = rsBD!Plaza
                   reg922 = reg922 + 1
                End If
           'Detail Diesel Titulares
               clinea = "2102"
               clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
               elnombre = Trim(rsBD!Nombre)
               elnombre = Mid(Trim(elnombre), 1, 26)
               If Len(elnombre) = 0 Then
                  elnombre = ""
               End If
               clinea = clinea & rellena(elnombre, 26, " ", "I") & rellena(rsBD!Tarjeta, 16, "")
               Print #1, clinea
               rsBD.MoveNext
               reg93 = reg93 + 1
               reg91 = reg91 + 1
               reg92 = reg92 + 1
           Loop
           'Trailer C Diesel
            clinea = "91" & Format(reg91, "000000") & rellena("0", 48, "")
            Print #1, clinea
            
       Next
   End With
   'Trailer B Diesel
   clinea = "93" & Format(reg93, "000000") & rellena("0", 48, "")
   Print #1, clinea
    rsBD.Close
    Set rsBD = Nothing
End If

'Trailer A
clinea = "92" & Format(reg922, "000000") & Format(reg92, "000000") & rellena("0", 42, "")
Print #1, clinea

Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Baja de titulares"
   Exit Sub
End Sub

Sub CargaDatosAjustes()
On Error GoTo ERRO:
Dim i As Integer
Dim total As Double
Dim TotEmp As Long
'***Uniformes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesC',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    total = 0
    totalemp = 0
    i = 0
    With spdAjustes
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!cliente)
        .Col = 2
        .Text = Val(rsBD!clientedisp)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = rsBD!TotEmp
         totalemp = totalemp + rsBD!TotEmp
        .Col = 5
        .Text = CDbl(rsBD!total)
        total = total + rsBD!total
        rsBD.MoveNext
    Loop
    End With
    txtAjustes = total
    txtTotEmpAju = totalemp
    rsBD.Close
    Set rsBD = Nothing
    
'***Alimentos
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesC',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    total = 0
    totalemp = 0
    i = 0
    With spdAjustes2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!cliente)
        .Col = 2
        .Text = Val(rsBD!clientedisp)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = rsBD!TotEmp
         totalemp = totalemp + rsBD!TotEmp
        .Col = 5
        .Text = CDbl(rsBD!total)
        total = total + rsBD!total
        rsBD.MoveNext
    Loop
    End With
    txtAjustes2 = total
    txtTotEmpAju2 = totalemp
    rsBD.Close
    Set rsBD = Nothing
    
'***Regalo
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesC',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    total = 0
    totalemp = 0
    i = 0
    With spdAjustes3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!cliente)
        .Col = 2
        .Text = Val(rsBD!clientedisp)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = rsBD!TotEmp
         totalemp = totalemp + rsBD!TotEmp
        .Col = 5
        .Text = CDbl(rsBD!total)
        total = total + rsBD!total
        rsBD.MoveNext
    Loop
    End With
    txtAjustes3 = total
    txtTotEmpAju3 = totalemp
    rsBD.Close
    Set rsBD = Nothing

'***Premium
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesC',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    total = 0
    totalemp = 0
    i = 0
    With spdAjustes4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!cliente)
        .Col = 2
        .Text = Val(rsBD!clientedisp)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = rsBD!TotEmp
         totalemp = totalemp + rsBD!TotEmp
        .Col = 5
        .Text = CDbl(rsBD!total)
        total = total + rsBD!total
        rsBD.MoveNext
    Loop
    End With
    txtAjustes4 = total
    txtTotEmpAju4 = totalemp
    rsBD.Close
    Set rsBD = Nothing

'***Regalo
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesC',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    total = 0
    totalemp = 0
    i = 0
    With spdAjustes5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!cliente)
        .Col = 2
        .Text = Val(rsBD!clientedisp)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = rsBD!TotEmp
         totalemp = totalemp + rsBD!TotEmp
        .Col = 5
        .Text = CDbl(rsBD!total)
        total = total + rsBD!total
        rsBD.MoveNext
    Loop
    End With
    txtAjustes5 = total
    txtTotEmpAju5 = totalemp
    rsBD.Close
    Set rsBD = Nothing

Exit Sub
ERRO:
  MsgBox "Errores generados", vbCritical, "Carga de Ajustes"
  Exit Sub
End Sub

Sub CargaDatosDispersion()
On Error GoTo ERRO:
Dim i As Integer
Dim total As Double
Dim TotEmp As Long
'***Uniformes
    total = 0
    totalemp = 0
    '+++Por pedido
    With spdDisp
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    i = 0
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Cargar',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "P"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
    '++Por ajustes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesA',1"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Folio)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "A"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    End With
    txtTotImp = total
    txtTotEmp = totalemp
    rsBD.Close
    Set rsBD = Nothing
    
'***Alimentos
    total = 0
    totalemp = 0
    '+++Por pedido
    With spdDisp2
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    i = 0
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Cargar',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "P"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
    '++Por ajustes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesA',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Folio)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "A"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    End With
    txtTotImp2 = total
    txtTotEmp2 = totalemp
    rsBD.Close
    Set rsBD = Nothing
    
'***Regalo
    total = 0
    totalemp = 0
    '+++Por pedido
    With spdDisp3
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    i = 0
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Cargar',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "P"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
    '++Por ajustes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesA',3"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Folio)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "A"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    End With
    txtTotImp3 = total
    txtTotEmp3 = totalemp
    rsBD.Close
    Set rsBD = Nothing

    
'***Premium
    total = 0
    totalemp = 0
    '+++Por pedido
    With spdDisp4
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    i = 0
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Cargar',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "P"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
    '++Por ajustes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesA',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Folio)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "A"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    End With
    txtTotImp4 = total
    txtTotEmp4 = totalemp
    rsBD.Close
    Set rsBD = Nothing

'***Diesel
    total = 0
    totalemp = 0
    '+++Por pedido
    With spdDisp5
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    i = 0
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Cargar',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "P"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
    '++Por ajustes
    sqls = " SP_Vistas_PagoGas '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "',0,'AjustesA',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Folio)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = Val(rsBD!clientedisp)
        .Col = 4
        .Text = rsBD!Nombre
        .Col = 5
        .Text = rsBD!Empleados
        totalemp = totalemp + rsBD!Empleados
        .Col = 6
        .Text = CDbl(rsBD!importe)
        .Col = 8
        .Text = "A"
        total = total + rsBD!importe
        rsBD.MoveNext
    Loop
    End With
    txtTotImp5 = total
    txtTotEmp5 = totalemp
    rsBD.Close
    Set rsBD = Nothing

Exit Sub
ERRO:
  MsgBox "Errores generados", vbCritical, "Carga de Dispersiones"
  Exit Sub
End Sub

Sub GeneraArchivoAjustes()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre, j As Long, importe As Double
Dim Empleadora As String
Dim RetVal, reg97 As Long, total97 As Double, reg96 As Long, total96 As Double
Dim ImporteCliente As Double, EmpleadosCliente As Long
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Long
Dim clientedisp As Long
On Error GoTo err_gral

Call busca_ventana

'Header A
If (chkUni_Ajustes.value = 1 And spdAjustes.MaxRows > 0) Or _
   (chkali_Ajustes.value = 1 And spdAjustes2.MaxRows > 0) Or _
   (chkreg_Ajustes.value = 1 And spdAjustes3.MaxRows > 0) Or _
   (chkpre_Ajustes.value = 1 And spdAjustes4.MaxRows > 0) Or _
   (chkdie_Ajustes.value = 1 And spdAjustes5.MaxRows > 0) Then
   Open dir_prueba & "AJ" & Format(CDate(mskFechaProc.Text), "yymmdd") & Ventana & "_SC.vlt" For Output As #1
   clinea = "15AJUSTES04" & Format(CDate(mskFechaProc.Text), "YYYYMMDD")
   Print #1, clinea
Else
   Exit Sub
End If

total96 = 0
reg96 = 0
'***UNIFORMES
'HEADER B
If (chkUni_Ajustes.value = 1 And spdAjustes.MaxRows > 0) Then
clinea = "1701" & rellena("0", 15, "")
Print #1, clinea

With spdAjustes
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosAjustes',1"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'Ajustes_Archivos',1," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***ALIMENTACION
'HEADER B
If (chkali_Ajustes.value = 1 And spdAjustes2.MaxRows > 0) Then
clinea = "1702" & rellena("0", 15, "")
Print #1, clinea

With spdAjustes2
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosAjustes',2"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'Ajustes_Archivos',2," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       importe = importe + rsBD2!importe
       j = j + 1
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***REGALO
'HEADER B
If (chkreg_Ajustes.value = 1 And spdAjustes3.MaxRows > 0) Then
clinea = "1704" & rellena("0", 15, "")
Print #1, clinea

With spdAjustes3
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosAjustes',3"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'Ajustes_Archivos',3," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***Premium
'HEADER B
If (chkpre_Ajustes.value = 1 And spdAjustes4.MaxRows > 0) Then
clinea = "1704" & rellena("0", 15, "")
Print #1, clinea

With spdAjustes4
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosAjustes',10"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'Ajustes_Archivos',10," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***Diesel
'HEADER B
If (chkdie_Ajustes.value = 1 And spdAjustes5.MaxRows > 0) Then
clinea = "1704" & rellena("0", 15, "")
Print #1, clinea

With spdAjustes5
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosAjustes',11"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'Ajustes_Archivos',11," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'Trailer A
clinea = "96" & Format(reg96, "000000") & Format(total96 * 100, "00000000000")
Print #1, clinea

Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Archivo de Ajustes"
   Exit Sub
End Sub

Sub GeneraArchivoDispersiones()
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre, j As Long, importe As Double
Dim Empleadora As String
Dim RetVal, reg97 As Long, total97 As Double, reg96 As Long, total96 As Double
Dim ImporteCliente As Double, EmpleadosCliente As Long
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Long
Dim clientedisp As Long
On Error GoTo err_gral

Call busca_ventana

'Header A
If (chkUni_Disp.value = 1 And spdDisp.MaxRows > 0) Or _
   (chkali_Disp.value = 1 And spdDisp2.MaxRows > 0) Or _
   (chkreg_Disp.value = 1 And spdDisp3.MaxRows > 0) Or _
   (chkpre_Disp.value = 1 And spdDisp4.MaxRows > 0) Or _
   (chkdie_Disp.value = 1 And spdDisp5.MaxRows > 0) Then
   Open dir_prueba & "DS" & Format(CDate(mskFechaProc.Text), "yymmdd") & Ventana & "_SC.vlt" For Output As #1
   clinea = "15SALDOS 04" & Format(CDate(mskFechaProc.Text), "YYYYMMDD")
   Print #1, clinea
Else
   Exit Sub
End If

total96 = 0
reg96 = 0
'***UNIFORMES
'HEADER B
If (chkUni_Disp.value = 1 And spdDisp.MaxRows > 0) Then
clinea = "1701" & rellena("0", 15, "")
Print #1, clinea

With spdDisp
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosDispersion',1"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'ArchivosDispersion2',1," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***ALIMENTACION
'HEADER B
If (chkali_Disp.value = 1 And spdDisp2.MaxRows > 0) Then
clinea = "1702" & rellena("0", 15, "")
Print #1, clinea

With spdDisp2
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosDispersion',2"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'ArchivosDispersion2',2," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       importe = importe + rsBD2!importe
       j = j + 1
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***REGALO
'HEADER B
If (chkreg_Disp.value = 1 And spdDisp3.MaxRows > 0) Then
clinea = "1703" & rellena("0", 15, "")
Print #1, clinea

With spdDisp3
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosDispersion',3"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'ArchivosDispersion2',3," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***Premium
'HEADER B
If (chkpre_Disp.value = 1 And spdDisp4.MaxRows > 0) Then
clinea = "1710" & rellena("0", 15, "")
Print #1, clinea

With spdDisp4
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosDispersion',10"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'ArchivosDispersion2',10," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'***DIESEL
'HEADER B
If (chkdie_Disp.value = 1 And spdDisp5.MaxRows > 0) Then
clinea = "1711" & rellena("0", 15, "")
Print #1, clinea

With spdDisp5
 total97 = 0
 reg97 = 0
 sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'ArchivosDispersion',11"
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 Do While Not rsBD.EOF
    'HEADER C
    clientedisp = rsBD!clientedisp
    If clientedisp < 4 And rsBD!Plaza = 1 Then
         clinea = "16" & rellena(CStr(clientedisp), 10, " ", "I") & rellena("0", 7, "")
    Else
         clinea = "16" & Format(clientedisp, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
    End If
    Print #1, clinea
    
    sqls = "SP_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'," & clientedisp & ",'ArchivosDispersion2',11," & rsBD!Plaza
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    'Detail
    j = 0
    importe = 0
    Do While Not rsBD2.EOF
       j = j + 1
       importe = importe + rsBD2!importe
       clinea = "23"
       clinea = clinea & Format(rsBD2!Cuenta, "00000000")
       clinea = clinea & Format(rsBD2!importe * 100, "000000000")
       Print #1, clinea
       reg97 = reg97 + 1
       reg96 = reg96 + 1
       total97 = total97 + rsBD2!importe
       total96 = total96 + rsBD2!importe
       rsBD2.MoveNext
    Loop
    'Trailer C
    clinea = "95" & Format(j, "000000")
    clinea = clinea & Format(importe * 100, "00000000000")
    Print #1, clinea
    rsBD.MoveNext
 Loop
 'Trailer B
 clinea = "97" & Format(reg97, "000000") & Format(total97 * 100, "00000000000")
 Print #1, clinea
End With
rsBD.Close
Set rsBD = Nothing
End If

'Trailer A
clinea = "96" & Format(reg96, "000000") & Format(total96 * 100, "00000000000")
Print #1, clinea

Close #1
Exit Sub
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Archivo de Dispersiones"
   Exit Sub
End Sub

Sub GrabaTarjetas()
Dim i As Long
Dim idarchivo As String
'***Uniformes
idarchivo = "AT" & Format(CDate(mskFechaProc.Text), "mmddyyyy")
sqls = "delete sbiAltaTarjetas where fechaproc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=1"
cnxBD.Execute sqls, intRegistros
  
With spdAltaTit
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =1"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =1"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

With spdAltaStockTit
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =1"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =1"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Alimentacion
sqls = "delete sbiAltaTarjetas where fechaproc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=2"
cnxBD.Execute sqls, intRegistros
  
With spdAltaTit2
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =2"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =2"
        cnxBD.Execute sqls, intRegistros
   Next i
End With
With spdAltaStockTit2
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =2"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =2"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Regalo
sqls = "delete sbiAltaTarjetas where fechaproc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=3"
cnxBD.Execute sqls, intRegistros
  
With spdAltaTit3
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =3"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =3"
        cnxBD.Execute sqls, intRegistros
   Next i
End With
With spdAltaStockTit3
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =3"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =3"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Premium
sqls = "delete sbiAltaTarjetas where fechaproc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=10"
cnxBD.Execute sqls, intRegistros
  
With spdAltaTit4
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =10"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =10"
        cnxBD.Execute sqls, intRegistros
   Next i
End With
With spdAltaStockTit4
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =10"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =10"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Diesel
sqls = "delete sbiAltaTarjetas where fechaproc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=11"
cnxBD.Execute sqls, intRegistros
  
With spdAltaTit5
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =11"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =11"
        cnxBD.Execute sqls, intRegistros
   Next i
End With
With spdAltaStockTit5
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'T'" & _
                ", @Producto =11"
        cnxBD.Execute sqls, intRegistros
        sqls = "  exec  sp_SBIAltaTarjetas " & _
                "  @Id_Archivo= '" & idarchivo & "'" & _
                ", @Cliente = " & cliente & _
                ", @FechaProc = '" & Format(CDate(mskFechaProc.Text), "mm/dd/yyyy") & "'" & _
                ", @TipoTar = 'A'" & _
                ", @Producto =11"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

End Sub

Sub GrabaAjustes()
Dim i As Long

'***Uniformes
sqls = "delete sbiajustes where fechaproc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=1"
cnxBD.Execute sqls, intRegistros

With spdAjustes
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        .Col = 2
        clientedisp = Val(.Text)
        sqls = "  exec  sp_SBIAjustesBE " & _
                       "  @Id_Archivo= 'AJ " & Format(CDate(mskFechaProc.Text), "yymmdd") & "'" & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=1"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Alimentacion
sqls = "delete sbiajustes where fechaproc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=2"
cnxBD.Execute sqls, intRegistros

With spdAjustes2
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        .Col = 2
        clientedisp = Val(.Text)
        sqls = "  exec  sp_SBIAjustesBE " & _
                       "  @Id_Archivo= 'AJ " & Format(CDate(mskFechaProc.Text), "yymmdd") & "'" & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=2"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Regalo
sqls = "delete sbiajustes where fechaproc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=3"
cnxBD.Execute sqls, intRegistros

With spdAjustes3
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        .Col = 2
        clientedisp = Val(.Text)
        sqls = "  exec  sp_SBIAjustesBE " & _
                       "  @Id_Archivo= 'AJ " & Format(CDate(mskFechaProc.Text), "yymmdd") & "'" & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=3"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Premium
sqls = "delete sbiajustes where fechaproc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=10"
cnxBD.Execute sqls, intRegistros

With spdAjustes4
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        .Col = 2
        clientedisp = Val(.Text)
        sqls = "  exec  sp_SBIAjustesBE " & _
                       "  @Id_Archivo= 'AJ " & Format(CDate(mskFechaProc.Text), "yymmdd") & "'" & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=10"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

'***Diesel
sqls = "delete sbiajustes where fechaproc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=11"
cnxBD.Execute sqls, intRegistros

With spdAjustes5
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        cliente = Val(.Text)
        .Col = 2
        clientedisp = Val(.Text)
        sqls = "  exec  sp_SBIAjustesBE " & _
                       "  @Id_Archivo= 'AJ " & Format(CDate(mskFechaProc.Text), "yymmdd") & "'" & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=11"
        cnxBD.Execute sqls, intRegistros
   Next i
End With

End Sub

Sub GrabaDispersiones()
Dim i As Long
Dim idarchivo As String
idarchivo = "SD" & Format(mskFechaProc.Text, "mmddyyyy")
'**Uniformes
sqls = "sp_Vistas_AltasSBI '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'BorraSBIDisp',1"
cnxBD.Execute sqls, intRegistros
With spdDisp
    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        Status = .value
        If Status <> 0 Then
            .Col = 1
            Pedido = Val(.Text)
            .Col = 2
            cliente = Val(.Text)
            .Col = 3
            clientedisp = Val(.Text)
            .Col = 8
            tipo = .Text
            
            sqls = "  exec  sp_SBIDispersionesBE " & _
                       "  @Id_Archivo= '" & idarchivo & "'" & _
                       ", @Pedido = " & Pedido & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=1"
            cnxBD.Execute sqls, intRegistros
        End If
   Next i
End With

'***Alimentacion
sqls = "sp_Vistas_AltasSBI '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'BorraSBIDisp',2"
cnxBD.Execute sqls, intRegistros
With spdDisp2
    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        Status = .value
        If Status <> 0 Then
            .Col = 1
            Pedido = Val(.Text)
            .Col = 2
            cliente = Val(.Text)
            .Col = 3
            clientedisp = Val(.Text)
            .Col = 8
            tipo = .Text
            
            sqls = "  exec  sp_SBIDispersionesBE " & _
                       "  @Id_Archivo= '" & idarchivo & "'" & _
                       ", @Pedido = " & Pedido & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=2"
            cnxBD.Execute sqls, intRegistros
        End If
   Next i
End With

'***Regalo
sqls = "sp_Vistas_AltasSBI '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'BorraSBIDisp',3"
cnxBD.Execute sqls, intRegistros
With spdDisp3
    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        Status = .value
        If Status <> 0 Then
            .Col = 1
            Pedido = Val(.Text)
            .Col = 2
            cliente = Val(.Text)
            .Col = 3
            clientedisp = Val(.Text)
            .Col = 8
            tipo = .Text
            
            sqls = "  exec  sp_SBIDispersionesBE " & _
                       "  @Id_Archivo= '" & idarchivo & "'" & _
                       ", @Pedido = " & Pedido & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=3"
            cnxBD.Execute sqls, intRegistros
        End If
   Next i
End With

'***Premium
sqls = "sp_Vistas_AltasSBI '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'BorraSBIDisp',10"
cnxBD.Execute sqls, intRegistros
With spdDisp4
    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        Status = .value
        If Status <> 0 Then
            .Col = 1
            Pedido = Val(.Text)
            .Col = 2
            cliente = Val(.Text)
            .Col = 3
            clientedisp = Val(.Text)
            .Col = 8
            tipo = .Text
            
            sqls = "  exec  sp_SBIDispersionesBE " & _
                       "  @Id_Archivo= '" & idarchivo & "'" & _
                       ", @Pedido = " & Pedido & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=10"
            cnxBD.Execute sqls, intRegistros
        End If
   Next i
End With

'***Diesel
sqls = "sp_Vistas_AltasSBI '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'BorraSBIDisp',11"
cnxBD.Execute sqls, intRegistros
With spdDisp5
    For i = 1 To .MaxRows
        .Row = i
        .Col = 7
        Status = .value
        If Status <> 0 Then
            .Col = 1
            Pedido = Val(.Text)
            .Col = 2
            cliente = Val(.Text)
            .Col = 3
            clientedisp = Val(.Text)
            .Col = 8
            tipo = .Text
            
            sqls = "  exec  sp_SBIDispersionesBE " & _
                       "  @Id_Archivo= '" & idarchivo & "'" & _
                       ", @Pedido = " & Pedido & _
                       ", @Cliente = " & cliente & _
                       ", @ClienteDisp = " & clientedisp & _
                       ", @FechaProc = '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "'" & _
                       ", @Producto=11"
            cnxBD.Execute sqls, intRegistros
        End If
   Next i
End With

End Sub
Sub GrabaCombustibles()
    sqls = "sp_Vistas_PagoGas '" & Format(mskFechaProc.Text, "mm/dd/yyyy") & "',NULL,'Combustible'"
    cnxBD.Execute sqls, intRegistros
End Sub
Sub busca_ventana()

If Format(Time, "HH:MM:SS") >= "08:00:00" And Format(Time, "HH:MM:SS") <= "11:59:59" Then
   Ventana = "_01"
ElseIf Format(Time, "HH:MM:SS") >= "12:00:00" And Format(Time, "HH:MM:SS") <= "13:59:59" Then
   Ventana = "_02"
ElseIf Format(Time, "HH:MM:SS") >= "14:00:00" And Format(Time, "HH:MM:SS") <= "15:00:00" Then
   Ventana = "_03"
Else
   Ventana = ""
End If
End Sub

