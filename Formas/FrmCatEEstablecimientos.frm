VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCatEEstablecimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Establecimientos"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   7395
   Begin MSComctlLib.Toolbar BarMenu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nueva Captura"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Modificación"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar Captura"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Limpiar"
            Object.ToolTipText     =   "Limpiar Captura"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reporte"
            Object.ToolTipText     =   "Reporte"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "A"
                  Text            =   "Activos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "I"
                  Text            =   "Inactivo"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "T"
                  Text            =   "Todos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   8490
      Left            =   0
      TabIndex        =   46
      Top             =   615
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   14975
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Búsqueda"
      TabPicture(0)   =   "FrmCatEEstablecimientos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImgMenu"
      Tab(0).Control(1)=   "FraBusqueda"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Captura"
      TabPicture(1)   =   "FrmCatEEstablecimientos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraCaptura"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contratos"
      TabPicture(2)   =   "FrmCatEEstablecimientos.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraContrato"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Datos Banco"
      TabPicture(3)   =   "FrmCatEEstablecimientos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FraBanco"
      Tab(3).ControlCount=   1
      Begin VB.Frame FraBanco 
         Height          =   6210
         Left            =   -74910
         TabIndex        =   109
         Top             =   345
         Width           =   7200
         Begin VB.TextBox TxtCuenta 
            Height          =   315
            Left            =   1110
            MaxLength       =   18
            TabIndex        =   45
            Top             =   1950
            Width           =   1890
         End
         Begin VB.ComboBox CboTipoPago 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1455
            Width           =   3660
         End
         Begin VB.TextBox TxtSucursal 
            Height          =   315
            Left            =   1110
            MaxLength       =   10
            TabIndex        =   43
            Top             =   960
            Width           =   1050
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   480
            Width           =   5640
         End
         Begin VB.Label Label45 
            Caption         =   "Cuenta"
            Height          =   165
            Left            =   315
            TabIndex        =   113
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label44 
            Caption         =   "Tipo Pago"
            Height          =   195
            Left            =   315
            TabIndex        =   112
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label Label43 
            Caption         =   "Sucursal"
            Height          =   165
            Left            =   315
            TabIndex        =   111
            Top             =   1065
            Width           =   810
         End
         Begin VB.Label Label42 
            Caption         =   "Banco"
            Height          =   240
            Left            =   315
            TabIndex        =   110
            Top             =   585
            Width           =   600
         End
      End
      Begin VB.Frame fraContrato 
         Height          =   6210
         Left            =   90
         TabIndex        =   83
         Top             =   345
         Width           =   7200
         Begin VB.OptionButton OptContInact 
            Caption         =   "Inactivo"
            Height          =   240
            Left            =   2310
            TabIndex        =   40
            Top             =   1425
            Width           =   1065
         End
         Begin VB.OptionButton OptContAct 
            Caption         =   "Activo"
            Height          =   240
            Left            =   1155
            TabIndex        =   39
            Top             =   1425
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid GrdContrato 
            Height          =   4290
            Left            =   105
            TabIndex        =   86
            Top             =   1830
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   7567
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            Height          =   315
            Left            =   3600
            TabIndex        =   41
            Top             =   1365
            Width           =   1095
         End
         Begin VB.ComboBox cboProd 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   465
            Width           =   2385
         End
         Begin VB.TextBox TxtComision 
            Height          =   315
            Left            =   4665
            TabIndex        =   36
            Top             =   465
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker DTPFechaContIni 
            Height          =   330
            Left            =   1155
            TabIndex        =   37
            Top             =   915
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            Format          =   59179009
            CurrentDate     =   40273
         End
         Begin MSComCtl2.DTPicker DTPFechaContFin 
            Height          =   330
            Left            =   4665
            TabIndex        =   38
            Top             =   915
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            Format          =   59179009
            CurrentDate     =   40275
         End
         Begin VB.Label Label39 
            Caption         =   "Status"
            Height          =   180
            Left            =   210
            TabIndex        =   105
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "Fecha Fin"
            Height          =   210
            Left            =   3885
            TabIndex        =   104
            Top             =   1035
            Width           =   885
         End
         Begin VB.Label Label38 
            Caption         =   "Fecha Inicio"
            Height          =   210
            Left            =   210
            TabIndex        =   103
            Top             =   1035
            Width           =   900
         End
         Begin VB.Label Label25 
            Caption         =   "Producto"
            Height          =   180
            Left            =   210
            TabIndex        =   85
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label13 
            Caption         =   "Comisión"
            Height          =   180
            Left            =   3885
            TabIndex        =   84
            Top             =   570
            Width           =   675
         End
      End
      Begin VB.Frame FraBusqueda 
         Height          =   6225
         Left            =   -74925
         TabIndex        =   50
         Top             =   345
         Width           =   7230
         Begin VB.ComboBox CboProductoB 
            Height          =   315
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1620
            Width           =   4725
         End
         Begin VB.OptionButton OptInactivoB 
            Caption         =   "Inactivo"
            Height          =   195
            Left            =   1890
            TabIndex        =   5
            Top             =   2100
            Width           =   1125
         End
         Begin VB.OptionButton OptActivoB 
            Caption         =   "Activo"
            Height          =   195
            Left            =   930
            TabIndex        =   4
            Top             =   2100
            Width           =   765
         End
         Begin VB.ComboBox CboSucursalB 
            Height          =   315
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   4725
         End
         Begin VB.TextBox TxtClaveB 
            Height          =   315
            Left            =   930
            TabIndex        =   0
            Top             =   345
            Width           =   1200
         End
         Begin VB.TextBox TxtNombreB 
            Height          =   315
            Left            =   930
            MaxLength       =   45
            TabIndex        =   1
            Top             =   765
            Width           =   4725
         End
         Begin VB.CommandButton CmdBuscar 
            Height          =   420
            Left            =   6075
            Picture         =   "FrmCatEEstablecimientos.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Generar Búsqueda"
            Top             =   1965
            Width           =   480
         End
         Begin VB.CommandButton CmdLimpiar 
            Height          =   420
            Left            =   6600
            Picture         =   "FrmCatEEstablecimientos.frx":088A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Búsqueda"
            Top             =   1965
            Width           =   480
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3690
            Left            =   75
            TabIndex        =   51
            Top             =   2430
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   6509
            _Version        =   393216
            FixedCols       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label40 
            Caption         =   "Producto"
            Height          =   195
            Left            =   225
            TabIndex        =   107
            Top             =   1725
            Width           =   675
         End
         Begin VB.Label Label18 
            Caption         =   "Status"
            Height          =   210
            Left            =   225
            TabIndex        =   67
            Top             =   2100
            Width           =   570
         End
         Begin VB.Label Label17 
            Caption         =   "Sucursal"
            Height          =   210
            Left            =   225
            TabIndex        =   66
            Top             =   1290
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre"
            Height          =   210
            Left            =   225
            TabIndex        =   65
            Top             =   870
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
            Height          =   210
            Left            =   225
            TabIndex        =   64
            Top             =   450
            Width           =   615
         End
      End
      Begin VB.Frame FraCaptura 
         Height          =   7950
         Left            =   -74925
         TabIndex        =   47
         Top             =   345
         Width           =   7215
         Begin VB.TextBox TxtDC 
            Height          =   300
            Left            =   4545
            TabIndex        =   34
            Top             =   5835
            Width           =   1050
         End
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   315
            Left            =   4560
            TabIndex        =   29
            Top             =   5055
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            Format          =   59179009
            CurrentDate     =   38511
         End
         Begin VB.TextBox TxtComAnt 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5835
            TabIndex        =   106
            Top             =   5430
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Frame FraDes 
            Height          =   1575
            Left            =   105
            TabIndex        =   87
            Top             =   6270
            Width           =   7020
            Begin VB.TextBox TxtNomComer 
               Height          =   315
               Left            =   1290
               TabIndex        =   91
               Top             =   270
               Width           =   5580
            End
            Begin VB.TextBox TxtTpv 
               Height          =   315
               Left            =   1290
               TabIndex        =   90
               Top             =   690
               Width           =   2400
            End
            Begin VB.TextBox TxtAfiliado 
               Height          =   315
               Left            =   4290
               TabIndex        =   89
               Top             =   690
               Width           =   2400
            End
            Begin VB.CommandButton CmdGrabaDes 
               Caption         =   "Ok"
               Height          =   315
               Left            =   2970
               TabIndex        =   88
               Top             =   1110
               Width           =   1455
            End
            Begin VB.Label Label26 
               Caption         =   "Nom. Comercial"
               Height          =   240
               Left            =   105
               TabIndex        =   94
               Top             =   375
               Width           =   1110
            End
            Begin VB.Label Label27 
               Caption         =   "T.P.V."
               Height          =   240
               Left            =   105
               TabIndex        =   93
               Top             =   810
               Width           =   1065
            End
            Begin VB.Label Label28 
               Caption         =   "Afiliado"
               Height          =   240
               Left            =   3705
               TabIndex        =   92
               Top             =   780
               Width           =   600
            End
         End
         Begin VB.TextBox TxtCurp 
            Height          =   315
            Left            =   4440
            TabIndex        =   26
            Top             =   4305
            Width           =   2445
         End
         Begin VB.TextBox TxtEmail 
            Height          =   315
            Left            =   1245
            TabIndex        =   14
            Top             =   1695
            Width           =   5625
         End
         Begin VB.TextBox TxtComercial 
            Height          =   315
            Left            =   1245
            TabIndex        =   13
            Top             =   1335
            Width           =   5625
         End
         Begin VB.TextBox TxtNoInt 
            Height          =   315
            Left            =   4530
            TabIndex        =   21
            Top             =   3165
            Width           =   2340
         End
         Begin VB.TextBox TxtNoExt 
            Height          =   315
            Left            =   1245
            TabIndex        =   20
            Top             =   3165
            Width           =   2325
         End
         Begin VB.TextBox TxtAMaterno 
            Height          =   315
            Left            =   4575
            MaxLength       =   30
            TabIndex        =   12
            Top             =   975
            Width           =   2295
         End
         Begin VB.TextBox TxtAPaterno 
            Height          =   315
            Left            =   1245
            MaxLength       =   30
            TabIndex        =   11
            Top             =   975
            Width           =   2295
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   3300
            TabIndex        =   74
            Top             =   120
            Width           =   2880
            Begin VB.OptionButton OptMoral 
               Caption         =   "Moral"
               Height          =   285
               Left            =   2100
               TabIndex        =   9
               Top             =   180
               Width           =   735
            End
            Begin VB.OptionButton OptFisica 
               Caption         =   "Física"
               Height          =   285
               Left            =   1215
               TabIndex        =   8
               Top             =   180
               Width           =   735
            End
            Begin VB.Label Label30 
               Caption         =   "Tipo Persona"
               Height          =   285
               Left            =   135
               TabIndex        =   75
               Top             =   210
               Width           =   960
            End
         End
         Begin VB.TextBox TxtNombre 
            Height          =   315
            Left            =   1245
            MaxLength       =   100
            TabIndex        =   10
            Top             =   615
            Width           =   5625
         End
         Begin VB.TextBox TxtColonia 
            Height          =   315
            Left            =   1245
            MaxLength       =   40
            TabIndex        =   22
            Top             =   3555
            Width           =   5625
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2415
            Width           =   2370
         End
         Begin VB.ComboBox CboMunicipio 
            Height          =   315
            Left            =   4515
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2415
            Width           =   2370
         End
         Begin VB.TextBox TxtEncargado 
            Height          =   315
            Left            =   1245
            MaxLength       =   40
            TabIndex        =   27
            Top             =   4680
            Width           =   5625
         End
         Begin VB.TextBox TxtDomicilio 
            Height          =   315
            Left            =   1245
            MaxLength       =   45
            TabIndex        =   19
            Top             =   2790
            Width           =   5625
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2385
            TabIndex        =   73
            Top             =   6570
            Width           =   1875
         End
         Begin VB.CommandButton CmdSucursales 
            Caption         =   "&Sucursales"
            Height          =   330
            Left            =   5685
            TabIndex        =   72
            Top             =   5835
            Width           =   1185
         End
         Begin VB.ComboBox CboProducto 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   5430
            Width           =   2445
         End
         Begin VB.ComboBox CboGiro 
            Height          =   315
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   5055
            Width           =   2475
         End
         Begin VB.ComboBox CboImpuesto 
            Height          =   315
            ItemData        =   "FrmCatEEstablecimientos.frx":0DD4
            Left            =   4545
            List            =   "FrmCatEEstablecimientos.frx":0DDB
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   5430
            Width           =   960
         End
         Begin VB.OptionButton OptInactivo 
            Caption         =   "Inactivo"
            Height          =   195
            Left            =   2295
            TabIndex        =   33
            Top             =   5925
            Width           =   900
         End
         Begin VB.OptionButton OptActivo 
            Caption         =   "Activo"
            Height          =   195
            Left            =   1245
            TabIndex        =   32
            Top             =   5925
            Width           =   810
         End
         Begin VB.TextBox TxtRFC 
            Height          =   315
            Left            =   1245
            TabIndex        =   25
            Top             =   4305
            Width           =   2325
         End
         Begin VB.TextBox TxtTel 
            Height          =   315
            Left            =   4440
            MaxLength       =   20
            TabIndex        =   24
            Top             =   3930
            Width           =   2445
         End
         Begin VB.TextBox TxtCP 
            Height          =   315
            Left            =   1245
            TabIndex        =   23
            Top             =   3930
            Width           =   2325
         End
         Begin VB.ComboBox CboSucursal 
            Height          =   315
            Left            =   2205
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2055
            Width           =   4680
         End
         Begin VB.TextBox TxtNSucursal 
            Height          =   315
            Left            =   1245
            TabIndex        =   15
            Top             =   2055
            Width           =   930
         End
         Begin VB.Frame FraGas 
            Height          =   1500
            Left            =   105
            TabIndex        =   95
            Top             =   6195
            Visible         =   0   'False
            Width           =   7020
            Begin VB.TextBox TxtNGpo 
               Height          =   315
               Left            =   1005
               MaxLength       =   50
               TabIndex        =   99
               Top             =   675
               Width           =   5520
            End
            Begin VB.TextBox TxtGpo 
               Height          =   315
               Left            =   1005
               MaxLength       =   10
               TabIndex        =   98
               Top             =   270
               Width           =   2325
            End
            Begin VB.TextBox TxtServicio 
               Height          =   315
               Left            =   4440
               MaxLength       =   10
               TabIndex        =   97
               Top             =   270
               Width           =   1830
            End
            Begin VB.CommandButton CmdGrabaGas 
               Caption         =   "Ok"
               Height          =   315
               Left            =   2730
               TabIndex        =   96
               Top             =   1095
               Width           =   1455
            End
            Begin VB.Label Label24 
               Caption         =   "Nom. Gpo"
               Height          =   225
               Left            =   135
               TabIndex        =   102
               Top             =   780
               Width           =   900
            End
            Begin VB.Label Label23 
               Caption         =   "Grupo"
               Height          =   210
               Left            =   135
               TabIndex        =   101
               Top             =   375
               Width           =   630
            End
            Begin VB.Label Label21 
               Caption         =   "No. Servicio"
               Height          =   210
               Left            =   3480
               TabIndex        =   100
               Top             =   375
               Width           =   960
            End
         End
         Begin VB.Label Label41 
            Caption         =   "Días Cred."
            Height          =   180
            Left            =   3735
            TabIndex        =   108
            Top             =   5925
            Width           =   780
         End
         Begin VB.Label Label37 
            Caption         =   "C.U.R.P."
            Height          =   225
            Left            =   3720
            TabIndex        =   82
            Top             =   4410
            Width           =   735
         End
         Begin VB.Label Label36 
            Caption         =   "E-Mail"
            Height          =   165
            Left            =   300
            TabIndex        =   81
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label35 
            Caption         =   "Nom. Comer."
            Height          =   165
            Left            =   285
            TabIndex        =   80
            Top             =   1425
            Width           =   930
         End
         Begin VB.Label Label34 
            Caption         =   "No. Interior"
            Height          =   225
            Left            =   3705
            TabIndex        =   79
            Top             =   3270
            Width           =   795
         End
         Begin VB.Label Label33 
            Caption         =   "No. Exterior"
            Height          =   195
            Left            =   300
            TabIndex        =   78
            Top             =   3270
            Width           =   885
         End
         Begin VB.Label Label32 
            Caption         =   "A. Materno"
            Height          =   180
            Left            =   3660
            TabIndex        =   77
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Label31 
            Caption         =   "A. Paterno"
            Height          =   210
            Left            =   300
            TabIndex        =   76
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Label14 
            Caption         =   "Giro"
            Height          =   180
            Left            =   300
            TabIndex        =   71
            Top             =   5190
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "Producto"
            Height          =   180
            Left            =   300
            TabIndex        =   70
            Top             =   5535
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Impuesto"
            Height          =   225
            Left            =   3825
            TabIndex        =   69
            Top             =   5535
            Width           =   690
         End
         Begin VB.Label Label19 
            Caption         =   "Calle"
            Height          =   180
            Left            =   270
            TabIndex        =   68
            Top             =   2910
            Width           =   660
         End
         Begin VB.Label Label16 
            Caption         =   "Status"
            Height          =   165
            Left            =   300
            TabIndex        =   63
            Top             =   5895
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha"
            Height          =   180
            Left            =   3840
            TabIndex        =   62
            Top             =   5190
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Encargado"
            Height          =   210
            Left            =   300
            TabIndex        =   61
            Top             =   4800
            Width           =   930
         End
         Begin VB.Label Label11 
            Caption         =   "R.F.C."
            Height          =   180
            Left            =   300
            TabIndex        =   60
            Top             =   4395
            Width           =   585
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfono"
            Height          =   210
            Left            =   3720
            TabIndex        =   59
            Top             =   4035
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "Cod. Postal"
            Height          =   210
            Left            =   300
            TabIndex        =   58
            Top             =   4035
            Width           =   945
         End
         Begin VB.Label Label8 
            Caption         =   "Municipio"
            Height          =   225
            Left            =   3810
            TabIndex        =   57
            Top             =   2520
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "Estado"
            Height          =   225
            Left            =   300
            TabIndex        =   56
            Top             =   2520
            Width           =   690
         End
         Begin VB.Label Label6 
            Caption         =   "Colonia"
            Height          =   225
            Left            =   300
            TabIndex        =   55
            Top             =   3660
            Width           =   720
         End
         Begin VB.Label Label5 
            Caption         =   "Sucursal"
            Height          =   225
            Left            =   300
            TabIndex        =   54
            Top             =   2160
            Width           =   1125
         End
         Begin VB.Label LblFolio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1245
            TabIndex        =   53
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label Label3 
            Caption         =   "Folio"
            Height          =   195
            Left            =   300
            TabIndex        =   49
            Top             =   330
            Width           =   1380
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre"
            Height          =   210
            Left            =   300
            TabIndex        =   48
            Top             =   720
            Width           =   1515
         End
      End
      Begin MSComctlLib.ImageList ImgMenu 
         Left            =   -74820
         Top             =   1665
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":0DE3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":1ABD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":2797
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":3471
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":414B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCatEEstablecimientos.frx":4E25
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCatEEstablecimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FOpcion As Integer, OpcionDet As Integer, Contrato As Integer

Private Sub BarMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
   Case "Nuevo"
        Nuevo
   Case "Editar"
        Editar
   Case "Guardar"
        Guardar
   Case "Limpiar"
        Inicio
   Case "Reporte"
        Imprimir
   Case "Salir"
        Unload Me
End Select
End Sub
Private Sub BarMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Opcion As Integer, Status As Integer
Select Case ButtonMenu.Key
    Case "A":   Opcion = 1
                Status = 1
                Call Imprimir
    Case "I":   Opcion = 1
                Status = 0
                Call Imprimir
    Case "T":   Opcion = 0
                Status = 2
                Call Imprimir
End Select
'A: Activos
'I: Inactivos
'T: Todos
End Sub

Private Sub CboEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If CboEstado.ListIndex <> -1 Then
      CboMunicipio.SetFocus
   Else
      Call Mensajes(5)
      CboEstado.SetFocus
   End If
End If
End Sub

Private Sub CboEstado_LostFocus()
If CboEstado.ListIndex <> -1 Then
   CboMunicipio.Clear
   Call CargaPoblaciones(CboMunicipio, CboEstado.ItemData(CboEstado.ListIndex))
End If
End Sub

Private Sub CboGiro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If CboGiro.ListIndex <> -1 Then
      DTPFecha.SetFocus
   Else
      Call Mensajes(5)
      CboGiro.SetFocus
   End If
End If
End Sub
Private Sub CboImpuesto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If CboImpuesto.ListIndex <> -1 Then
      OptActivo.SetFocus
   Else
      Call Mensajes(5)
      CboImpuesto.SetFocus
   End If
End If
End Sub

Private Sub CboMunicipio_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If CboMunicipio.ListIndex <> -1 Then
      txtDomicilio.SetFocus
   Else
      Call Mensajes(5)
      CboMunicipio.SetFocus
   End If
End If
End Sub

Private Sub cboProd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If cboProd.ListIndex <> -1 Then
        txtComision.SetFocus
    Else
        Call Mensajes(5)
        cboProd.SetFocus
    End If
End If
End Sub

Private Sub CboProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If cboProducto.ListIndex <> -1 Then
        CboImpuesto.SetFocus
    Else
        Call Mensajes(5)
        cboProducto.SetFocus
    End If
End If
End Sub

Private Sub CboProductoB_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   OptActivoB.SetFocus
End If
End Sub

Private Sub CboSucursal_Click()
If CboSucursal.ListIndex <> -1 Then
   TxtNSucursal.Text = CboSucursal.ItemData(CboSucursal.ListIndex)
End If
End Sub

Private Sub CboSucursal_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   If CboSucursal.ListIndex <> -1 Then
      TxtNSucursal.Text = CboSucursal.ItemData(CboSucursal.ListIndex)
      CboEstado.SetFocus
   Else
      Call Mensajes(5)
      CboSucursal.SetFocus
   End If
End If
End Sub

Private Sub CboSucursalB_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   CboProductoB.SetFocus
End If
End Sub

Private Sub CmdAgregar_Click()
Dim i As Integer, StatusCont As String

If cboProd.ListIndex = -1 Then
    Call Mensajes(5)
    cboProd.SetFocus
    Exit Sub
End If

If txtComision.Text = vbNullString Then
    Call Mensajes(5)
    txtComision.SetFocus
    Exit Sub
End If

If Format(DTPFechaContFin.value, "YYYYMMDD") < Format(Date, "YYYYMMDD") Then
    MsgBox "La fecha final no puede ser menor a la fecha actual, Favor de verificar", vbInformation, "Vale Total"
    DTPFechaContFin.SetFocus
    Exit Sub
End If

For i = 1 To GrdContrato.Rows - 1
    If cboProd.ItemData(cboProd.ListIndex) = GrdContrato.TextMatrix(i, 0) Then
        MsgBox "Ya esta activo este producto, Favor de verificar", vbInformation, "Vale Total"
        cboProd.SetFocus
        Exit Sub
    End If
Next

'Select Case CboProd.ItemData(CboProd.ListIndex)
'    Case gBONODESPENSA, gBONOALIMENTACION, gBONOUNIFORME, gPAGOBONOBE, gUNIFORMEBE:
'        If CboProducto.ItemData(CboProducto.ListIndex) = gBONOGASOLINA Or CboProducto.ItemData(CboProducto.ListIndex) = gPAGOGASBE Then
'            MsgBox "No puedes dar de alta el producto " & CboProd.Text & " para este establecimiento, Favor de verificar", vbInformation, "Vale Total"
'            CboProd.SetFocus
'            Exit Sub
'        End If
'    Case gBONOGASOLINA:
'        If CboProducto.ItemData(CboProducto.ListIndex) <> gBONOGASOLINA Or CboProducto.ItemData(CboProducto.ListIndex) <> gPAGOGASBE Then
'            MsgBox "No puedes dar de alta el producto " & CboProd.Text & " para este establecimiento, Favor de verificar", vbInformation, "Vale Total"
'            CboProd.SetFocus
'            Exit Sub
'        End If
'End Select

If OptContAct.value = True Then
    StatusCont = "A"
End If

If OptContInact.value = True Then
    StatusCont = "I"
End If
GrdContrato.AddItem cboProd.ItemData(cboProd.ListIndex) & vbTab & cboProd.Text & vbTab & txtComision.Text & vbTab & _
    Format(DTPFechaContIni.value, "dd/mm/yyyy") & vbTab & Format(DTPFechaContFin.value, "dd/mm/yyyy") & vbTab & StatusCont

cboProd.ListIndex = -1
txtComision.Text = vbNullString
DTPFechaContIni.value = Format(Date, "dd/mm/yyyy")
DTPFechaContFin.value = Format(Date, "dd/mm/yyyy")
OptContAct.value = False
OptContInact.value = False

End Sub

Private Sub CmdBuscar_Click()
Dim OpcionBusca As Integer, Folio As String, Nombre As String, Sucursal As String, Status As String
Dim StrStatus As String, Producto As String

On Error GoTo errbonos
'**********************************************************************************************************************************************************************
If TxtClaveB.Text = vbNullString And TxtNombreB.Text = vbNullString And CboSucursalB.ListIndex = -1 And OptActivoB.value = False And OptInactivoB.value = False Then
   Call Mensajes(5)
   TxtClaveB.SetFocus
   Exit Sub
End If
'**********************************************************************************************************************************************************************
If TxtClaveB.Text <> vbNullString Then
   Folio = TxtClaveB.Text
Else
    Folio = "Null"
End If

If TxtNombreB.Text <> vbNullString Then
    Nombre = "'" & Trim$(TxtNombreB.Text) & "'"
Else
    Nombre = "Null"
End If

If CboSucursalB.ListIndex <> -1 Then
   Sucursal = CboSucursalB.ItemData(CboSucursalB.ListIndex)
Else
    Sucursal = "Null"
End If

If CboProductoB.ListIndex <> -1 Then
   Producto = CboProductoB.ItemData(CboProductoB.ListIndex)
Else
    Producto = "Null"
End If

If OptActivoB.value = True Then
   Status = 1
End If

If OptInactivoB.value = True Then
   Status = 0
End If



sqls = "Exec Sp_Establecimiento_Sel " & Folio & "," & Nombre & "," & Sucursal & "," & Producto & "," & Val(Status) & ""
consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
Grid.Rows = 1
With consulta
   While Not .EOF
      If consulta!Status = 1 Then
         StrStatus = "ACTIVO"
      Else
         StrStatus = "INACTIVO"
      End If
      Grid.AddItem !CveEstablecimiento & Chr(9) & Trim$(!descripcion) + " " + Trim$(!APaterno & "") + " " + Trim$(!AMaterno & "") & Chr(9) & StrStatus
      .MoveNext
   Wend
End With

If consulta.State = 1 Then
   consulta.Close
End If

Exit Sub

errbonos:
Call doErrorLog(1, "OPE", ERR.Number, ERR.Description, Usuario, "frmCatEstablecimientos.CmdBuscarClick")
Call Mensajes(6)
End Sub

Private Sub CmdGrabaDes_Click()
On Error GoTo errbonos
Call Mensajes(0)
If RespMsg = vbYes Then
    If OpcionDet = 0 Then
        sqls = "Exec Sp_EstablecimientosDesp_Ins " & CDbl(LblFolio.Caption) & "," & 0 & ",'" & Trim$(TxtNomComer.Text) & "','" & _
            Trim$(TxtTpv.Text) & "','" & Trim$(TxtAfiliado.Text) & "'"
'        consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        sqls = "Exec Sp_EstablecimientosDesp_Upd " & CDbl(LblFolio.Caption) & "," & 0 & ",'" & Trim$(TxtNomComer.Text) & "','" & _
            Trim$(TxtTpv.Text) & "','" & Trim$(TxtAfiliado.Text) & "'"
'        consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    MsgBox "Información Guardada", vbInformation, "Vale Total"
End If
    
Exit Sub

errbonos:
Call doErrorLog(1, "OPE", ERR.Number, ERR.Description, Usuario, "frmCatEstablecimientos.GuardarDesp")
Call Mensajes(6)
End Sub

Private Sub CmdGrabaGas_Click()
On Error GoTo errbonos

Call Mensajes(0)
If RespMsg = vbYes Then
    If OpcionDet = 0 Then
        sqls = "Exec Sp_EstablecimientosGas_Ins " & CDbl(LblFolio.Caption) & ",'" & Trim$(TxtGpo.Text) & "','" & Trim$(TxtNGpo.Text) & "','" & Trim$(TxtServicio.Text) & "'"
'        consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        sqls = "Exec Sp_EstablecimientosGas_Upd " & CDbl(LblFolio.Caption) & ",'" & Trim$(TxtGpo.Text) & "','" & Trim$(TxtNGpo.Text) & "','" & Trim$(TxtServicio.Text) & "'"
'        consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    MsgBox "Información Guardada", vbInformation, "Vale Total"
End If

Exit Sub

errbonos:
Call doErrorLog(1, "OPE", ERR.Number, ERR.Description, Usuario, "frmCatEstablecimientos.cmdGrabaGas")
Call Mensajes(6)
End Sub

Private Sub CmdLimpiar_Click()
Call Inicio
End Sub

Private Sub CmdSucursales_Click()
'FrmEstablecimientosDet.LblNEstableB.Caption = Trim$(LblFolio.Caption)
'FrmEstablecimientosDet.LblEstableB.Caption = Trim$(TxtNombre.Text)
'FrmEstablecimientosDet.Show vbModal
End Sub

Private Sub DTPFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   cboProducto.SetFocus
End If
End Sub

Private Sub Form_Load()
Call CargaEstados(CboEstado)
Call CargaGiros(CboGiro)
Call CargaBodegas(CboSucursal)
Call CargaBodegas(CboSucursalB)
Call CargaComboBE(cboProducto, "sp_Producto_All")
Call CargaComboBE(cboProd, "sp_Producto_All")
Call CargaComboBE(CboProductoB, "sp_Producto_All")
Call CargaComboBE(cboBanco, "sp_Bancos_Sel")
Call CargaComboBE(CboTipoPago, "sp_Claves_Sel 'Establecimientos','Tipo Pago'")
Call Inicio
Call CentraFormaMDI(Me)
End Sub

Private Sub Inicio()
FraBusqueda.Enabled = True

TxtClaveB.Text = vbNullString
TxtNombreB.Text = vbNullString
CboSucursalB.ListIndex = -1
CboProductoB.ListIndex = -1
OptActivoB.value = True
OptInactivoB.value = False

Grid.Cols = 3
Grid.Rows = 1
Grid.FormatString = "Folio|Nombre|Status"
Grid.ColWidth(0) = 900
Grid.ColWidth(1) = 4600
Grid.ColWidth(2) = 1200

Grid.ColAlignment(1) = 1

SSTab.Tab = 0
OptFisica.value = False
OptMoral.value = True
LblFolio.Caption = vbNullString
TxtNombre.Text = vbNullString
TxtAPaterno.Text = vbNullString
TxtAMaterno.Text = vbNullString
TxtComercial.Text = vbNullString
TxtEmail.Text = vbNullString
CboSucursal.ListIndex = -1
CboEstado.ListIndex = -1
CboMunicipio.ListIndex = -1
TxtColonia.Text = vbNullString
TxtCP.Text = vbNullString
txtTel.Text = vbNullString
txtRFC.Text = vbNullString
TxtCurp.Text = vbNullString
TxtEncargado.Text = vbNullString
txtComision.Text = vbNullString
TxtNSucursal.Text = vbNullString
txtDomicilio.Text = vbNullString
TxtNoExt.Text = vbNullString
TxtNoInt.Text = vbNullString
CboGiro.ListIndex = -1
DTPFecha.value = Format(Date, "dd/mm/yyyy")
TxtGpo.Text = vbNullString
TxtServicio.Text = vbNullString
TxtNGpo.Text = vbNullString
CboImpuesto.ListIndex = -1
OptActivo.value = True
TxtNomComer.Text = vbNullString
TxtTpv.Text = vbNullString
TxtAfiliado.Text = vbNullString
FraDes.Visible = False
FraGas.Visible = False
FraCaptura.Enabled = False
fraContrato.Enabled = False
FraBanco.Enabled = False
FOpcion = 0
OpcionDet = 0
cboProducto.ListIndex = -1
txtComision.Text = 0 '--Eliminar
cboProd.ListIndex = -1
txtComision.Text = vbNullString
DTPFechaContIni.value = Format(Date, "dd/mm/yyyy")
DTPFechaContFin.value = Format(Date, "dd/mm/yyyy")
OptContAct.value = False
OptContInact.value = False
TxtDC.Text = vbNullString

GrdContrato.Cols = 6
GrdContrato.Rows = 1
GrdContrato.FormatString = "|Producto|Comisión|Fecha Inicio|Fecha Fin|Status"
GrdContrato.ColWidth(0) = 0
GrdContrato.ColWidth(1) = 2200
GrdContrato.ColWidth(2) = 1000
GrdContrato.ColWidth(3) = 1400
GrdContrato.ColWidth(4) = 1400
GrdContrato.ColWidth(5) = 600

GrdContrato.ColAlignment(5) = flexAlignCenterCenter


cboBanco.ListIndex = -1
TxtSucursal.Text = vbNullString
CboTipoPago.ListIndex = -1
TxtCuenta.Text = vbNullString

FrmCatEEstablecimientos.Height = 7665
SSTab.Height = 6645
FraCaptura.Height = 6210
End Sub

Private Sub Nuevo()
'Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gGUARDAR)
FraBusqueda.Enabled = False
TxtClaveB.Text = vbNullString
TxtNombreB.Text = vbNullString
CboSucursalB.ListIndex = -1
OptActivoB.value = True
Grid.Rows = 1

FraCaptura.Enabled = True
fraContrato.Enabled = True
FraBanco.Enabled = True
SSTab.Tab = 1
LblFolio.Caption = vbNullString
TxtNombre.Text = vbNullString
TxtAPaterno.Text = vbNullString
TxtAMaterno.Text = vbNullString
TxtComercial.Text = vbNullString
TxtEmail.Text = vbNullString
CboSucursal.ListIndex = -1
CboEstado.ListIndex = -1
CboMunicipio.ListIndex = -1
txtDomicilio.Text = vbNullString
TxtNoExt.Text = vbNullString
TxtNoInt.Text = vbNullString
TxtColonia.Text = vbNullString
TxtCP.Text = vbNullString
txtTel.Text = vbNullString
txtRFC.Text = vbNullString
TxtCurp.Text = vbNullString
TxtEncargado.Text = vbNullString
txtComision.Text = vbNullString
CboGiro.ListIndex = -1
CboImpuesto.ListIndex = -1
TxtServicio.Text = vbNullString
TxtGpo.Text = vbNullString
TxtNGpo.Text = vbNullString
TxtComAnt.Visible = True
OptActivo.value = True
TxtDC.Text = vbNullString
cboBanco.ListIndex = -1
TxtSucursal.Text = vbNullString
CboTipoPago.ListIndex = -1
TxtCuenta.Text = vbNullString

FOpcion = 0
sqls = "Exec Sp_Folio_Sel_Upd 'SEL',0,'EST'"
consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
If Not consulta.EOF Then
   LblFolio.Caption = consulta!Folio
End If

If consulta.State = 1 Then
   consulta.Close
End If

TxtNombre.SetFocus
End Sub

Private Sub Editar()
'Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gGUARDAR)
FraBusqueda.Enabled = False
TxtClaveB.Text = vbNullString
TxtNombreB.Text = vbNullString
CboSucursalB.ListIndex = -1
OptActivoB.value = True
Grid.Rows = 1
FOpcion = 1
FraCaptura.Enabled = True
fraContrato.Enabled = True
FraBanco.Enabled = True
SSTab.Tab = 1
TxtNombre.SetFocus
TxtComAnt.Visible = True
End Sub

Private Sub Guardar()
Dim Status As Integer, Servicio As String, Grupo As String, NomGpo As String ', Contrato As Integer
Dim TipoPersona As Integer, NoExt As String, NoInt As String, APaterno As String, AMaterno As String
Dim Comision As Double, i As Integer, StatusCont As Integer, Sucursal As String, TipoPago As Integer, Cuenta As String
Dim Banco As Integer

On Error GoTo errbonos

If TxtNombre.Text = vbNullString Then
   Call Mensajes(5)
   TxtNombre.SetFocus
   Exit Sub
End If

If OptFisica.value = True Then
    TipoPersona = 1
    If TxtAPaterno.Text = vbNullString Then
        Call Mensajes(5)
        TxtAPaterno.SetFocus
        Exit Sub
    Else
        APaterno = TxtAPaterno.Text
    End If
    If TxtAMaterno.Text = vbNullString Then
        Call Mensajes(5)
        TxtAMaterno.SetFocus
        Exit Sub
    Else
        AMaterno = TxtAMaterno.Text
    End If
Else
    TipoPersona = 2
    APaterno = ""
    AMaterno = ""
End If

If TxtComercial.Text = vbNullString Then
   Call Mensajes(5)
   TxtComercial.SetFocus
   Exit Sub
End If

If CboSucursal.ListIndex = -1 Then
   Call Mensajes(5)
   CboSucursal.SetFocus
   Exit Sub
End If

If CboEstado.ListIndex = -1 Then
   Call Mensajes(5)
   CboEstado.SetFocus
   Exit Sub
End If

If CboMunicipio.ListIndex = -1 Then
   Call Mensajes(5)
   CboMunicipio.SetFocus
   Exit Sub
End If

If txtDomicilio.Text = vbNullString Then
   Call Mensajes(5)
   txtDomicilio.SetFocus
   Exit Sub
End If

If TxtNoExt.Text = vbNullString Then
    NoExt = "S\N"
Else
    NoExt = TxtNoExt.Text
End If

If TxtNoInt.Text = vbNullString Then
    NoInt = "S\N"
Else
    NoInt = TxtNoInt.Text
End If

If TxtColonia.Text = vbNullString Then
   Call Mensajes(5)
   TxtColonia.SetFocus
   Exit Sub
End If

If TxtCP.Text = vbNullString Then
   Call Mensajes(5)
   TxtCP.SetFocus
   Exit Sub
End If

If txtTel.Text = vbNullString Then
   Call Mensajes(5)
   txtTel.SetFocus
   Exit Sub
End If

If txtRFC.Text = vbNullString Then
   Call Mensajes(5)
   txtRFC.SetFocus
   Exit Sub
End If

If TxtCurp.Text = vbNullString Then
   Call Mensajes(5)
   TxtCurp.SetFocus
   Exit Sub
End If

If TxtEncargado.Text = vbNullString Then
   Call Mensajes(5)
   TxtEncargado.SetFocus
   Exit Sub
End If

If CboGiro.ListIndex = -1 Then
   Call Mensajes(5)
   CboGiro.SetFocus
   Exit Sub
End If

If TxtServicio = vbNullString Then
    Servicio = 0
Else
    Servicio = Trim$(TxtServicio.Text)
End If

If TxtGpo.Text = vbNullString Then
    Grupo = "Null"
Else
    Grupo = Trim$(TxtGpo.Text)
End If

If TxtNGpo.Text = vbNullString Then
    NomGpo = "Null"
Else
    NomGpo = Trim$(TxtNGpo.Text)
End If

If CboImpuesto.ListIndex = -1 Then
   Call Mensajes(5)
   CboImpuesto.SetFocus
   Exit Sub
End If

If OptActivo.value = True Then
   Status = 1 'Activo
Else
   Status = 0 'Inactivo
End If

If TxtDC.Text = vbNullString Then
    Call Mensajes(5)
    TxtDC.SetFocus
    Exit Sub
End If

If cboBanco.ListIndex = -1 Then
    MsgBox "Falta capturar el banco", vbInformation, "Vale Total"
    cboBanco.SetFocus
    Exit Sub
Else
    Banco = cboBanco.ItemData(cboBanco.ListIndex)
End If

If cboProducto.ListIndex = -1 Then
    MsgBox "Falta capturar el producto", vbInformation, "Vale Total"
    cboProducto.SetFocus
    Exit Sub
End If

If TxtSucursal.Text = vbNullString Then
    Sucursal = "S/S"
Else
    Sucursal = TxtSucursal.Text
End If

If CboTipoPago.ListIndex = -1 Then
    TipoPago = 0
Else
    TipoPago = CboTipoPago.ItemData(CboTipoPago.ListIndex)
End If

If TxtCuenta.Text = vbNullString Then
    Cuenta = "S/C"
Else
    Cuenta = TxtCuenta.Text
End If


Contrato = 0 '- -Reactivar
Comision = Val(TxtComAnt.Text) '-- poner igual a cero

If FOpcion = 0 Then
   Call Mensajes(0)
   If RespMsg = vbYes Then
        sqls = "Exec Sp_Establecimiento_Ins " & CboSucursal.ItemData(CboSucursal.ListIndex) & "," & Val(LblFolio.Caption) & "," & TipoPersona & ",'" & Trim$(TxtNombre.Text) & "','" & _
          Trim$(APaterno) & "','" & Trim$(AMaterno) & "','" & Trim$(TxtComercial.Text) & "','" & Trim$(TxtEmail.Text) & "','" & Trim$(txtRFC.Text) & " ','" & Trim$(TxtCurp.Text) & "','" & _
          Trim$(txtDomicilio.Text) & "','" & NoExt & "','" & NoInt & "','" & _
          Trim$(TxtColonia.Text) & "','" & TxtCP.Text & "','" & Trim$(txtTel.Text) & "'," & CboMunicipio.ItemData(CboMunicipio.ListIndex) & "," & _
           CboEstado.ItemData(CboEstado.ListIndex) & ",'" & Trim$(TxtEncargado.Text) & "'," & Comision & "," & CboGiro.ItemData(CboGiro.ListIndex) & ",'" & Format(DTPFecha.value, "YYYYMMDD") & "'," & _
           Status & "," & CboImpuesto.ItemData(CboImpuesto.ListIndex) & "," & cboProducto.ItemData(cboProducto.ListIndex) & "," & Contrato & "," & Val(TxtDC.Text) & "," & Banco & ",'" & _
           Trim$(Sucursal) & "'," & TipoPago & ",'" & Trim$(Cuenta) & "'"
        consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If consulta.State = 1 Then
           consulta.Close
        End If
        
        sqls = "Exec Sp_Folio_Sel_Upd 'UPD',0,'EST'," & Val(LblFolio.Caption) & ""
        consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If consulta.State = 1 Then
           consulta.Close
        End If
      
        If GrdContrato.Rows > 1 Then
            For i = 1 To GrdContrato.Rows - 1
                If GrdContrato.TextMatrix(i, 5) = "A" Then
                    StatusCont = 1
                Else
                    StatusCont = 0
                End If
                sqls = "Exec Sp_EstableCont_Ins " & LblFolio.Caption & "," & GrdContrato.TextMatrix(i, 2) & "," & _
                    GrdContrato.TextMatrix(i, 0) & ",'" & Format(GrdContrato.TextMatrix(i, 3), "YYYYMMDD") & "','" & _
                    Format(GrdContrato.TextMatrix(i, 4), "YYYYMMDD") & "'," & StatusCont & ""
                consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
            Next
        End If

        If consulta.State = 1 Then
           consulta.Close
        End If
   End If
Else
   Call Mensajes(2)
   If RespMsg = vbYes Then
      sqls = "Exec Sp_Establecimiento_Upd " & CboSucursal.ItemData(CboSucursal.ListIndex) & "," & Val(LblFolio.Caption) & "," & TipoPersona & ",'" & Trim$(TxtNombre.Text) & "','" & _
            Trim$(APaterno) & "','" & Trim$(AMaterno) & "','" & Trim$(TxtComercial.Text) & "','" & Trim$(TxtEmail.Text) & "','" & Trim$(txtRFC.Text) & "','" & Trim$(TxtCurp.Text) & "','" & _
            Trim$(txtDomicilio.Text) & "','" & NoExt & "','" & NoInt & "','" & _
            Trim$(TxtColonia.Text) & "','" & TxtCP.Text & "','" & Trim$(txtTel.Text) & "'," & CboMunicipio.ItemData(CboMunicipio.ListIndex) & "," & _
            CboEstado.ItemData(CboEstado.ListIndex) & ",'" & Trim$(TxtEncargado.Text) & "'," & Comision & "," & CboGiro.ItemData(CboGiro.ListIndex) & ",'" & Format(DTPFecha.value, "YYYYMMDD") & "'," & _
            Status & "," & CboImpuesto.List(CboImpuesto.ListIndex) & "," & cboProducto.ItemData(cboProducto.ListIndex) & "," & Contrato & "," & Val(TxtDC.Text) & "," & Banco & ",'" & _
            Trim$(Sucursal) & "'," & TipoPago & ",'" & Trim$(Cuenta) & "','" & Usuario & "'"
      consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
      
        If GrdContrato.Rows > 1 Then
            For i = 1 To GrdContrato.Rows - 1
                If GrdContrato.TextMatrix(i, 5) = "A" Then
                    StatusCont = 1
                Else
                    StatusCont = 0
                End If
                sqls = "Exec Sp_EstableCont_Ins " & LblFolio.Caption & "," & GrdContrato.TextMatrix(i, 2) & "," & _
                    GrdContrato.TextMatrix(i, 0) & ",'" & Format(GrdContrato.TextMatrix(i, 3), "YYYYMMDD") & "','" & _
                    Format(GrdContrato.TextMatrix(i, 4), "YYYYMMDD") & "'," & StatusCont & ",'" & Usuario & "'"
                consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
            Next
        End If
        
        If consulta.State = 1 Then
           consulta.Close
        End If
   End If
End If

If consulta.State = 1 Then
   consulta.Close
End If

Call Mensajes(8)

Call Inicio
Exit Sub

errbonos:
Call doErrorLog(1, "OPE", ERR.Number, ERR.Description, Usuario, "frmCatEstablecimientos.Guardar")
Call Mensajes(6)
End Sub

Private Sub GrdContrato_DblClick()
If GrdContrato.Rows > 1 Then
    Call PosicionaComboEnItemData(cboProd, GrdContrato.TextMatrix(GrdContrato.Row, 0))
    txtComision.Text = GrdContrato.TextMatrix(GrdContrato.Row, 2)
    DTPFechaContIni.value = Format(GrdContrato.TextMatrix(GrdContrato.Row, 3), "dd/mm/yyyy")
    DTPFechaContFin.value = Format(GrdContrato.TextMatrix(GrdContrato.Row, 4), "dd/mm/yyyy")
    If GrdContrato.TextMatrix(GrdContrato.Row, 5) = "A" Then
        OptContAct.value = True
    Else
        OptContInact.value = True
    End If
    
    If GrdContrato.Rows > 2 Then
        GrdContrato.RemoveItem (GrdContrato.Row)
    Else
        GrdContrato.Rows = 1
    End If
End If
End Sub

Private Sub Grid_DblClick()
Dim Municipio As String, StatusCont As String

On Error GoTo errbono

If Grid.Rows > 1 Then
   sqls = "Exec Sp_Establecimiento_Sel " & Grid.TextMatrix(Grid.Row, 0) & ""
   consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   
    If Not consulta.EOF Then
        LblFolio.Caption = consulta!CveEstablecimiento
        If consulta!TipoPersona = 1 Then
            OptFisica.value = True
        Else
            OptMoral.value = True
        End If
        TxtNombre.Text = Trim$(consulta!descripcion)
        TxtAPaterno.Text = Trim$(consulta!APaterno & "")
        TxtAMaterno.Text = Trim$(consulta!AMaterno & "")
        TxtComercial.Text = Trim$(consulta!NomComer & "")
        TxtEmail.Text = Trim$(consulta!Email & "")
        Call PosicionaComboEnItemData(CboSucursal, consulta!Sucursal)
        TxtNSucursal.Text = consulta!Sucursal
        Call PosicionaComboEnItemData(CboEstado, consulta!estado)
        Municipio = consulta!Ciudad
        txtDomicilio.Text = Trim$(consulta!Domicilio)
        TxtNoExt.Text = Trim$(consulta!NoExterior & "")
        TxtNoInt.Text = Trim$(consulta!NoInterior & "")
        TxtColonia.Text = Trim$(consulta!Colonia)
        TxtCP.Text = consulta!CodigoPostal
        txtTel.Text = Trim$(consulta!Telefono)
        txtRFC.Text = Trim$(consulta!Rfc)
        TxtCurp.Text = Trim$(consulta!Curp & "")
        TxtEncargado.Text = Trim$(consulta!Encargado)
        TxtComAnt.Text = consulta!Comision 'Quitar
        Contrato = consulta!Contrato 'Quitar
        If Not IsNull(consulta!Giro) Then
            Call PosicionaComboEnItemData(CboGiro, consulta!Giro)
        End If
               
        DTPFecha.value = Format(consulta!Fecha, "DD/MM/YYYY")
        Call PosicionaComboEnItemData(cboProducto, consulta!TipoBono)
        Call PosicionaComboEnItemData(CboImpuesto, consulta!CveImpuesto)
        If consulta!Status = 1 Then
            OptActivo.value = True
        Else
            OptInactivo.value = True
        End If
        
        TxtDC.Text = Val(consulta!DiasCred)
      
        If Not IsNull(consulta!Banco) Then
            Call PosicionaComboEnItemData(cboBanco, consulta!Banco)
        End If
        
        If Not IsNull(consulta!SucBanco) Then
            If Trim$(consulta!SucBanco) = "S/S" Then
                TxtSucursal.Text = vbNullString
            Else
                TxtSucursal.Text = Trim$(consulta!SucBanco & "")
            End If
        Else
            TxtSucursal.Text = vbNullString
        End If
        
        If Not IsNull(consulta!TipoPago) Then
            If (consulta!TipoPago) = 0 Then
                CboTipoPago.ListIndex = -1
            Else
                Call PosicionaComboEnItemData(CboTipoPago, consulta!TipoPago)
            End If
        Else
            CboTipoPago.ListIndex = -1
        End If
        
        If Not IsNull(consulta!Cuenta) Then
            If Trim$(consulta!Cuenta) = "S/C" Then
                TxtCuenta.Text = vbNullString
            Else
                TxtCuenta.Text = Trim$(consulta!Cuenta & "")
            End If
        Else
            TxtCuenta.Text = vbNullString
        End If
        
        If consulta.State = 1 Then
            consulta.Close
        End If
        
        GrdContrato.Rows = 1
        
        sqls = "Exec Sp_EstableCont_Sel " & Grid.TextMatrix(Grid.Row, 0) & ""
        consulta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        With consulta
            While Not .EOF
                If consulta!Status = 1 Then
                    StatusCont = "A"
                Else
                    StatusCont = "I"
                End If
                GrdContrato.AddItem !Producto & vbTab & Trim$(!Bon_Pro_Descripcion) & vbTab & !Comision & vbTab & Format(!FechaInicio, "dd/mm/yyyy") & vbTab & Format(!FechaFinal, "dd/mm/yyyy") & vbTab & StatusCont
                .MoveNext
            Wend
        End With
               
        If consulta.State = 1 Then
            consulta.Close
        End If
        
      '  Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gEDITAR)
        FOpcion = 1
        Grid.SetFocus
        Grid.Row = 0
    Else
        MsgBox "No hay informacion para el establecimiento" & Grid.TextMatrix(Grid.Row, 0), vbInformation, "Vale Total"
        If consulta.State = 1 Then
            consulta.Close
        End If
        Exit Sub
    End If
   
    If consulta.State = 1 Then
        consulta.Close
    End If
   
    Call CargaPoblaciones(CboMunicipio, CboEstado.ItemData(CboEstado.ListIndex))
    Call PosicionaComboEnItemData(CboMunicipio, Val(Municipio))
   
    If consulta.State = 1 Then
        consulta.Close
    End If
    
  '  If cboProducto.ItemData(cboProducto.ListIndex) = 1 Then
  '      sqls = "Exec Sp_EstablecimientosDesp_Sel " & CDbl(LblFolio.Caption) & ""
   '     consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
   '     If Not consulta.EOF Then
   '         If Not IsNull(consulta!Nombre) Then
   '             TxtNomComer.Text = Trim$(consulta!Nombre)
   '         Else
   '             TxtNomComer.Text = vbNullString
   '         End If
   '         If Not IsNull(consulta!Tpv) Then
   '             TxtTpv.Text = Trim$(consulta!Tpv)
   '         Else
   '             TxtTpv.Text = vbNullString
   '         End If
   '         If Not IsNull(consulta!Afiliado) Then
   '             TxtAfiliado.Text = Trim$(consulta!Afiliado)
   '         Else
   '             TxtAfiliado.Text = vbNullString
   '         End If
   '         OpcionDet = 1
   '     End If
   '     FrmCatEEstablecimientos.Height = 9360
   '     SSTab.Height = 8325
   '     FraCaptura.Height = 7890
   '
   '     FraDes.Visible = True
   '     FraDes.Top = 6195
   ' Else
   '     sqls = "Exec Sp_EstablecimientosGas_Sel " & CDbl(LblFolio.Caption) & ""
   '     consulta.Open sqls, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
   '     If Not consulta.EOF Then
   '         If Not IsNull(consulta!Grupo) Then
   '             TxtGpo.Text = Trim$(consulta!Grupo)
   '         Else
   '             TxtGpo.Text = vbNullString
   '         End If
   '         If Not IsNull(consulta!Servicio) Then
   '             TxtServicio.Text = Trim$(consulta!Servicio)
   '         Else
   '             TxtServicio.Text = vbNullString
   '         End If
   '         If Not IsNull(consulta!NomGrupo) Then
   '             TxtNGpo.Text = Trim$(consulta!NomGrupo)
   '         Else
   '             TxtNGpo.Text = vbNullString
   '         End If
   '         OpcionDet = 1
   '     End If
        FrmCatEEstablecimientos.Height = 9360
        SSTab.Height = 8325
        FraCaptura.Height = 7890
        
   '     FraGas.Visible = True
   '     FraDes.Top = 6195
    'End If
    
   SSTab.Tab = 1
   FOpcion = 1
End If

If consulta.State = 1 Then
    consulta.Close
End If

Exit Sub

errbono:
Call doErrorLog(1, "OPE", ERR.Number, ERR.Description, Usuario, "frmCatEstablecimientos.GridDblClick")
Call Mensajes(6)
End Sub

Private Sub OptFisica_Click()
If OptFisica.value = True Then
    TxtAPaterno.Enabled = True
    TxtAMaterno.Enabled = True
End If
End Sub

Private Sub OptMoral_Click()
If OptMoral.value = True Then
    TxtAPaterno.Enabled = False
    TxtAMaterno.Enabled = False
End If
End Sub

Private Sub TxtAfiliado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtAMaterno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtAMaterno.Text <> vbNullString Then
        TxtNSucursal.SetFocus
    Else
        Call Mensajes(5)
        TxtAMaterno.SetFocus
    End If
End If
End Sub

Private Sub TxtAPaterno_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtAPaterno.Text <> vbNullString Then
        TxtAMaterno.SetFocus
    Else
        Call Mensajes(5)
        TxtAPaterno.SetFocus
    End If
End If
End Sub

Private Sub TxtClaveB_KeyPress(KeyAscii As Integer)
If CapNumerica(KeyAscii) = True Then
   If KeyAscii = vbKeyReturn Then
      TxtNombreB.SetFocus
   End If
Else
   KeyAscii = 0
End If
End Sub

Private Sub TxtColonia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If TxtColonia.Text <> vbNullString Then
      TxtCP.SetFocus
   Else
      Call Mensajes(5)
      TxtColonia.SetFocus
   End If
End If
End Sub


Private Sub TxtComercial_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtComercial.Text <> vbNullString Then
        TxtEmail.SetFocus
    Else
        Call Mensajes(5)
        TxtComercial.SetFocus
    End If
End If
End Sub

Private Sub TxtComision_KeyPress(KeyAscii As Integer)
If CapImporte(KeyAscii) = True Then
   If KeyAscii = vbKeyReturn Then
      If txtComision.Text <> vbNullString Then
         DTPFechaContIni.SetFocus
      Else
         Call Mensajes(5)
         txtComision.SetFocus
      End If
   End If
Else
   KeyAscii = 0
End If
End Sub

Private Sub TxtComisionBE_KeyPress(KeyAscii As Integer)
If CapImporte(KeyAscii) = False Then
    KeyAscii = 0
End If
End Sub

Private Sub TxtCP_KeyPress(KeyAscii As Integer)
If CapNumerica(KeyAscii) = True Then
   If KeyAscii = vbKeyReturn Then
      If TxtCP.Text <> vbNullString Then
         txtTel.SetFocus
      Else
         Call Mensajes(5)
         TxtCP.SetFocus
      End If
   End If
Else
   KeyAscii = 0
End If
End Sub

Private Sub TxtDiasCred_KeyPress(KeyAscii As Integer)
If CapNumerica(KeyAscii) = False Then
    KeyAscii = 0
End If
End Sub

Private Sub TxtCurp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If TxtCurp.Text <> vbNullString Then
      TxtEncargado.SetFocus
   Else
      Call Mensajes(5)
      TxtCurp.SetFocus
   End If
End If
End Sub

Private Sub TxtDomicilio_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If txtDomicilio.Text <> vbNullString Then
      TxtNoExt.SetFocus
   Else
      Call Mensajes(5)
      txtDomicilio.SetFocus
   End If
End If
End Sub

Private Sub TxtEncargado_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If TxtEncargado.Text <> vbNullString Then
      CboGiro.SetFocus
   Else
      Call Mensajes(5)
      TxtEncargado.SetFocus
   End If
End If
End Sub

Private Sub TxtGpo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtServicio.SetFocus
End If
End Sub
Private Sub TxtNGpo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNoExt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtNoInt.SetFocus
End If
End Sub

Private Sub TxtNoInt_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtColonia.SetFocus
End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtNombre.Text <> vbNullString Then
        If TxtAPaterno.Enabled = True Then
            TxtAPaterno.SetFocus
        Else
        TxtNSucursal.SetFocus
        End If
    Else
        Call Mensajes(5)
        TxtNombre.SetFocus
    End If
End If
End Sub

Private Sub TxtNombreB_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   CboSucursalB.SetFocus
End If
End Sub

Private Sub TxtNomComer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtTpv.SetFocus
End If
End Sub

Private Sub TxtNSucursal_KeyPress(KeyAscii As Integer)
If CapNumerica(KeyAscii) = True Then
   If KeyAscii = vbKeyReturn Then
      If TxtNSucursal.Text <> vbNullString Then
         Call PosicionaComboEnItemData(CboSucursal, TxtNSucursal.Text)
      End If
      CboSucursal.SetFocus
   End If
 Else
   KeyAscii = 0
 End If
   
End Sub
Private Sub TxtRFC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If txtRFC.Text <> vbNullString Then
      TxtEncargado.SetFocus
   Else
      Call Mensajes(5)
      txtRFC.SetFocus
   End If
End If
End Sub
Private Sub TxtServicio_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtNGpo.SetFocus
End If
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
   If txtTel.Text <> vbNullString Then
      txtRFC.SetFocus
   Else
      Call Mensajes(5)
      txtTel.SetFocus
   End If
End If
End Sub

Private Sub Imprimir()
'If CboSucursalB.ListIndex = -1 Then
'    Call Mensajes(5)
'    CboSucursalB.SetFocus
'    Exit Sub
'End If

'plazasuc = CboSucursalB.ItemData(CboSucursalB.ListIndex)
'Reporte = "RptEstablecimientos"
'FrmReporte.Show
End Sub

Private Sub TxtTpv_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtAfiliado.SetFocus
End If
End Sub
