VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepSolTarjetas 
   Caption         =   "Solicitudes de Tarjetas"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRepSolTarjetas.frx":0000
         Left            =   120
         List            =   "frmRepSolTarjetas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3120
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   3375
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   2355
         Picture         =   "frmRepSolTarjetas.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblNombre 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   855
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   240
         Picture         =   "frmRepSolTarjetas.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   450
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   240
         Picture         =   "frmRepSolTarjetas.frx":0290
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   3375
      Begin VB.CheckBox chk1 
         Caption         =   "Orden x Solicitud"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmRepSolTarjetas.frx":0392
         Left            =   360
         List            =   "frmRepSolTarjetas.frx":0394
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (mm/dd/yyyy) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
      Begin VB.CheckBox chkEnt 
         Caption         =   "Imprimir reporte entregas"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   960
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   960
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   915
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRepSolTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Dim ordena As Boolean

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Sub InicializaForma()
   lblNombre.Caption = ""
   chkEnt.value = 0
End Sub

Private Sub chk1_Click()
 If chk1.value = 1 Then
    ordena = True
 Else
    ordena = False
 End If
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "ClienteBE"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtCliente = frmConsulta.cliente
       lblNombre = frmConsulta.Nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Private Sub cmdPresentar_Click()
    Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Sub carga_status()
   cbostatus.Clear
   cbostatus.AddItem "Todas"
   cbostatus.AddItem "Solicitadas"
   cbostatus.AddItem "Aceptadas"
   cbostatus.AddItem "En Ruta"
   cbostatus.AddItem "Entregadas"
   cbostatus.AddItem "Facturadas"
   cbostatus.AddItem "Por Facturar"
End Sub

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  
  lblNombre = ""
  mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
  mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YyYY")
 ' carga_status
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  CboProducto.Text = UCase("Winko Mart")
  Carga_Status_Solicitud cbostatus
  cbostatus.ListIndex = 0
  cbostatus.AddItem "Por Facturar"
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "sp_SolicitudesBE_varios '','',''," & Product & ",'Reajusta'"
    cnxbdMty.Execute sqls
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptSolTarjetas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 1
    mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = cbostatus.ItemData(cbostatus.ListIndex) 'cbostatus.ListIndex
    If cbostatus.ItemData(cbostatus.ListIndex) = 0 And cbostatus.Text <> "<< TODAS >>" Then
       mdiMain.cryReport.StoredProcParam(4) = 7
    End If
    
'    If ordena = False Then
'       If cboStatus.ListIndex = 6 Then
'          mdiMain.cryReport.StoredProcParam(4) = 7
'       Else
'          mdiMain.cryReport.StoredProcParam(4) = cboStatus.ListIndex 'cboStatus.ItemData(cboStatus.ListIndex)
'       End If
'    Else
'       mdiMain.cryReport.StoredProcParam(4) = 6
'    End If
    producto_cual
    mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
    mdiMain.cryReport.StoredProcParam(6) = 0
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
    
    If chkEnt.value = 1 Then
        MsgBar "Generando Reporte Entregas", True
        Limpia_CryReport
        mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEntregas.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = 1
        mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(4) = cbostatus.ListIndex 'cboStatus.ItemData(cboStatus.ListIndex)
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
         
        On Error Resume Next
        Result = mdiMain.cryReport.PrintReport
        MsgBar "", False
        If Result <> 0 Then
            MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
