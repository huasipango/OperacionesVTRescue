VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepStatus 
   Caption         =   "Estatus de Clientes"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   248
      TabIndex        =   9
      Top             =   0
      Width           =   3015
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
         ItemData        =   "frmRepStatus.frx":0000
         Left            =   120
         List            =   "frmRepStatus.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (dd/mm/yyyy) "
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2415
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   840
         TabIndex        =   5
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
         Left            =   840
         TabIndex        =   6
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
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   915
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   735
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   120
         Picture         =   "frmRepStatus.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   200
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         Picture         =   "frmRepStatus.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   720
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmRepStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
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

Private Sub Form_Load()
    Dim fe As String, fe2 As String
    Set mclsAniform = New clsAnimated
    
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    mskFechaIni = Format((Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Trim(Str(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))))), "mm/dd/yyyy")
    mskFechaFin = Format((FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date) - 1), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)))), "mm/dd/yyyy")
    
    If TipoRep = "KARDEX" Then
        Me.Caption = "IndicadoresBE"
    ElseIf TipoRep = "CLIENTES" Then
        Me.Caption = "Estatus de Clientes"
    ElseIf TipoRep = "TARJETAS" Then
        Me.Caption = "Entrega de Tarjetas"
    ElseIf TipoRep = "DEPOB" Then
        Me.Caption = "Depositos en Otros Bancos"
    ElseIf TipoRep = "FACTAR" Then
        Me.Caption = "Facturación de Tarjetas"
    ElseIf TipoRep = "TARCANC" Then
        Me.Caption = "Tarjetas Canceladas"
        mskFechaIni = Date - 1
        mskFechaFin = Date - 1
    ElseIf TipoRep = "INDTRAN" Then
        Me.Caption = "Transacciones"
    ElseIf TipoRep = "EDOCTAGLOBAL" Then
        Me.Caption = "Estado de Cuenta Global"
    ElseIf TipoRep = "Status de Solicitudes" Then
        Me.Caption = "Status de Solicitudes"
        fe = Format(Date - 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    End If
     
   ' Call CargaVendedores(cboVendedores)
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    If TipoRep = "CLIENTES" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAltaClienteBE.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "KARDEX" Then
        sqls = "EXEC SP_REPKARDEX '" & Format(mskFechaIni, "mm/dd/yyyy") & "' , '" & Format(mskFechaFin, "mm/dd/yyyy") & "'," & CStr(Product)
        cnxBD.Execute sqls, intRegistros
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\Indicadoresbe.rpt"
    ElseIf TipoRep = "TARJETAS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rpttarjetasEnt.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "DEPOB" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptDepositosOB.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = 1
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "FACTAR" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFacturasTarjetas.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = 2
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "TARCANC" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTarjetasCanc.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "INDTRAN" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptIndTransacciones.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "EDOCTAGLOBAL" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEdoCuentaGlobal.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(Format((FechaFinMes(IIf((Month(Date)) = 1, 11, Month(Date) - 2), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)))), "mm/dd/yyyy"), "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "Status de Solicitudes" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptStatus_Solicitud.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    End If
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Error..."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
