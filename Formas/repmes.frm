VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepMes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro de comisiones"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FRM 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   345
      TabIndex        =   15
      Top             =   3720
      Width           =   2535
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "repmes.frx":0000
         Left            =   480
         List            =   "repmes.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
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
         ItemData        =   "repmes.frx":005A
         Left            =   120
         List            =   "repmes.frx":0064
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame frmbodegas 
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   2535
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   80
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   1440
         Picture         =   "repmes.frx":0076
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   720
         Picture         =   "repmes.frx":01E8
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (mm/dd/yyyy) "
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   840
         TabIndex        =   4
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   240
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
         TabIndex        =   5
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   720
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
         Left            =   360
         TabIndex        =   3
         Top             =   795
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame frmDetalle 
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   2535
      Begin VB.CheckBox chkDetalle 
         Caption         =   "Detalle de las no conciliadas"
         Height          =   255
         Left            =   55
         TabIndex        =   13
         Top             =   240
         Width           =   2350
      End
   End
End
Attribute VB_Name = "frmRepMes"
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

Private Sub cmdImprimir_Click()
    Imprime crptToPrinter
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
Dim ms As Byte, uf As Date
Dim fe As String, fe2 As String
    Set mclsAniform = New clsAnimated
    
    CentraForma Me
    CargaBodegas cboBodegas
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")

    Me.Height = 4110
    frmbodegas.Visible = True
    frmDetalle.Visible = False
     
    mskFechaIni = "01/" & Mid(Date, 4, 2) & "/" & Format(Date, "yyyy")
    ms = Month(Date) + 1
    'ms = IIf(ms > 12, 1, ms)
    If ms > 12 Then
       ms = 1
       uf = "01/" & Format(ms, "00") & "/" & (Mid(Date, 7, 4) + 1)
    Else
       uf = "01/" & Format(ms, "00") & "/" & Mid(Date, 7, 4)
    End If
    uf = uf - 1
    mskFechaFin = Format(uf, "dd/mm/yyyy")
    
    If TipoRep = "PAPEL" Then
        Me.Caption = "COMISIONES A PROV. PAPEL"
    ElseIf TipoRep = "BE" Then
        Me.Caption = "COMISIONES A PROVEEDORES"
    ElseIf TipoRep = "FACDISP" Then
        Me.Caption = "Pedidos por Fecha Disp"
        mskFechaIni = Date + 1
        mskFechaFin = Date + 1
    ElseIf TipoRep = "CONCTRAN" Then
        Me.Caption = "Conciliacion de Transacciones"
        frmbodegas.Visible = False
        frmDetalle.Visible = True
    ElseIf TipoRep = "CXP" Then
        Me.Caption = "Resumen de Transacciones"
        frmbodegas.Visible = False
    ElseIf TipoRep = "CBCO" Then
        Me.Caption = "Cargos del Banco"
        frmbodegas.Visible = False
     ElseIf TipoRep = "FACVSDISPMES" Or TipoRep = "AJUSTESMES" Then
       frmbodegas.Visible = False
       Me.Caption = "Dispersion Mensual"
    ElseIf TipoRep = "DTM" Then
       Me.Caption = "Detalle de Movimientos"
       frmbodegas.Visible = False
    ElseIf TipoRep = "AJUSFONDOS" Then
          Me.Caption = "Ajuste Fondos Insuficientes"
        frmbodegas.Visible = False
    ElseIf TipoRep = "TRANSASNOC" Then
        Me.Caption = "Transacciones no conciliadas"
        frmbodegas.Visible = False
    ElseIf TipoRep = "Status de Solicitudes" Then
        Me.Caption = "Status de Solicitudes"
        Me.Height = 5130
        FRM.Top = 2760
        Frame1.Top = 3720
        fe = Format(Date - 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
        cbostatus.Text = "Todas"
    ElseIf TipoRep = "Estancadas" Then
        Me.Caption = "Tarjetas estancadas"
        fe = Format(Date - 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    ElseIf TipoRep = "CANTRANPROD" Then
        frmbodegas.Visible = False
        Frame4.Visible = False
        Me.Caption = "Cantidad de Transacciones SyC"
    ElseIf TipoRep = "FACTURAS" Then
        Me.Caption = "Libro de Ventas de Dispersiones"
        Frame4.Visible = False
        frmbodegas.Visible = False
        fe = Format(Date - Day(Date) + 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    ElseIf TipoRep = "TARJETAS" Then
        Me.Caption = "Libro de Ventas de Tarjetas"
        Frame4.Visible = False
        frmbodegas.Visible = False
        fe = Format(Date - Day(Date) + 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    ElseIf TipoRep = "LIQXCTE" Then
        Me.Caption = "Liquidaciones por Cliente"
        Frame4.Visible = False
        frmbodegas.Visible = False
        fe = Format(Date - Day(Date) + 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    ElseIf TipoRep = "Notas" Then
        Me.Caption = "Diario Notas de Credito"
        fe = Format(Date - Day(Date) + 1, "mm/dd/yyyy")
        fe2 = Format(Date, "mm/dd/yyyy")
        mskFechaIni = Format(fe, "mm/dd/yyyy")
        mskFechaFin = Format(fe2, "mm/dd/yyyy")
    Else
       Me.Caption = ""
    End If
       
End Sub
Private Sub txtAño_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaNumericos(KeyAscii, txtAño.Text, 0)
End Sub
Sub Imprime(Destino)
Dim Result As Integer
Dim Rango As String
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    
    Rango = "Fecha = 'Del " & Format(mskFechaIni, "dd/mmm/yyyy") & " al " & Format(mskFechaFin, "dd/mmm/yyyy") & "'"
    If TipoRep = "PAPEL" Then
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptLibroCom.rpt"
    ElseIf TipoRep = "BE" Then
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = 1
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptLibroComBE.rpt"
    ElseIf TipoRep = "FACDISP" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptBitacora.rpt"
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = "2"
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(4) = CStr(Product)
    ElseIf TipoRep = "FACVSDISP" Then
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFacVSDisp.rpt"
    ElseIf TipoRep = "CONCTRAN" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptconctransacciones.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = 0
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "CXP" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCXPGlobal.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "CBCO" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCargosBanco.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "FACVSDISPMES" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptfacvsdispmensual.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "AJUSTESMES" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAjustesmensual.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "DTM" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFinTransBE.rpt"
        If Mid(Format(mskFechaIni, "mm/dd/yyyy"), 7, 4) = Mid(Date, 7, 4) Then
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni) - 1, Year(mskFechaIni)), "0000")
           If Year(mdiMain.cryReport.StoredProcParam(2)) <= 2009 Then
              mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin))
           Else
              mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin))
           End If
           mdiMain.cryReport.StoredProcParam(4) = 1
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        Else
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni), Year(mskFechaIni)), "0000")
           mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin))
           mdiMain.cryReport.StoredProcParam(4) = 1
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        End If
    ElseIf TipoRep = "AJUSFONDOS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAjusFondosInsDet.rpt"
        mdiMain.cryReport.StoredProcParam(0) = 0
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "TRANSASNOC" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTransas_Noconc.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
    ElseIf TipoRep = "Status de Solicitudes" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptStatus_Solicitud.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
        mdiMain.cryReport.StoredProcParam(3) = "Reporte"
        mdiMain.cryReport.StoredProcParam(4) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(5) = cbostatus.ListIndex
    ElseIf TipoRep = "Estancadas" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptStatus_Solicitud.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
        mdiMain.cryReport.StoredProcParam(3) = "Reporte"
        mdiMain.cryReport.StoredProcParam(4) = cboBodegas.ItemData(cboBodegas.ListIndex)
    ElseIf TipoRep = "CANTRANPROD" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\RepCantTranProd.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
    ElseIf TipoRep = "FACTURAS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\LibroDisp.rpt"
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = 1
        mdiMain.cryReport.Formulas(0) = Rango
        mdiMain.cryReport.Formulas(1) = "Reporte ='LIBRO DE VENTAS DE DISPERSIONES'"
    ElseIf TipoRep = "TARJETAS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\LibroTarj.rpt"
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = 12
        mdiMain.cryReport.Formulas(0) = Rango
        mdiMain.cryReport.Formulas(1) = "Reporte ='LIBRO DE VENTAS DE TARJETAS'"
    ElseIf TipoRep = "LIQXCTE" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptLiquidacionesxCliente.rpt"
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.Formulas(0) = Rango
        mdiMain.cryReport.Formulas(1) = "Reporte ='LIQUIDACIONES POR CLIENTE'"
    ElseIf TipoRep = "NOTAS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptDiario.rpt"
        mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = 1
        mdiMain.cryReport.Formulas(0) = Rango
        mdiMain.cryReport.Formulas(1) = "Reporte ='DIARIO DE NOTAS DE CREDITO'"
    Else
        MsgBox "Error en reporte, verifiquelo con sistemas", vbCritical, "Errores generados"
        Exit Sub
    End If
    
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
   
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Error..."
    End If
    
    
    If TipoRep = "CONCTRAN" And chkDetalle.value = 1 Then
        MsgBar "Generando Reporte Detallado", True
        Limpia_CryReport
        mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptconctransaccionesDet.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = "1" 'aqki cambie
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
        
        On Error Resume Next
        Result = mdiMain.cryReport.PrintReport
        MsgBar "", False
        If Result <> 0 Then
            MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
        End If
        
    ElseIf TipoRep = "AJUSFONDOS" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAjustesFondosInsuf.rpt"
        mdiMain.cryReport.StoredProcParam(0) = 0
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
     
        On Error Resume Next
        Result = mdiMain.cryReport.PrintReport
        MsgBar "", False
        If Result <> 0 Then
            MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
        End If
  
    End If
    
    If TipoRep = "DTM" Then
        MsgBar "Generando Reporte Tarjetas Saldos Negativos", True
        Limpia_CryReport
        mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTarjetasSaldoNeg.rpt"
        If Mid(Format(mskFechaIni, "mm/dd/yyyy"), 7, 4) = Mid(Date, 7, 4) Then
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni) - 1, Year(mskFechaIni)), "0000")
           mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin) - 1)
           mdiMain.cryReport.StoredProcParam(4) = 2
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        Else
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni), Year(mskFechaIni)), "0000")
           mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin))
           mdiMain.cryReport.StoredProcParam(4) = 2
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        End If
        
     
        On Error Resume Next
        Result = mdiMain.cryReport.PrintReport
        MsgBar "", False
        If Result <> 0 Then
            MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
        End If
        
        MsgBar "Generando Reporte Ventas sin Dispersar", True
        Limpia_CryReport
        mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptVentasinDisp.rpt"
        If Mid(Format(mskFechaIni, "mm/dd/yyyy"), 7, 4) = Mid(Date, 7, 4) Then
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni) - 1, Year(mskFechaIni)), "0000")
           mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin) - 1)
           mdiMain.cryReport.StoredProcParam(4) = 3
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        Else
           mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
           mdiMain.cryReport.StoredProcParam(2) = Format(IIf(Month(mskFechaIni) - 1 = 0, 12, Month(mskFechaIni) - 1), "00") & "/01/" & Format(IIf(Month(mskFechaIni) = 1, Year(mskFechaIni), Year(mskFechaIni)), "0000")
           mdiMain.cryReport.StoredProcParam(3) = FechaFinMes(IIf(Month(mskFechaFin) = 1, 12, Month(mskFechaFin) - 1), Year(mskFechaFin))
           mdiMain.cryReport.StoredProcParam(4) = 3
           'prod = IIf(Product = 8, 6, Product)
           producto_cual
           mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        End If
        
    
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
