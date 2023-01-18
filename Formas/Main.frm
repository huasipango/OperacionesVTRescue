VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Operaciones Vale Total"
   ClientHeight    =   4125
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   7980
   Icon            =   "Main.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "mdiForm1"
   Picture         =   "Main.frx":1CFA
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1020
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   7920
      TabIndex        =   2
      Top             =   2580
      Visible         =   0   'False
      Width           =   7980
      Begin VB.Image Image1 
         Height          =   615
         Left            =   3480
         Top             =   120
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   480
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ProgressBar prAvance 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3855
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2769
            MinWidth        =   2187
            Text            =   "Listo..."
            TextSave        =   "Listo..."
            Object.ToolTipText     =   "Estado del Sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2011
            MinWidth        =   1764
            TextSave        =   "11/02/2020"
            Object.ToolTipText     =   "Fecha Actual"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1499
            TextSave        =   "04:01 p.m."
            Object.ToolTipText     =   "Hora Actual"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Ver."
            TextSave        =   "Ver."
            Key             =   "Verso"
            Object.ToolTipText     =   "Version oficial del Sistema"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1658
            MinWidth        =   1658
            Key             =   "Usuario"
            Object.ToolTipText     =   "Usuario del sistema"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "PCname"
            Object.ToolTipText     =   "Nombre de la PC"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnOper 
      Caption         =   "&Operación"
      Begin VB.Menu mnuOpeHacienda 
         Caption         =   "&Facturación"
         Begin VB.Menu FacCtes 
            Caption         =   "&Facturas a Clientes"
         End
         Begin VB.Menu mnOI 
            Caption         =   "&Tarjetas y Otros Ingresos"
         End
         Begin VB.Menu FacProv 
            Caption         =   "&Comisiones a Proveedores"
         End
         Begin VB.Menu mnCancFacBE 
            Caption         =   "C&ancelacion de Facturas Dispersiones"
         End
         Begin VB.Menu mnCancOI 
            Caption         =   "Cancelacion de Facturas &Otros Ingresos"
         End
      End
      Begin VB.Menu mnNotas 
         Caption         =   "&Notas de Crédito"
         Begin VB.Menu Mn_PorReem 
            Caption         =   "&Generacion de Notas de Crédito"
         End
         Begin VB.Menu mnCancel 
            Caption         =   "&Cancelación de Notas de Crédito"
         End
      End
      Begin VB.Menu mnNotasCar 
         Caption         =   "Notas de &Cargo"
         Begin VB.Menu mnGeneraNota 
            Caption         =   "&Generar Notas de Cargo"
         End
         Begin VB.Menu mnImpNotascar 
            Caption         =   "&Cancelacion de  Notas de Cargo"
         End
      End
      Begin VB.Menu mnOpCuentas 
         Caption         =   "&Cuentas"
         Begin VB.Menu mnAjustes 
            Caption         =   "Ajustes"
         End
      End
      Begin VB.Menu mnTarjetas 
         Caption         =   "T&arjetas"
         Begin VB.Menu StatTarjetas 
            Caption         =   "Status de Solicitudes de Tarjetas"
         End
         Begin VB.Menu TarjetaEstanca 
            Caption         =   "Tarjetas Estancadas"
         End
      End
      Begin VB.Menu mnEnvios 
         Caption         =   "&Envios"
         Begin VB.Menu mnEnvioTarjetas 
            Caption         =   "Envio de Tarjetas"
         End
         Begin VB.Menu mnCE 
            Caption         =   "Control de Envíos"
         End
      End
      Begin VB.Menu mnGenArchivoBanco12 
         Caption         =   "&Generar archivos para Banco"
         Begin VB.Menu mnGenArchivoBanco3 
            Caption         =   "Archivos SyC"
         End
      End
      Begin VB.Menu mnIntegraLiqBEPaso 
         Caption         =   "Integracion de Archivos del Banco"
      End
   End
   Begin VB.Menu mnSaldos 
      Caption         =   "&Atencion a Clientes"
      Begin VB.Menu mnConsSaldos 
         Caption         =   "&Consulta de Saldos"
      End
      Begin VB.Menu mnAclara 
         Caption         =   "&Aclaraciones"
      End
      Begin VB.Menu mnuMovsCteEMPL 
         Caption         =   "Movs. de empleados x Cliente"
      End
   End
   Begin VB.Menu mnMesa 
      Caption         =   "M&esa de Control"
      Begin VB.Menu mnuIntBanco 
         Caption         =   "Integracion Depositos Banco"
      End
      Begin VB.Menu mnAutPedidos 
         Caption         =   "Autorizar Pedidos"
      End
   End
   Begin VB.Menu mnuComisionbe 
      Caption         =   "Co&misiones"
      Begin VB.Menu mnuIng 
         Caption         =   "Ingresos por Comisiones"
      End
      Begin VB.Menu mnuDepositos 
         Caption         =   "Depositos por Comisiones"
      End
      Begin VB.Menu mnuEdoCTAoi2 
         Caption         =   "Estado de Cuenta Otros Ingresos"
      End
      Begin VB.Menu mnuCorteCaja 
         Caption         =   "Corte de Caja Otros Ingresos"
      End
      Begin VB.Menu mnuAplicacion 
         Caption         =   "Aplicacion de la cobranza"
      End
      Begin VB.Menu mnuValidacion 
         Caption         =   "Validación de la Cobranza Otros Ingresos"
      End
      Begin VB.Menu mnuAntig 
         Caption         =   "Antig. de Saldos Concentrada Otros Ingresos"
      End
   End
   Begin VB.Menu mnCons 
      Caption         =   "C&onsultas"
      Begin VB.Menu mnConPed 
         Caption         =   "&Consulta de Pedidos"
      End
      Begin VB.Menu mnconsultaActBancos 
         Caption         =   "Consulta de Actividades del Banco"
      End
   End
   Begin VB.Menu mnCat 
      Caption         =   "&Catálogos"
      Begin VB.Menu mnEstab 
         Caption         =   "&Establecimientos"
      End
      Begin VB.Menu mnProv 
         Caption         =   "&Comercios"
      End
      Begin VB.Menu mnclientes 
         Caption         =   "C&lientes"
      End
      Begin VB.Menu mnCorreoFE 
         Caption         =   "Correos &Fact. Elect."
      End
      Begin VB.Menu mnclientesOI 
         Caption         =   "Clientes &Otros Ingresos"
      End
   End
   Begin VB.Menu mnRep 
      Caption         =   "&Reportes"
      Begin VB.Menu mncombe 
         Caption         =   "&Comisiones"
         Begin VB.Menu mnTranscom 
            Caption         =   "&Transacciones por Comercio"
         End
         Begin VB.Menu mnTransDet 
            Caption         =   "Transacciones &Detalladas"
         End
         Begin VB.Menu mnRepLibroComBE 
            Caption         =   "&Libro de Comisiones"
         End
         Begin VB.Menu mnNoRef 
            Caption         =   "Pagos no referenciados"
         End
      End
      Begin VB.Menu mnTarj 
         Caption         =   "&Tarjetas"
         Begin VB.Menu mnsolTarjetas 
            Caption         =   "&Estatus  de Tarjetas"
         End
         Begin VB.Menu mnProdTarjetas 
            Caption         =   "&Producción de Tarjetas"
         End
         Begin VB.Menu mnEntregas 
            Caption         =   "En&trega deTarjetas"
         End
         Begin VB.Menu mnRepMens 
            Caption         =   "Reporte &Mensual de Entregas"
         End
         Begin VB.Menu mnTarCanc 
            Caption         =   "Tarjetas &Canceladas"
         End
         Begin VB.Menu mnSustTar 
            Caption         =   "&Sustitucion de Tarjetas"
         End
         Begin VB.Menu mnuStockRep 
            Caption         =   "Tarjetas en Stock x Cliente"
         End
         Begin VB.Menu mnuTarjxFactura 
            Caption         =   "Tarjetas x Factura"
         End
      End
      Begin VB.Menu mnrepcuentas 
         Caption         =   "C&uentas"
         Begin VB.Menu mnRepAjus 
            Caption         =   "Ajustes a Cuentas"
         End
         Begin VB.Menu mnctasxcte 
            Caption         =   "Cuentas x Cliente"
         End
         Begin VB.Menu mnctasajusfondos 
            Caption         =   "Cuentas x ajustar por Fondos Insuf"
         End
         Begin VB.Menu mnSaldosFinalesver 
            Caption         =   "Consulta Saldos Finales"
         End
      End
      Begin VB.Menu mnInd 
         Caption         =   "&Indicadores"
         Begin VB.Menu mnIndBE 
            Caption         =   "&Indicadores"
         End
         Begin VB.Menu mnDepOB 
            Caption         =   "&Depositos Otros Bancos"
         End
         Begin VB.Menu mnFacTarjetas 
            Caption         =   "&Facturas de Tarjetas"
         End
         Begin VB.Menu mnIndTransacciones 
            Caption         =   "&Indicador de Transacciones"
         End
         Begin VB.Menu mnEdoCuentaGlobal 
            Caption         =   "E&stado de Cuenta Global"
         End
         Begin VB.Menu mnTarClientes 
            Caption         =   "&Tarjetas por Cliente"
         End
      End
      Begin VB.Menu mnRepDia 
         Caption         =   "&Diarios"
         Begin VB.Menu mnLibro 
            Caption         =   "&Libro de Dispersiones"
         End
         Begin VB.Menu mnLibroTarj 
            Caption         =   "L&ibro de Tarjetas"
         End
         Begin VB.Menu mnLiqxCliente 
            Caption         =   "Li&quidaciones por Cliente"
         End
         Begin VB.Menu mnDiario 
            Caption         =   "&Diario de Notas de Crédito"
         End
         Begin VB.Menu mnTransBE 
            Caption         =   "&Transacciones"
         End
         Begin VB.Menu mnConsOI 
            Caption         =   "Consultas de Movimientos &Otros Ingresos"
         End
         Begin VB.Menu mnPedFechaDisp 
            Caption         =   "&Pedidos por Fecha de Dispersion"
         End
         Begin VB.Menu mnRepDiaDispDia 
            Caption         =   "D&ispersión Diaria"
         End
         Begin VB.Menu mnDispvsFact 
            Caption         =   "&Dispersion VS Facturacion"
         End
         Begin VB.Menu mnrepfacdispmens 
            Caption         =   "Dispersion VS &Facturacion Mensual"
         End
         Begin VB.Menu mnConcTran 
            Caption         =   "Conciliacion de &Transacciones"
         End
         Begin VB.Menu mnAjustesMens 
            Caption         =   "&Ajustes Mensual"
         End
         Begin VB.Menu mnResTran 
            Caption         =   "&Resumen de Transacciones"
         End
         Begin VB.Menu mnDetMovMens 
            Caption         =   "D&etalle de Movimientos Mensuales"
         End
         Begin VB.Menu mnuTransNoconc 
            Caption         =   "Transacciones &No conciliadas"
         End
      End
      Begin VB.Menu Estadisticos 
         Caption         =   "Estadisticos"
         Begin VB.Menu mnuIngTarj 
            Caption         =   "&Ingresos de Tarjetas"
         End
         Begin VB.Menu msnuIngAfiliados 
            Caption         =   "Ingresos de &Afiliados"
         End
         Begin VB.Menu mnuEmpActivos 
            Caption         =   "&Empleados Activos x Cliente y Plaza"
         End
      End
   End
   Begin VB.Menu mnUtilerias 
      Caption         =   "&Utilerias"
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnCambioPass 
         Caption         =   "&Cambio de Password"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim vmay As Integer, vmen As Integer, vrev As Integer, sist As String
Dim cierra As Boolean

Private Sub act_envios_Click()
   'ActualizaMenus ("FBE")
   frmactualiza_envios.Show vbModal
End Sub

Private Sub Candado_Click()
  'ActualizaMenus ("FBE")
  frmCandado.Show
End Sub

Private Sub DelphiPlantas_Click()
   'ActualizaMenus ("FBE")
   frmRelacionPlantas.Show
End Sub

Private Sub DelphiSaldos_Click()
   'ActualizaMenus ("FBE")
   frmDispDelphi.Show
End Sub

Private Sub DelphiTarjetas_Click()
  'ActualizaMenus ("FBE")
  frmAsignaDelphi.Show
End Sub

Private Sub Dispersion_media_Click()
  'ActualizaMenus ("FBE")
  frmdisper_media.Show
End Sub

Private Sub FacCtes_Click()
     frmImpFactura.Show
End Sub

Private Sub GenPoliza_Click()
    frmGenPoliza.Show
End Sub

Private Sub FacProv_Click()
   frmFacturasComisiones.Show 1
End Sub

Private Sub frm7Eleven_Click()
  'ActualizaMenus ("FBE")
  frm7Elevena.Show vbModal
End Sub

Private Sub frmLigaCtas_Click()
  'ActualizaMenus ("FBE")
  frmLigar.Show vbModal
End Sub

Private Sub frmPolConsumo_Click()
  'ActualizaMenus ("FBE")
  frmPolizaNConsumo.Show
End Sub

Private Sub frmReimpresion_Click()
  'ActualizaMenus ("FBE")
  frmReimpresionB.Show vbModal
End Sub

Private Sub mnConBitPed_Click()
    frmAutorizaDisp.Show 1
End Sub

Private Sub mnCancOI_Click()
    frmCancFacOI.Show 1
End Sub

Private Sub mnConPed_Click()
    inicia_consulped = 0
    tblConsultaPedidos.Show
End Sub

Private Sub mnGenConsOI_Click()
    sReporte = "OI"
    frmConsultasMov.Show 1
End Sub

Private Sub mnEstab_Click()
    FrmCatEEstablecimientos.Show
End Sub

Private Sub mnGeneraNota_Click()
    frmNotasCargo.Show
End Sub

Private Sub mnImpNotascar_Click()
    frmCancNotasCargo.Show
End Sub

Private Sub mnLibro_Click()
    TipoRep = "FACTURAS"
    frmRepMes.Show
End Sub

Private Sub mnLibroTarj_Click()
    TipoRep = "TARJETAS"
    frmRepMes.Show
End Sub

Private Sub mnLiqxCliente_Click()
    TipoRep = "LIQXCTE"
    frmRepMes.Show
End Sub

Private Sub mnNoRef_Click()
    TipoRep = "NoRef"
    frmRepMes.Show
End Sub

Private Sub mnOI_Click()
    frmOI.Show
End Sub

Private Sub MDIForm_Load()
      
    Set mclsAniform = New clsAnimated
    
       If FileExist("C:\Sistemas\OperacionesVT\Logotipo.gif") Then
         Image1.Picture = LoadPicture("C:\Sistemas\OperacionesVT\Logotipo.gif")
       ElseIf FileExist("C:\Sistemas\OperacionesVT\Logotipo.jpg") Then
         Image1.Picture = LoadPicture("C:\Sistemas\OperacionesVT\Logotipo.jpg")
       ElseIf FileExist("C:\Sistemas\OperacionesVT\Logotipo.bmp") Then
         Image1.Picture = LoadPicture("C:\Sistemas\OperacionesVT\Logotipo.bmp")
       End If
    
    cierra = False
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    gnSucursal = 1
    sbStatusBar.Panels(4).Text = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    sist = "SISTEMA DE OPERACIONES VALE TOTAL"
    vmay = App.Major
    vmen = App.Minor
    vrev = App.Revision
    
End Sub

Private Sub MDIForm_Resize()
Dim i As Integer, j As Integer, iMode As Integer
   
'   If MDIMain.WindowState = vbHide Then Exit Sub
   If Image1.Picture = 0 Then Exit Sub
   
   Picture1.AutoRedraw = True
   Picture1.Cls
   Picture1.Height = Me.ScaleHeight + Picture1.Height - Picture1.ScaleHeight

   iMode = 1
   Select Case iMode
   Case 0 'centered
      Picture1.PaintPicture Image1.Picture, (Me.ScaleWidth - Image1.Width) / 2, (Me.ScaleHeight - Image1.Height) / 2
   Case 1 'stretched
      If Me.ScaleWidth <> 0 And Me.ScaleHeight <> 0 Then
        Picture1.PaintPicture Image1.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight
      End If
   Case 2 'tiled
      For i = 0 To Screen.Height Step Image1.Height
         For j = 0 To Screen.Width Step Image1.Width
            Picture1.PaintPicture Image1, j, i, Image1.Width, Image1.Height
         Next
      Next
   End Select
   
   Me.Picture = Picture1.Image
   Picture1.AutoRedraw = False
End Sub

Private Sub MDIForm_Terminate()
  sql = "SELECT TOP 1 * FROM usuario_ajustes"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rsBD.EOF Then
      If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
         sql = "DELETE usuario_ajustes"
         cnxbdMty.Execute sql
     End If
   End If
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("¿Estas seguro de salir?", vbQuestion + vbYesNo + vbDefaultButton2, "Saliendome...") = vbYes Then
   If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    Call mclsAniform.Animated(Me, eHIDE, 500, AW_BLEND)
    'Set cnxbdMty = New ADODB.Connection
    'cnxbdMty.CommandTimeout = 6000
    'cnxbdMty.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    'cnxbdMty.Close
    'Set cnxbdMty = Nothing
    sql = "DELETE usuario_ajustes"
    cnxbdMty.Execute sql
    'cnxbdMatriz.Close
    'Set cnxbdMatriz = Nothing
    End
Else
   Cancel = 1
End If
End Sub
Private Sub Mn_porFact_Click()
    TipoAcceso = "Factura"
    frmNotas.Show 1
End Sub

Private Sub Mn_PorReem_Click()
    TipoAcceso = "Reembolso"
    frmNotas.Show 1
End Sub

Private Sub mnCancel_Click()
    frmCancNotas.Show
End Sub

Private Sub mnCaptEntrada_Click()
    tblRecBon.Show
End Sub

Private Sub mnDiario_Click()
    TipoRep = "NOTAS"
    frmRepMes.Show
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\" & Reporte & ".rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = cboUENS.ItemData(cboUENS.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = cboBodegas.ItemData(cboBodegas.ListIndex)
    If Reporte = "RembxSuc" Then
        mdiMain.cryReport.StoredProcParam(2) = txtAño1 & Format(cboMes1.ItemData(cboMes1.ListIndex), "00")
    Else
        mdiMain.cryReport.StoredProcParam(2) = txtAño1
        mdiMain.cryReport.StoredProcParam(3) = cboMes1.ItemData(cboMes1.ListIndex)
    End If
    If Reporte = "VenCliMe" Then
        mdiMain.cryReport.Formulas(0) = "Fecha='" & cboMes1.List(cboMes1.ListIndex) & " del " & txtAño1 & "'"
    ElseIf Reporte = "RembxSuc" Then
        mdiMain.cryReport.Formulas(0) = "Fecha='" & cboMes1.List(cboMes1.ListIndex) & " del " & txtAño1 & "'"
        mdiMain.cryReport.StoredProcParam(3) = cboProductos.ItemData(cboProductos.ListIndex)
    ElseIf Reporte = "Comparat" Then
        mdiMain.cryReport.Formulas(0) = "Fecha='De " & cboMes1.List(cboMes1.ListIndex) & " " & txtAño1 & " contra " & cboMes2.List(cboMes2.ListIndex) & " " & txtAño2 & "'"
        mdiMain.cryReport.StoredProcParam(4) = txtAño2
        mdiMain.cryReport.StoredProcParam(5) = cboMes2.ItemData(cboMes2.ListIndex)
    Else
        mdiMain.cryReport.StoredProcParam(4) = txtAño2
        mdiMain.cryReport.StoredProcParam(5) = cboMes2.ItemData(cboMes2.ListIndex)
    End If
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Private Sub mnLecBonos_Click()
    frmBonificaBonos.Show
End Sub

Private Sub mnAltasBanco_Click()
    frmAltaBanco.Show
End Sub

Private Sub mnAsigSaldos_Click()
    frmAsigSaldos.Show
End Sub

Private Sub mnAclara_Click()
    frmAclara.Show
    mskFechaProb = Date
End Sub
Private Sub mnActivaCuentas_Click()
    frmCambiarEstatusTJ.Show
End Sub

Private Sub mnActTarjetas_Click()
    frmActivacion.Show 1
End Sub

Private Sub mnAjustes_Click()
    frmAjustes.Show
End Sub

Private Sub mnAjustesMens_Click()
    TipoRep = "AJUSTESMES"
    frmRepMes.Show vbModal
End Sub

'Private Sub mnAsigPap_Click()
'    frmAsigPapFact.Show
'End Sub

Private Sub mnAutPedidos_Click()
    Dim sqls As String
    sqls = "SELECT * FROM Claves Where Tabla='SistemaFBE'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If rsBD.EOF Then
       MsgBox "El interruptor ha sido cambiado de su posicion original", vbCritical, "Interruptor off-line"
       frmAutorizaDisp.Show 1
    Else
       If rsBD!Status = 1 And gnBodega <> 1 Then
          MsgBox "Lo siento ... la opcion ha sido bloqueada temporalmente debido a que se esta realizando un cierre en este momento", vbCritical, "Opcion bloqueada"
       Else
          frmAutorizaDisp.Show 1
       End If
    End If
End Sub

Private Sub mnBitacora_Click()
    frmAutorizaDisp.Show 1
End Sub

'Private Sub mnBoveda_Click()
'   frmStockPapFact.Show
'End Sub

Private Sub mnCambioPass_Click()
  'ActualizaMenus ("FBE")
  frmCambiaPass.Show vbModal
End Sub

Private Sub mnCancCuentas_Click()
    frmCancCuentas.Show
End Sub

Private Sub mnCancFacBE_Click()
    frmCancelaFactura.Show
End Sub

Private Sub mnCancFact_Click()
    frmCancFacOI.Show
End Sub

Private Sub mnCapManSBI_Click()
    frmCapturaMovManualSBI.Show
End Sub

Private Sub mnCaptFechacorte_Click()
    frmCargosBanco.Show 1
End Sub
Private Sub mnCapturaEnvios_Click()
    frmEnvios.Show
End Sub

Private Sub mnCarBanco_Click()
    TipoRep = "CBCO"
    frmRepMes.Show
End Sub

Private Sub mnCE_Click()
    frmControlEnvios.Show
End Sub

Private Sub mnClientes_Click()
    frmCatClientes.Show
End Sub

Private Sub mnclientesOI_Click()
    frmCatClientesOI.Show
End Sub

Private Sub mnConcTran_Click()
    TipoRep = "CONCTRAN"
    frmRepMes.Show
End Sub

Private Sub mnConsOI_Click()
    frmConsultasMov.Show
End Sub

Private Sub mnConsPed_Click()
    inicia_consulped = 0
    tblConsultaPedidos.Show
End Sub
Private Sub mnConsSaldos_Click()
    frmMovEmpleados.Show
End Sub

Private Sub mnConsTarj_Click()
  'ActualizaMenus ("FBE")
  frmPolxConsumo.Show
End Sub

Private Sub mnconsultaActBancos_Click()
    frmActivBanco.Show 1
End Sub

Private Sub mnCorreoFE_Click()
    frmCorreosFE.Show
End Sub

Private Sub mnCostoTar_Click()
    frmCostoTarjetas.Show
End Sub

Private Sub mnctasajusfondos_Click()
    TipoRep = "AJUSFONDOS"
    frmRepMes.Show
End Sub

Private Sub mnctasxcte_Click()
    TipoRep = "Cuentas"
    frmRepProd.Show
End Sub

Private Sub mnCuentas_Click()
   frmCuentas.Show 1
End Sub

Private Sub mnEntreegas_Click()
End Sub

Private Sub mnDepOB_Click()
    TipoRep = "DEPOB"
    frmRepStatus.Show
End Sub

Private Sub mnDetMovMens_Click()
    TipoRep = "DTM"
    frmRepMes.Show
End Sub

Private Sub mnDispDiaria_Click()
    'frmRepDispersion.Show
    frmSelSucursal.Show 1
End Sub

Private Sub mnDispvsFact_Click()
    TipoRep = "FACVSDISP"
    frmRepMes.Show
End Sub

Private Sub mnEdoCtaG_Click()
  'ActualizaMenus ("FBE")
  frmEdoCtaG.Show
End Sub

Private Sub mnEdoCuentaGlobal_Click()
    TipoRep = "EDOCTAGLOBAL"
    frmRepStatus.Show
End Sub

Private Sub mnEntrega_Click()
    frmEntregaTarjetas.Show
End Sub

Private Sub mnGenConMo_Click()

End Sub

Private Sub mnEntregas_Click()
    TipoRep = "Entregas"
    frmRepProd.Show
End Sub

Private Sub mnEnvioTarjetas_Click()
    frmGuiaxempleado.Show
End Sub

Private Sub mnFacTarjetas_Click()
    TipoRep = "FACTAR"
    frmRepStatus.Show
End Sub

Private Sub mnGenArchivoBanco_Click()
    'ActualizaMenus ("FBE")
    frmSBI.Show 1
End Sub

Private Sub mnGenArchivoBanco2_Click()
  'ActualizaMenus ("FBE")
  SBIPGNew.Show 1
End Sub

Private Sub mnGenArchivoBanco3_Click()
   frmEnviaSyC.Show
'    frmGeneraSyC.Show
End Sub


'Private Sub mnGenArchivos_Click()
' frmReprocesa.Show
'End Sub

Private Sub mnIndBE_Click()
    TipoRep = "KARDEX"
    frmRepStatus.Show
End Sub

Private Sub mnIndTransacciones_Click()
    TipoRep = "INDTRAN"
    frmRepStatus.Show
End Sub

Private Sub mnIng_Click()
    frmPagosCom.Show
End Sub

Private Sub mnInt_Click()
   frmIntereses.Show
End Sub

Private Sub mnIntegraLiqBEPaso_Click()
    frmLiquidacionesBEPaso.Show vbModal
End Sub

Private Sub mnLiqEstabbe_Click()
'    frmLiqEstabBe.Show
    frmValCuentas.Show 1
End Sub

Private Sub mnModifT_Click()
  'ActualizaMenus ("FBE")
  frmModifT.Show
End Sub

Private Sub mnNCOI_Click()
  FrmPolizaNCSYT.Show
End Sub

Private Sub mnNotasCO_Click()
  'ActualizaMenus ("FBE")
   frmNotasCON.Show
End Sub

Private Sub mnNuevas_Click()
   frmTarjetasNva.Show
End Sub

Private Sub mnPedFechaDisp_Click()
    TipoRep = "FACDISP"
    frmRepMes.Show
End Sub

Private Sub mnPlazas_Click()
    frmPlazas.Show
End Sub

Private Sub mnPolComPBE_Click()
    polizaBE = 1
    frmGenPolizaProvBE.Show 1
End Sub

Private Sub mnPolDispConsInt_Click()
   frmDispPolCI.Show
End Sub

Private Sub mnPolOI_Click()
  frmGenPoliza.Show
End Sub

Private Sub mnPResupuesto_Click()
   frmRepPResup.Show 1
End Sub

Private Sub mnProdTarjetas_Click()
    TipoRep = "Produccion"
    frmRepProd.Show
End Sub

Private Sub mnProv_Click()
    frmProveedores.Show
End Sub

Private Sub mnReimpresion_Click()
    frmReimpNotas.Show
End Sub

Private Sub mnPSV_Click()
  'ActualizaMenus ("FBE")
  frmPSV.Show
End Sub

Private Sub mnRepAjus_Click()
    frmRepAjustes.Show
End Sub

Private Sub mnrepfacdispmens_Click()
    TipoRep = "FACVSDISPMES"
    frmRepMes.Show
End Sub

Private Sub mnrepLibroCom_Click()
    TipoRep = "PAPEL"
    frmRepMes.Show
End Sub

Private Sub mnRepLibroComBE_Click()
    TipoRep = "BE"
    frmRepMes.Show
End Sub
Private Sub mnRepMens_Click()
    TipoRep = "TARJETAS"
    frmRepStatus.Show
End Sub
Private Sub mnReposicion_Click()
    frmAsigTarjetas.Show 1
End Sub
Private Sub mnRepSaldos_Click()
    frmRepSaldos.Show
End Sub
Private Sub mnRepTarjetas_Click()
    frmReposicionTarjetas.Show 1
End Sub

Private Sub mnResTran_Click()
    TipoRep = "CXP"
    frmRepMes.Show
End Sub

Private Sub mnSaldosFin_Click()
  'ActualizaMenus ("FBE")
  frmSaldosFin.Show
End Sub

Private Sub mnSaldosFinalesver_Click()
  'ActualizaMenus ("FBE")
  frmConsultaSaldos.Show
End Sub

Private Sub mnsolTarjetas_Click()
    frmRepSolTarjetas.Show
End Sub

Private Sub mnStatus_Click()
    TipoRep = "CLIENTES"
'    frmRepStatus.Show
End Sub

Private Sub mnSubeArch_Click()
    frmLiquidaciones.Show
End Sub

Private Sub mnSustTar_Click()
    TipoRep = "TARSUST"
    frmRepProd.Show
End Sub

Private Sub mnSYT_Click()
    FrmPolizaSYT.Show
End Sub

Private Sub mnTarCanc_Click()
TipoRep = "TARCANC"
frmRepStatus.Show 1
End Sub

Private Sub mnTarClientes_Click()
' ActualizaMenus ("FBE")
' frmTarxCliente.Show
Dim Result As Integer
Screen.MousePointer = 11
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTarjetasXCliente.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 0

    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
    Screen.MousePointer = 1
End Sub

Private Sub mnTransBE_Click()
    frmRepFecha.Show
End Sub

Private Sub mnTranscom_Click()
   TipoRep = "TRANS"
   frmRepTrans.Show
End Sub

Private Sub mnTransDet_Click()
    TipoRep = "TRANSDET"
    frmRepTrans.Show
End Sub

Private Sub mntransSoriana_Click()
    frmTransSori.Show
End Sub

Private Sub mnTransxcomer_Click()
   frmTransXComercio.Show
End Sub

Private Sub mnuAcercaDe_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuBuscarAyuda_Click()
    Dim nRet As Integer
    'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro
    'de diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If ERR Then
            MsgBox ERR.Description
        End If
    End If

End Sub
Private Sub mnuContenido_Click()
    Dim nRet As Integer
    'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
    'puede establecer el archivo de Ayuda para la aplicación en el cuadro
    'de diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If ERR Then
            MsgBox ERR.Description
        End If
    End If

End Sub
Private Sub mnuMosaicoVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuMosaicoHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuCascada_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuActivamasivo_Click()
  'ActualizaMenus ("FBE")
  frmActivaMasivo.Show
End Sub

Private Sub mnuActivaTarj_Click()
   'ActualizaMenus ("FBE")
   frmActiva_Tarjeta.Show
End Sub

Private Sub mnuActSycTarj_Click()
  'ActualizaMenus ("FBE")
  frmActSycSto.Show
End Sub

Private Sub mnuAntig_Click()
  'ActualizaMenus ("FBE")
  frmAntigBE.Show
End Sub

Private Sub mnuAntigBE_Click()
  'ActualizaMenus ("FBE")
  frmAntigBE.Show
End Sub

Private Sub mnuAplicacion_Click()
  tipo_estad = 2
  frmCorteCajaBE.Show
End Sub

Private Sub mnuAsocia_Click()
   frmAsociar.Show
End Sub

Private Sub mnuBloqueo_Click()
   fmrBlockCard.Show
End Sub

Private Sub mnuCaratulas_Click()
  'ActualizaMenus ("FBE")
  frmEnvios_Caratulas.Show
End Sub

Private Sub mnuComerciosSyc_Click()
  'ActualizaMenus ("FBE")
  frmComerciosSyc.Show
End Sub

Private Sub mnuCorteCaja_Click()
 'ActualizaMenus ("FBE")
 tipo_estad = 1
 frmCorteCajaBE.Show
End Sub

Private Sub mnuCreaCtasSyC_Click()
  'ActualizaMenus ("FBE")
  frmCreaCtasSyc.Show
End Sub

Private Sub mnuCuotas_Click()
  'ActualizaMenus ("FBE")
  tipo_estad = 1
  frmEstadisitica.Show
End Sub

Private Sub mnuDepositos_Click()
  'ActualizaMenus ("FBE")
  frmDepositos.Show
End Sub

Private Sub mnuDeshacer_Click()
  'ActualizaMenus ("FBE")
  frmDeshacer.Show
End Sub

Private Sub mnuEdoCtaOI_Click()
  'ActualizaMenus ("FBE")
  frmEdoCtaOI.Show
End Sub

Private Sub mnuEdoCTAoi2_Click()
  'ActualizaMenus ("FBE")
  frmEdoCtaOI.Show
End Sub

Private Sub mnuEmpActivos_Click()
  'ActualizaMenus ("FBE")
  tipo_estad = 3
  frmEstadisitica.Show
End Sub

Private Sub mnuEspPG_Click()
   frmespgas.Show
End Sub

Private Sub mnuFolioSol_Click()
  'ActualizaMenus ("FBE")
  frmfoliosolicitud.Show vbModal
End Sub

Private Sub mnuInformeSYC_Click()
   'ActualizaMenus ("FBE")
   frmSycdetalle.Show
End Sub

Private Sub mnuIng_Click()
   'ActualizaMenus ("FBE")
   frmPagosCom.Show
End Sub

Private Sub mnuIngTarj_Click()
  'ActualizaMenus ("FBE")
  tipo_estad = 0
  frmEstadisitica.Show
End Sub

Private Sub mnuInterruptor_Click()
  'ActualizaMenus ("FBE")
  frmInterruptor_PedBE.Show 1
End Sub

Private Sub mnuLote_Click()
   frmLote.Show
End Sub

Private Sub mnuIntBanco_Click()
    frmReferBanc.Show
End Sub

Private Sub mnuMovsCteEMPL_Click()
    mov_o_stok = 1
    frmRepStock.Show
End Sub

Private Sub mnuOpeCierreMes_Click()
    frmPedidos.Show
End Sub

Private Sub mnuRepDiaLibroVentas_Click()
    Reporte = "LibroVen"
    subReporte = "Libro de Ventas"
    frmRepFechas.Show
End Sub
Private Sub mnuRepDiaNotasConsumo_Click()
    Reporte = "LibroVen"
    subReporte = "Notas de Consumo"
    frmRepFechas.Show
End Sub
Private Sub mnuRepMenCompara_Click()
    Reporte = "Comparat"
    frmRepClientes.Show
End Sub

Private Sub mnuRepMenVenCliAn_Click()
    Reporte = "VenCliAn"
    frmRepClientes.Show
End Sub

Private Sub mnuRepMenVenCliMes_Click()
    Reporte = "VenCliMe"
    frmRepClientes.Show
End Sub

Private Sub mnuRepProFactRuta_Click()
    Reporte = "FactRuta"
    frmRepMes.Show
End Sub

Private Sub mnuRepProSuc_Click()
    Reporte = "ProdxSuc"
    frmRepUENS.Show
End Sub

Private Sub mnuRepRemCte_Click()
    Reporte = "RembxCte"
    frmRepMesClientes.Show
End Sub
Private Sub mnuRepRemMen_Click()
    Reporte = "RembxSuc"
    frmRepClientes.Show
End Sub
Private Sub mnuRepVenCanal_Click()
    Reporte = "VenCanal"
    frmRepFechas.Show
End Sub
Private Sub mnuRepVenxRango_Click()
    Reporte = "Rango"
    frmRepFechas.Show
End Sub

Private Sub mnuOrigenCta_Click()
   'ActualizaMenus ("FBE")
   frmOrigenSolic.Show
End Sub

Private Sub mnuPerfilmaster_Click()
   'ActualizaMenus ("FBE")
  fmrPerfil.Show
End Sub

Private Sub mnuProvBE_Click()
  polizaBE = 2
  frmGenPolizaProvBE.Show 1
End Sub

Private Sub mnuSaldoONline_Click()
   frmSaldoOL.Show
End Sub
Private Sub mnRepDiaDispDia_Click()
    frmSelSucursal.Show 1
End Sub
Private Sub mnuSalir_Click()
   Unload Me
End Sub
Private Sub mnuEspecificarImpresora_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Especificar Impresora"
        .CancelError = True
        .ShowPrinter
    End With
End Sub
Private Sub ReimpOI_Click()
    frmReimpOI.Show
End Sub

Private Sub mnuStatusTarjeta_Click()
   frmStatusTarjeta.Show
End Sub

Private Sub mnuStock_Click()
'  ActualizaMenus ("FBE")
  frmStock.Show
End Sub

Private Sub mnuStockRep_Click()
  ' ActualizaMenus ("FBE")
  frmRepStock.Show
End Sub

Private Sub mnuStockSyC_Click()
  'ActualizaMenus ("FBE")
  frmStockSyC.Show
End Sub

Private Sub mnuTarjxFactura_Click()
  'ActualizaMenus ("FBE")
  frmDetalleFactura.Show
End Sub

Private Sub mnuTransaSyc_Click()
   TipoRep = "CANTRANPROD"
   frmRepMes.Show
End Sub

Private Sub mnuTransNoconc_Click()
  TipoRep = "TRANSASNOC"
  frmRepMes.Show
End Sub

Private Sub mnuUsuarios_Click()
    frmPerfil.Show
End Sub

Private Sub mnuValidacion_Click()
  'ActualizaMenus ("FBE")
  mnuValida.Show
End Sub

Private Sub msnuIngAfiliados_Click()
 ' ActualizaMenus ("FBE")
  tipo_estad = 2
  frmEstadisitica.Show
End Sub

Private Sub PolCOMBE_Click()
 'ActualizaMenus ("FBE")
 polizaBE = 2
 frmGenPolizaProvBE.Show 1
End Sub

Private Sub StatTarjetas_Click()
  TipoRep = "Status de Solicitudes"
  'frmRepStatus.Show
  frmRepMes.Show
End Sub

Private Sub TarjetaEstanca_Click()
   TipoRep = "Estancadas"
   frmRepMes.Show
End Sub
