VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepProd 
   Caption         =   "Produccion de Tarjetas"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spdPlazas 
      Height          =   1935
      Left            =   128
      OleObjectBlob   =   "frmRepProd.frx":0000
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdBuscaPlaza 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      Picture         =   "frmRepProd.frx":029A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtPlaza 
      Height          =   375
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   19
      Top             =   5280
      Width           =   855
   End
   Begin VB.CheckBox chkCaratula 
      Caption         =   "Generar Caratula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   795
      TabIndex        =   17
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   3255
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRepProd.frx":039C
         Left            =   120
         List            =   "frmRepProd.frx":039E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3255
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
         ItemData        =   "frmRepProd.frx":03A0
         Left            =   120
         List            =   "frmRepProd.frx":03AA
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3000
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   2640
         Picture         =   "frmRepProd.frx":03BC
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
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1575
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
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   2640
      TabIndex        =   11
      Top             =   2640
      Width           =   735
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         Picture         =   "frmRepProd.frx":04BE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   960
         Width           =   450
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   120
         Picture         =   "frmRepProd.frx":0630
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (dd/mm/yyyy) "
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
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
      Begin VB.CheckBox chkPlazas 
         Caption         =   "Separar por Plazas"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   960
         TabIndex        =   3
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
         TabIndex        =   4
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
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   915
         Width           =   255
      End
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Plaza:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   735
      TabIndex        =   21
      Top             =   5400
      Width           =   465
   End
End
Attribute VB_Name = "frmRepProd"
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
     InicializaForma
  End If
End Sub

Sub InicializaForma()
    chkPlazas.value = 0
    lblNombre = ""
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkCaratula_Click()
If chkCaratula.value = 1 Then
   Me.Height = 6360
   CentraForma Me
Else
   Me.Height = 5640
   CentraForma Me
End If
End Sub

Private Sub cmdBuscaPlaza_Click()
Dim i As Integer
If Val(txtCliente) = 0 Then
   Exit Sub
Else
   Me.Height = 8550
   CentraForma Me
End If
    
    spdPlazas.Visible = True
    sqls = "select * from plazasbe"
    If Trim(txtCliente) <> "" Then
        sqls = sqls & " where cliente = " & Val(txtCliente) & ""
    End If
    sqls = sqls & " order by cliente, plaza"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    With spdPlazas
    .Col = -1
    .Row = -1
    .Action = 12
    
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!Plaza
        .Col = 2
        .Text = rsBD!descripcion
        rsBD.MoveNext
    Loop

    rsBD.Close
    Set rsBD = Nothing
    End With
End Sub

Private Sub spdPlazas_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim Plaza As Integer, cliente As Integer

With spdPlazas
    .Row = Row
    .Col = 1
    Plaza = Val(.Text)
    Call BuscaPlaza(txtCliente, Plaza)
End With
spdPlazas.Visible = False
Me.Height = 6360
CentraForma Me
End Sub

Sub BuscaPlaza(cliente As Integer, Plaza As Integer)
    sqls = "select * from plazasbe where cliente = " & cliente & _
           " and plaza = " & Plaza
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        txtPlaza.Text = rsBD!Plaza
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
    If txtCliente.Text <> "" Then
       Imprime crptToWindow
       If TipoRep = "Entregas" Then
          Imprime2 crptToWindow
       End If
    End If
End Sub

Sub Imprime2(Destino)
  Dim Result As Integer, sql As String
    
    If Val(txtCliente) = 0 Or chkCaratula.value = 0 Then
       Exit Sub
    End If
    
    If chkCaratula.value = 1 And Val(txtPlaza) = 0 Then
        Exit Sub
    End If
    
    If chkCaratula.value = 0 And Val(txtPlaza) <> 0 Then
        Exit Sub
    End If
    
    sql = "UPDATE FOLIOS SET Consecutivo=Consecutivo+1"
    sql = sql & " WHERE Tipo='ENV' AND Prefijo='BE'"
    cnxbdMty.Execute sql
     
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase

    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCaratulas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtCliente.Text)
    mdiMain.cryReport.StoredProcParam(1) = Val(Usuario)
    mdiMain.cryReport.StoredProcParam(2) = Val(txtPlaza)
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Sub Imprime(Destino)
Dim stado As String
Dim Result As Integer
    
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    
    If TipoRep = "Entregas" Then
        If chkPlazas.value = 1 Then
            mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEntregas.rpt"
        Else
            mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEntregasSinAgrup.rpt"
        End If
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = 1
        mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(4) = 3
        mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
    ElseIf TipoRep = "Produccion" Then
'        If cbostatus.Text = "SOLICITADA" Then
'           stado = "1"
'        End If
'        If cbostatus.Text = "ACEPTADA" Then
'           stado = "2"
'        End If
'        If cbostatus.Text = "EN RUTA" Then
'           stado = "3"
'        End If
'        If cbostatus.Text = "ENTREGADA" Then
'           stado = "4"
'        End If
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptProdTarjetas.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = 1
        mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(4) = cbostatus.ItemData(cbostatus.ListIndex) 'cbostatus.ListIndex  'stado
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
    ElseIf TipoRep = "TARSUST" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTarjetasSust.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Val(txtCliente)
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    ElseIf TipoRep = "Cuentas" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCtasXCte.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Val(txtCliente)
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(1) = CStr(Product)
    End If
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
    
    If TipoRep = "TARSUST" Then
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTarjetasSustGlobal.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Val(txtCliente)
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
    
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub
Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  
  lblNombre = ""
  mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
  mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YyYY")
  If TipoRep = "Entregas" Then
    chkPlazas.Visible = True
    chkPlazas.value = 1
    Frame3.Visible = False
    Me.Height = 5640
  ElseIf TipoRep = "Entregas" Then
    Me.Caption = "Cuentas x cliente"
    Frame3.Visible = False
  ElseIf TipoRep = "TARSUST" Then
    Me.Caption = "Sustitucion de Tarjetas"
    Me.Height = 5160
    CentraForma Me
    chkPlazas.Visible = False
  Else
    Me.Height = 5160
    CentraForma Me
  End If
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  CboProducto.Text = UCase("Winko Mart")
  Carga_Status_Solicitud cbostatus
  cbostatus.ListIndex = 0
'  cbostatus.Clear
'  cbostatus.AddItem "SOLICITADA"
'  cbostatus.AddItem "ACEPTADA"
'  cbostatus.AddItem "EN RUTA"
'  cbostatus.AddItem "ENTREGADA"
'  cbostatus.Text = UCase("SOLICITADA")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

