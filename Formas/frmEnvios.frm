VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEnvios 
   Caption         =   "Captura de Envios"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   6135
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
         ItemData        =   "frmEnvios.frx":0000
         Left            =   1800
         List            =   "frmEnvios.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame frmTarjetas 
      Caption         =   "TARJETAS PENDIENTES POR ENTREGAR"
      Height          =   2175
      Left            =   4080
      TabIndex        =   17
      Top             =   3720
      Width           =   5895
      Begin FPSpread.vaSpread spdtarjetas 
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmEnvios.frx":001C
         TabIndex        =   18
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   11535
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   10920
         Picture         =   "frmEnvios.frx":0337
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Grabar"
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   10920
         Picture         =   "frmEnvios.frx":0439
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Salir"
         Top             =   960
         Width           =   450
      End
      Begin VB.ComboBox cboBodegas 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEnvios.frx":053B
         Left            =   1800
         List            =   "frmEnvios.frx":0542
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtGuia 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboMens 
         Height          =   315
         ItemData        =   "frmEnvios.frx":0564
         Left            =   1800
         List            =   "frmEnvios.frx":0571
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   345
         Left            =   8880
         TabIndex        =   7
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
      Begin MSMask.MaskEdBox mskFechaRec 
         Height          =   345
         Left            =   8880
         TabIndex        =   12
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaReal 
         Height          =   345
         Left            =   8880
         TabIndex        =   14
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Real Entrega:"
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
         Left            =   6960
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Recepcion:"
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
         Left            =   6945
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal:"
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
         Left            =   720
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No Guía:"
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
         Left            =   720
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Envio:"
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
         Left            =   6945
         TabIndex        =   8
         Top             =   405
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Mensajeria:"
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
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DETALLE DE LA GUIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   12855
      Begin FPSpread.vaSpread spdDetalle 
         Height          =   2175
         Left            =   240
         OleObjectBlob   =   "frmEnvios.frx":058E
         TabIndex        =   3
         Top             =   360
         Width           =   12375
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   22
         Top             =   2700
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   2700
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEnvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim FechaSol As Date
Dim Fecharesp As Date
Dim cliente As Long
Dim prod As Byte
Dim ahora As Date
Dim ahora2 As Date, dif As Date


Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptSolTarjetas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 1
    mdiMain.cryReport.StoredProcParam(1) = Val(cliente)
    mdiMain.cryReport.StoredProcParam(2) = Format(Fecharesp, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(Fecharesp, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = 3
    mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
    mdiMain.cryReport.StoredProcParam(6) = 0
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados"
    End If
End Sub
Sub ImprimeGuia(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptGuias.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = txtGuia.Text
    mdiMain.cryReport.StoredProcParam(1) = "0"
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFecha.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFecha.Text, "mm/dd/yyyy")
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
    
End Sub
Function BuscaFactura(cliente As Integer, Factura As Long)
    sqls = " select fecha from clientes_movimientos" & _
           " where cliente = " & cliente & " and refer_apl = " & Factura & _
           " and tipo_mov = 10"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        BuscaFactura = rsBD!Fecha
    Else
        BuscaFactura = 0
    End If
End Function
Function BuscaFolio()

sqls = "select consecutivo + 1  folio from Folios " & _
       " where tipo = 'CE' and Prefijo = 'MSJ'"
       
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    BuscaFolio = rsBD!folio
Else
    BuscaFolio = 1
End If

End Function

Private Sub cboMens_Click()
    If cboMens.Text <> "" Then
        If cboMens.ItemData(cboMens.ListIndex) = 20 Then
            txtGuia.Text = BuscaFolio
        Else
            txtGuia.Text = ""
        End If
    End If
End Sub

Private Sub cmdAgregar_Click()
 With spddetalle
    .MaxRows = .MaxRows + 1
    .Row = .ActiveRow + 1
    .Col = 0
    .SetFocus
    .Action = SS_ACTION_ACTIVE_CEL
 End With
End Sub

Private Sub cmdBorrar_Click()
With spddetalle
    .Row = .ActiveRow
    .Col = 5
    nombre = .Text
    resp = MsgBox("¿Esta seguro de que desea borrar esta fila?", vbYesNo + vbQuestion + vbDefaultButton2, "Quitando registro")
    
    If resp = vbYes And .MaxRows > 0 Then
        .Action = 5
        .MaxRows = .MaxRows - 1
    End If
    frmTarjetas.Visible = False
End With
End Sub

Private Sub cmdGrabar_Click()
Dim Bodega As Integer, CveMens As Integer, Guia As Long, FechaEnvio As Date
Dim accion As String, nacc As Byte
On Error GoTo ERR:

Screen.MousePointer = 11

If cboBodegas.ItemData(cboBodegas.ListIndex) < 0 Then
    MsgBox "Debe seleccionar la sucursal a la cual va el envío", vbInformation, "Falta sucursal"
    Exit Sub
Else
    Bodega = cboBodegas.ItemData(cboBodegas.ListIndex)
End If

If cboMens.ItemData(cboMens.ListIndex) < 0 Then
    MsgBox "Debe seleccionar la mensajeria por la cual se va a enviar ", vbInformation, "Tipo de mensajeria"
    Exit Sub
Else
    CveMens = cboMens.ItemData(cboMens.ListIndex)
End If

If Val(txtGuia.Text) = 0 Then
    MsgBox "Debe capturar el numero de guía", vbInformation, "Falta el No. de Guia"
    Exit Sub
Else
    Guia = txtGuia.Text
End If

With spddetalle
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        tipodoc = .TypeComboBoxCurSel + 1
        .Col = 2
        cliente = Val(.Text)
        .Col = 4
        Factura = .Text
        .Col = 5
        TipoTar = .Text
        .Col = 6
        FechaRespSB = .Text
        .Col = 7
        FechaRecSb = .Text
        .Col = 8
        total = .Text
        .Col = 9
        Contacto = .Text
        .Col = 11
        accion = .Text
        Select Case Trim(accion)
            Case "", "Entregada"
                 nacc = 4
            Case "Devolucion"
                 nacc = 20
            Case "Rechazada"
                 nacc = 21
            Case Else
                nacc = 4
        End Select
        .Col = 10
        Plaza = .Text
        If Plaza = "" Then Plaza = 0
        
        
        If Val(cliente) <> 0 And tipodoc <> 0 Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = " EXEC sp_controlEnvios "
            sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
            sqls = sqls & vbCr & ", @Guia         = " & Guia
            sqls = sqls & vbCr & ", @CveMensajeria= " & CveMens
            sqls = sqls & vbCr & ", @Refer        = " & i
            sqls = sqls & vbCr & ", @FechaEnvio   = '" & Format(mskFecha, "MM/DD/YYYY") & "'"
            sqls = sqls & vbCr & ", @TipoDoc      = " & tipodoc
            sqls = sqls & vbCr & ", @Cliente      = " & Val(cliente)
            sqls = sqls & vbCr & ", @Factura      = " & Factura
            sqls = sqls & vbCr & ", @TipoTar      = '" & TipoTar & "'"
            sqls = sqls & vbCr & ", @FechaRespSB  = '" & FechaRespSB & "'"
            sqls = sqls & vbCr & ", @FechaRecSB   = '" & FechaRecSb & "'"
            sqls = sqls & vbCr & ", @Total        = " & total
            sqls = sqls & vbCr & ", @Contacto     = '" & Trim(Contacto) & "'"
            sqls = sqls & vbCr & ", @status        =" & nacc
            sqls = sqls & vbCr & ", @TipoEnvio    =1"
            sqls = sqls & vbCr & ", @Plaza        = " & Plaza
            sqls = sqls & vbCr & ", @Producto     = " & Product
            cnxbdMty.Execute sqls, intRegistros
            
        End If
     Next i
End With

Screen.MousePointer = 1
ImprimeGuia crptToWindow
LimpiarControles Me
mskFecha.Text = Date
CargaBodegasServ cboBodegas
Call CargaMensajerias(cboMens)
Call CboPosiciona(cboMens, 20)
frmTarjetas.Visible = False
spddetalle.Col = -1
spddetalle.Row = -1
spddetalle.Action = 12
spddetalle.MaxRows = 0
cboBodegas.SetFocus
Exit Sub
ERR:
  MsgBox "Se presento el siguiente error: " & ERR.Description, vbCritical, "Errores presentados"
  Exit Sub
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Sub InicializaForma()
    frmTarjetas.Visible = False
    mskFechaRec.Text = "__/__/____"
    mskFechaReal.Text = "__/__/____"
    'spddetalle.MaxRows = 0
    CargaBodegasServ cboBodegas
    Call CargaMensajerias(cboMens)
    Call CboPosiciona(cboMens, 20)
    frmTarjetas.Visible = False
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
     ahora = Now
     Set mclsAniform = New clsAnimated
     mskFecha.Text = Date
     cboProducto.Clear
     Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
     Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
     cboProducto.Text = UCase("Despensa Total")

     If gnBodega <> 1 Then
        mdiMain.sbStatusBar.Panels(1).Text = "Listo... " & gstrServidor
     End If
     
     CargaBodegasServ cboBodegas
     Call CargaMensajerias(cboMens)
     Call CboPosiciona(cboMens, 20)
     frmTarjetas.Visible = False
     
End Sub

Function CargaSpreadTarjetas(cliente As Integer)
Dim rstar As ADODB.Recordset
 frmTarjetas.Visible = True
 With spdtarjetas
    .Col = -1
    .Row = -1
    .Action = 12
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = " select Isnull(a.FechaEnvio,'01/01/1900')FechaEnvio, a.tipo, isnull(a.plaza,0) Plaza, isnull(b.descripcion, '') Descrip,  count(a.empleado) Cant" & _
          " from solicitudesbe a with (nolock)" & _
          " left outer join plazasbe b with (nolock)" & _
          " on a.cliente = b.cliente" & _
          " and a.plaza =LTRIM(RTRIM(str(b.plaza,8)))" & _
          " Where a.Cliente = " & cliente & _
          " and a.status = 3 and a.tiposol = 1" & _
          " and a.Producto=" & Product & _
          " group by a.FechaEnvio,a.tipo, a.plaza, isnull(b.descripcion, '') " & _
          " order by a.FechaEnvio desc"
    
    Set rstar = New ADODB.Recordset
    rstar.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    .MaxRows = 0
    i = 0
    Do While Not rstar.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Format(rstar!FechaEnvio, "MM/DD/YYYY")
        .Col = 2
        .Text = rstar!tipo
        .Col = 3
        .Text = rstar!Cant
        .Col = 4
        .Text = rstar!Plaza
        rstar.MoveNext
    Loop
    frmTarjetas.Visible = True
    
 End With
End Function

Private Sub Form_Unload(Cancel As Integer)
     If gnBodega <> 1 Then
        mdiMain.sbStatusBar.Panels(1).Text = "Listo... " & gstrServidor
     End If
     Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub


Private Sub spddetalle_KeyPress(KeyAscii As Integer)
On Error GoTo ERRO:
With spddetalle

 If KeyAscii = 13 Then

    Select Case .ActiveCol
        Case 9
            
            Call cmdAgregar_Click
    End Select
 End If
End With
Exit Sub
ERRO:
  MsgBox "Error generado-> " & ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub spdDetalle_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

With spddetalle

 If (Col = 2 And NewCol = 3) Then
    .Row = Row
    .Col = 2
    ClienteS = .Text
    .Col = 1
    tipo = .TypeComboBoxCurSel
    
    sqls = "select Nombre from clientes with (nolock)"
    sqls = sqls & " where cliente = " & Val(ClienteS) & ""
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
            .Row = Row
            .Col = 3
            .Text = rsBD!nombre
            .Col = 4
            .SetFocus
            .Action = SS_ACTION_ACTIVE_CELL
            .Col = 6
            .Text = Format(Date, "MM/DD/YYYY")
    End If
     
    If tipo = 1 Then
        CargaSpreadTarjetas (ClienteS)
    ElseIf tipo = 2 Or tipo = 3 Then
        .Col = 4
        .Text = 0
        .Col = 6
        .Text = Format(Date, "MM/DD/YYYY")
        .Col = 7
        .Text = Format(Date, "mm/dd/yyyy")
        .Col = 8
        .SetFocus
        .Action = SS_ACTION_ACTIVE_CELL
    End If
    
 ElseIf Col = 4 And NewCol = 5 Then
 
    .Col = Col
    .Row = Row
    Factura = .Text
    .Col = 2
    cliente = .Text
    If Val(Factura) <> 0 Then
        sqls = "select refer_apl factura, importe valor, fecha   from clientes_movimientos with (nolock)"
        sqls = sqls & " where tipo_mov  in (10, 12, 12 , 13)  and refer_apl = " & Factura & " and cliente  = " & Val(cliente) & ""
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
           
        If Not rsBD.EOF Then
                .Row = Row
                .Col = 6
                .Text = Format(rsBD!Fecha, "mm/dd/yyyy")
                .Col = 7
                .Text = Format(Date, "mm/dd/yyyy")
                .Col = 8
                .Text = CDbl(rsBD!valor)
                .Col = 9
                .SetFocus
                .Action = SS_ACTION_ACTIVE_CELL
        Else
            MsgBox "La Factura no existe o no esta dada de alta para ese cliente", vbInformation, "Factura incorrecta"
            .Row = Row
            .Col = 4
            .SetFocus
            .Action = SS_ACTION_ACTIVE_CELL
        End If
   
    End If
 
 End If
End With
End Sub

Private Sub spdtarjetas_DblClick(ByVal Col As Long, ByVal Row As Long)
On Error GoTo ERR:
With spdtarjetas
If Col = 0 Then
    If Row = 0 Then
        frmTarjetas.Visible = False
        Exit Sub
    Else
        .Col = 1
        .Row = Row
        Fecharesp = Format(.Text, "mm/dd/yyyy")
        .Col = 2
        TipoTar = Trim(.Text)
        spddetalle.Col = 2
        spddetalle.Row = spddetalle.ActiveRow
        cliente = spddetalle.Text
        Imprime crptToWindow
        Exit Sub
    End If
End If
If .MaxRows < 1 Then
   frmTarjetas.Visible = False
   spddetalle.MaxRows = 0
   Exit Sub
End If
 
    .Row = Row
    .Col = 1
    Fecha = .Text
    .Col = 2
    tipo = .Text
    .Col = 3
    Cant = .Text
    .Col = 4
    Plaza = Trim(.Text)
    
    Select Case Trim(tipo)
        Case "T"
            tipo = 1
        Case "A"
            tipo = 2
        Case "RT"
            tipo = 3
        Case "RA"
            tipo = 4
    End Select
    
    With spddetalle
      .Row = spddetalle.ActiveRow
      .Col = 4
      .Text = 0
      .Col = 6
      .Text = Fecha
      .Col = 5
      .TypeComboBoxCurSel = IIf(IsNull(tipo), 1, tipo)
      .Col = 7
      .Text = Format(Date, "MM/DD/YYYY")
      .Col = 8
      .Text = Cant
       .Col = 10
      .Text = Plaza
      .Col = 9
      .SetFocus
      .Action = SS_ACTION_ACTIVE_CELL
    End With
    frmTarjetas.Visible = False
 End With
Exit Sub
ERR:
  MsgBox "Se encontro el siguiente error: " & ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub txtGuia_KeyPress(KeyAscii As Integer)
Dim TipoTar As Integer
    If KeyAscii = 13 Then
        If Trim(txtGuia.Text) <> "" Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            Guia = txtGuia.Text
'" @Bodega =0 '' " & cboBodegas.ItemData(cboBodegas.ListIndex)
           sqls = " exec spr_ControlEnvios " & _
           "  @Bodega =0" & _
           " ,@Guia = " & Val(txtGuia.Text) & _
           " ,@StatusI = 0" & _
           " ,@TipoEnvioI = 1" & _
           " ,@FechaIni = '" & Format(Date - 30, "MM/DD/YYYY") & "',@FechaFin = '" & Format(Date, "MM/DD/YYYY") & "'" & _
           " ,@Producto =" & CStr(Product)
           
              
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            
            i = 0
            spddetalle.MaxRows = 0
            
            If Not rsBD.EOF Then
                Call CboPosiciona(cboMens, rsBD!cvemensajeria)
                Call CboPosiciona(cboBodegas, rsBD!Bodega)
                mskFecha.Text = rsBD!FechaEnvio
                mskFechaRec.Text = rsBD!FechaRec
                mskFechaReal.Text = rsBD!Fechareal
                lblStatus.Caption = rsBD!Status
                txtGuia.Text = Guia
            End If
            Do While Not rsBD.EOF
                Select Case Trim(rsBD!TipoTar)
                    Case "T"
                        TipoTar = 1
                    Case "A"
                        TipoTar = 2
                    Case "RT"
                        TipoTar = 3
                    Case "RA"
                        TipoTar = 4
                End Select
                With spddetalle
                    i = i + 1
                    .MaxRows = i
                    .Row = i
                    .Col = 1
                    .TypeComboBoxCurSel = Val(rsBD!tipodoc) - 1
                    .Col = 2
                    .Text = rsBD!cliente
                    .Col = 3
                    .Text = rsBD!nombre
                    .Col = 4
                    .Text = rsBD!Factura
                    .Col = 5
                    .Text = TipoTar
                    .Col = 6
                    .Text = Format(rsBD!Fecharesp, "mm/dd/yyyy")
                    .Col = 7
                    .Text = Format(rsBD!FechaRecSb, "mm/dd/yyyy")
                    .Col = 8
                    .Text = CDbl(rsBD!total)
                    .Col = 9
                    .Text = rsBD!Contacto
                rsBD.MoveNext
                End With
            Loop
        End If
    
    End If
End Sub


