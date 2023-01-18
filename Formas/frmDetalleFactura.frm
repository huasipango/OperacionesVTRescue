VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalleFactura 
   Caption         =   "Detalle de Factura de Tarjetas"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   6015
      Begin VB.TextBox txtserie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtFactura 
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmDetalleFactura.frx":0000
         Left            =   1200
         List            =   "frmDetalleFactura.frx":0007
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblImporte 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   580
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
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
         ItemData        =   "frmDetalleFactura.frx":0029
         Left            =   1800
         List            =   "frmDetalleFactura.frx":0033
         Style           =   2  'Dropdown List
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   210
         Width           =   1545
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1429
      ButtonWidth     =   1058
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalle"
            Key             =   "Detalle"
            Object.ToolTipText     =   "Consulta detalle de factura"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleFactura.frx":0045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleFactura.frx":488DF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDetalleFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte, cliente As Integer

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  CboProducto.Text = UCase("Winko Mart")
  CargaBodegasServ cboBodegas
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Detalle"
         Call Detalle
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Sub Detalle()
Dim sql As String
On Error GoTo ERRO:
  If Val(txtFactura.Text) <> 0 Then
     Imprime crptToWindow
  Else
     MsgBox "Debe de proporcionar el numero de factura", vbExclamation, "Falta Factura"
     txtFactura.SetFocus
     Exit Sub
  End If
Exit Sub
ERRO:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
Dim sql As String
On Error GoTo ERRO:
   If Val(txtFactura.Text) <> 0 And KeyAscii = 13 Then
      sql = "SELECT f.*,c.Nombre FROM FM_facturas f"
      sql = sql & " INNER JOIN Clientes c on c.Cliente=f.Cliente"
      sql = sql & " WHERE  f.Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
      sql = sql & " AND f.Factura=" & Val(txtFactura.Text)
      sql = sql & " AND f.Rubro=12"
      Set rsBD = New ADODB.Recordset
      rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
      
      If Not rsBD.EOF Then
         If rsBD!Status = 2 Then
            MsgBox "Esta factura ya esta cancelada", vbInformation, "Factura cancelada"
            txtFactura = ""
            Exit Sub
         Else
            txtSerie.Text = rsBD!serie
            lblCliente.Caption = rsBD!Nombre
            lblImporte.Caption = rsBD!Subtotal + rsBD!iva
            cliente = rsBD!cliente
         End If
      Else
         MsgBox "No se encontro factura", vbExclamation, "Sin resultados"
         Exit Sub
      End If
   End If
Exit Sub
ERRO:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "sp_SolicitudesBE_varios '','',''," & Product & ",'Reajusta'"
    cnxBD.Execute sqls
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptSolTarjetas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 1
    mdiMain.cryReport.StoredProcParam(1) = cliente
    mdiMain.cryReport.StoredProcParam(2) = "01/01/2011"
    mdiMain.cryReport.StoredProcParam(3) = "01/01/2011"
    mdiMain.cryReport.StoredProcParam(4) = 5
    'prod = IIf(Product = 8, 6, Product)
    'producto_cual
    mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
    mdiMain.cryReport.StoredProcParam(6) = Val(txtFactura.Text)
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

