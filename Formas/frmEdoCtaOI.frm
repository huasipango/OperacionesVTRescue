VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEdoCtaOI 
   Caption         =   "Estado de cuenta Proveedores BE"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Proporcione las datos para consultar Edo de Cta"
      Height          =   5055
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   5055
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmEdoCtaOI.frx":0000
         Left            =   1200
         List            =   "frmEdoCtaOI.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3615
      End
      Begin VB.CheckBox chk1 
         Caption         =   "Historial Saldado"
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
         Left            =   360
         TabIndex        =   7
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   4695
      End
      Begin VB.TextBox txtFactura 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2400
         Picture         =   "frmEdoCtaOI.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   375
      End
      Begin VB.ComboBox cboBodegas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1200
         TabIndex        =   5
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   1200
         TabIndex        =   6
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta el:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   3915
         Width           =   780
      End
      Begin VB.Label lblAño1 
         AutoSize        =   -1  'True
         Caption         =   "Desde el:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   3345
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2685
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Establec:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   810
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar"
            Key             =   "Mostrar"
            Object.ToolTipText     =   "Muestra estado de cuenta"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdoCtaOI.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdoCtaOI.frx":489B8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdoCtaOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated, iva As Double
Dim prod As Byte
Dim ms As Byte, uf As Date
Dim fe As String, fe2 As String

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  CargaBodegas2 cboBodegas
  'mskFechaIni = "01/" & Mid(Date, 4, 2) & "/" & Format(Date, "yyyy")
  mskFechaIni = "01/01/1990"
    ms = Month(Date) + 1
    If ms > 12 Then
       ms = 1
       uf = "01/" & Format(ms, "00") & "/" & (Mid(Date, 7, 4) + 1)
    Else
       uf = "01/" & Format(ms, "00") & "/" & Mid(Date, 7, 4)
    End If
    uf = uf - 1
    mskFechaFin = Format(uf, "dd/mm/yyyy")
   CboProducto.Clear
   Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
   CboProducto.Text = UCase("Winko Mart")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Mostrar"
         Call Mostrar
End Select
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Establecimientos"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
      txtCliente = frmConsulta.cliente
    Else
      spPlazas.Enabled = False
    End If
    Text1.Text = cliente_busca
    cboBodegas.ListIndex = Bodegp
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCliente <> "" Then
   sqls = " EXEC sp_BuscaCliente_Datos '" & txtCliente.Text & "','Emisores'," & Product
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
   If rsBD.EOF Then
       MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
       txtCliente.SetFocus
       rsBD.Close
       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
   Else
      Text1.Text = UCase(rsBD!b)
      cboBodegas.ListIndex = rsBD!c
      txtFactura.SetFocus
   End If
End If
End Sub

Sub Mostrar()
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEdoCtaOI.rpt" 'NO ES NECESARIO CAMBIARLO
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente.Text)
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = Val(txtFactura.Text)
    mdiMain.cryReport.StoredProcParam(5) = Product
    If chk1.value = 1 Then
       mdiMain.cryReport.StoredProcParam(6) = "S"
    Else
       mdiMain.cryReport.StoredProcParam(6) = "N"
    End If
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

