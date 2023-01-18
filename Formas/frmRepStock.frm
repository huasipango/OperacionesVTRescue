VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepStock 
   Caption         =   "Reporte de Tarjetas de Stock"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   3240
         Picture         =   "frmRepStock.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   2400
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   3240
         Picture         =   "frmRepStock.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   3000
         Width           =   450
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
         Height          =   1455
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   345
            Left            =   960
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   11
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "A:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   915
            Width           =   255
         End
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   2400
         Picture         =   "frmRepStock.frx":0274
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   285
      End
      Begin VB.Frame Frame4 
         Caption         =   "Producto"
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   240
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
            ItemData        =   "frmRepStock.frx":0376
            Left            =   120
            List            =   "frmRepStock.frx":0380
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   3120
         End
      End
      Begin VB.Label lblNombre 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   3375
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
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRepStock"
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
   lblNombre.Caption = ""
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

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  
  lblNombre = ""
  mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
  mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YyYY")
  If mov_o_stok = 1 Then
     Me.Caption = "Movimientos de empleados"
  End If
 
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  CboProducto.Text = UCase("Winko Mart")
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    If mov_o_stok = 1 Then
      mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptMovsCte.rpt"
    Else
      mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptInforme_stock.rpt"
    End If
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Product
    mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
         
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And txtCliente.Text <> "" Then
     BUSCA_Cliente
   End If
End Sub

Sub BUSCA_Cliente()
 sqls = "SELECT Nombre From Clientes Where Cliente=" & Val(txtCliente)
 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
 If rsBD.EOF Then
    MsgBox "El cliente no existe", vbCritical, "Error en cliente"
    lblNombre.Caption = ""
    txtCliente.SetFocus
    Exit Sub
 Else
    lblNombre.Caption = rsBD!Nombre
    Set rsBD = Nothing
    mskFechaIni.SetFocus
 End If
End Sub
