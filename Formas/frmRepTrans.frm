VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepTrans 
   Caption         =   "Transacciones"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   120
      TabIndex        =   18
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
         ItemData        =   "frmRepTrans.frx":0000
         Left            =   120
         List            =   "frmRepTrans.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Buscar por fecha de "
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   2895
      Begin VB.OptionButton opFecha 
         Caption         =   "Conciliacion"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton opFecha 
         Caption         =   "Transaccion"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2895
      Begin VB.TextBox txtGrupo 
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
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2520
         Picture         =   "frmRepTrans.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2685
      End
      Begin VB.Label Label8 
         Caption         =   "Grupo:"
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
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   735
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         Picture         =   "frmRepTrans.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   450
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   120
         Picture         =   "frmRepTrans.frx":0290
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
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
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
      Begin VB.CheckBox chkRepGpo 
         Caption         =   "Trans. x Grupo"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   720
         TabIndex        =   2
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
         Left            =   720
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   600
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
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   675
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmRepTrans"
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
  lblGrupo.Caption = ""
  chkRepGpo.value = 0
End Sub

Private Sub cmdImprimir_Click()
    Imprime crptToPrinter
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Grupo"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtGrupo = frmConsulta.cliente
       lblGrupo = Trim(frmConsulta.Nombre)
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
    'CentraForma Me
    'CargaBodegas cboBodegas
    Set mclsAniform = New clsAnimated
    
    mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
    mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date) - 1), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YyYY")
    opFecha(0).value = True
    If TipoRep = "TRANSDET" Then
        fraFecha.Visible = True
    Else
        fraFecha.Visible = False
    End If
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
End Sub
Private Sub txtAño_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaNumericos(KeyAscii, txtAño.Text, 0)
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    If TipoRep = "TRANS" Then
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTrans_x_com.rpt" 'NO ES NECESARIO CAMBIARLO
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = Val(txtGrupo)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    Else
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTrans_x_tarjeta.rpt"
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = Val(txtGrupo)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         'mdiMain.cryReport.StoredProcParam(4) = CStr(Product)
         If opFecha(0).value Then mdiMain.cryReport.StoredProcParam(3) = "C"
         If opFecha(1).value Then mdiMain.cryReport.StoredProcParam(3) = "T"
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         mdiMain.cryReport.StoredProcParam(4) = CStr(Product)
    End If
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
    
    If chkRepGpo.value = 1 Then
        MsgBar "Generando Reporte", True
        Limpia_CryReport
        mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
        mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTrans_x_grupo.rpt"
        mdiMain.cryReport.Destination = Destino
        mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "mm/dd/yyyy")
        mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "mm/dd/yyyy")
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        mdiMain.cryReport.StoredProcParam(2) = CStr(Product)
        
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
