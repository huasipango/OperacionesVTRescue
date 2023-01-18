VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepDispersion 
   Caption         =   "Dispersion Diaria"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   1215
      Left            =   195
      TabIndex        =   7
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Facturas"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nota Consumo"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1455
      End
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
         ItemData        =   "frmRepDispersion.frx":0000
         Left            =   120
         List            =   "frmRepDispersion.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3120
      End
   End
   Begin VB.Frame fraPeriodo1 
      Height          =   855
      Left            =   255
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1920
         TabIndex        =   5
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
      Begin VB.Label lblAño1 
         Caption         =   "Fecha Dispersion:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   255
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   960
         Picture         =   "frmRepDispersion.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   1920
         Picture         =   "frmRepDispersion.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmRepDispersion"
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
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If Product <> 7 Or cboProducto = UCase("<< TODOS >>") Then
     Option1.Enabled = False
     Option2.Enabled = False
  Else
     Option1.Enabled = True
     Option2.Enabled = True
     Option1.value = True
  End If
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Private Sub cmdPresentar_Click()
    Imprime (crptToWindow)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    mskFechaIni.Text = Date + 1
    cboProducto.Clear
    Call CargaComboBE_All(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
    cboProducto.Text = UCase("<< TODOS >>")
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFactXDispersar.rpt"
    mdiMain.cryReport.connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni.Text, "mm/dd/yyyy")
    'prod = IIf(Product = 8, 6, Product)
    'producto_cual
    mdiMain.cryReport.StoredProcParam(1) = CStr(Product)
    If Option1.value = True And Product = 7 Then
       mdiMain.cryReport.StoredProcParam(2) = "NotasC"
    ElseIf Option2.value = True And Product = 7 Then
       mdiMain.cryReport.StoredProcParam(2) = "0"
    ElseIf Product <> 7 Then
       mdiMain.cryReport.StoredProcParam(2) = "0"
    End If
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaIni.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = Trim(plazasuc)
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
