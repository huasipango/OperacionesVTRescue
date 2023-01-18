VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepAjustes 
   Caption         =   "Ajustes"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   68
      TabIndex        =   12
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
         ItemData        =   "frmRepAjustes.frx":0000
         Left            =   120
         List            =   "frmRepAjustes.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2760
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   720
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
         Left            =   720
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
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   735
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   120
         Picture         =   "frmRepAjustes.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   120
         Picture         =   "frmRepAjustes.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2520
         Picture         =   "frmRepAjustes.frx":0290
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox txtfolio 
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
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Folio:"
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
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmRepAjustes"
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
  txtFolio.Text = ""
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
    Set mclsAniform = New clsAnimated
    mskFechaIni = "01/" & Mid(Date, 4, 2) & "/" & Format(Date, "yyyy")
    ms = Month(Date) + 1
    ms = IIf(ms > 12, 1, ms)
    If ms > 1 Then
       uf = "01/" & Format(ms, "00") & "/" & Year(Date)
    Else
       uf = "01/" & Format(ms, "00") & "/" & Year(Date) + 1
    End If
    uf = uf - 1
    mskFechaFin = Format(uf, "dd/mm/yyyy")
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase

    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptAjustesBE.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtFolio.Text)
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
