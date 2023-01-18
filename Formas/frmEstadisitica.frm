VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEstadisitica 
   Caption         =   "Estadisticas BE"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   1455
         TabIndex        =   9
         Top             =   2760
         Width           =   2535
         Begin VB.CommandButton cmdPresentar 
            Height          =   450
            Left            =   720
            Picture         =   "frmEstadisitica.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            Tag             =   "Det"
            ToolTipText     =   "Imprime en la Pantalla"
            Top             =   240
            Width           =   450
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Height          =   450
            Left            =   1440
            Picture         =   "frmEstadisitica.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.ComboBox cboBodegas 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2160
         Width           =   3615
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
         ItemData        =   "frmEstadisitica.frx":0274
         Left            =   1200
         List            =   "frmEstadisitica.frx":027E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         TabIndex        =   4
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         TabIndex        =   8
         Top             =   2280
         Width           =   810
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
         TabIndex        =   6
         Top             =   1020
         Width           =   825
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
         TabIndex        =   5
         Top             =   1635
         Width           =   780
      End
      Begin VB.Label Label11 
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
         TabIndex        =   2
         Top             =   420
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmEstadisitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Dim ms As Byte, uf As Date
Dim fe As String, fe2 As String

Private Sub cmdPresentar_Click()
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    If tipo_estad = 0 Then
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptIng_tarjetas.rpt" 'NO ES NECESARIO CAMBIARLO
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = CStr(CboProducto.ListIndex + 1)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(3) = cboBodegas.ItemData(cboBodegas.ListIndex)
    ElseIf tipo_estad = 1 Then
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCuotaInter.rpt" 'NO ES NECESARIO CAMBIARLO
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = CStr(CboProducto.ListIndex + 1)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(3) = cboBodegas.ItemData(cboBodegas.ListIndex)
    ElseIf tipo_estad = 2 Then
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptIng_Afiliados.rpt" 'NO ES NECESARIO CAMBIARLO
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = CStr(CboProducto.ListIndex + 1)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(3) = cboBodegas.ItemData(cboBodegas.ListIndex)
    ElseIf tipo_estad = 3 Then
         mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEmpleadosxPlazaBE.rpt" 'NO ES NECESARIO CAMBIARLO
         mdiMain.cryReport.Destination = Destino
         mdiMain.cryReport.StoredProcParam(0) = CStr(CboProducto.ListIndex + 1)
         mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
         mdiMain.cryReport.StoredProcParam(3) = cboBodegas.ItemData(cboBodegas.ListIndex)
    End If
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
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
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  CboProducto.Text = UCase("Winko Mart")
  CargaBodegas cboBodegas
  mskFechaIni = "01/" & Mid(Date, 4, 2) & "/" & Format(Date, "yyyy")
    ms = Month(Date) + 1
    If ms > 12 Then
       ms = 1
       uf = "01/" & Format(ms, "00") & "/" & (Mid(Date, 7, 4) + 1)
    Else
       uf = "01/" & Format(ms, "00") & "/" & Mid(Date, 7, 4)
    End If
    uf = uf - 1
    mskFechaFin = Format(uf, "dd/mm/yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
