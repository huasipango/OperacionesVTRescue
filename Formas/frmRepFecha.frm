VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepFecha 
   Caption         =   "Consulta de transacciones"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   615
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   350
         Left            =   120
         Picture         =   "frmRepFecha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   840
         Width           =   350
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   350
         Left            =   120
         Picture         =   "frmRepFecha.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   350
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboProductos 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboTipoMov 
         Height          =   315
         ItemData        =   "frmRepFecha.frx":0274
         Left            =   1320
         List            =   "frmRepFecha.frx":0276
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1440
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
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin: "
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
         TabIndex        =   11
         Top             =   2000
         Width           =   975
      End
      Begin VB.Label Label2 
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
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblAño1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ini: "
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
         Top             =   1540
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   940
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRepFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prod As Byte
Private Sub cmdPresentar_Click()
 Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    If cboTipoMov.ItemData(cboTipoMov.ListIndex) <> 4 Then
       mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
       mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTransaccionesbe.rpt"
       mdiMain.cryReport.Destination = Destino
       mdiMain.cryReport.StoredProcParam(0) = cboTipoMov.ItemData(cboTipoMov.ListIndex)
       mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
       mdiMain.cryReport.StoredProcParam(2) = CStr(cboProductos.ListIndex + 1)
       mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    Else
       mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
       mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptRechazos.rpt"
       mdiMain.cryReport.Destination = Destino
       mdiMain.cryReport.StoredProcParam(0) = cboTipoMov.ItemData(cboTipoMov.ListIndex)
       mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
       mdiMain.cryReport.StoredProcParam(2) = CStr(cboProductos.ListIndex + 1)
       mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    End If
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    
    If Result <> 0 Then
        Call doErrorLog(gnBodega, "OPE", mdiMain.cryReport.LastErrorNumber, mdiMain.cryReport.LastErrorString, Usuario, "frmRepFecha.Imprime")
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
        MsgBar "", False
    End If

 
End Sub
Private Sub Form_Load()
    
    mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
    mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YYYY")
    'mskFechaIni = Date
    sqls = " select * from claves where tabla = 'Liquidacionesbe'" & _
           " and campo = 'status' and nocve  > 0"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    intCount = -1
    cboTipoMov.Clear

    cboProductos.Clear
    Call CargaComboBE(cboProductos, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & Trim(cboProductos.Text) & " ','Leer'")
    cboProductos.Text = UCase("Winko Mart")
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cboTipoMov.AddItem Trim(strBodega)
       cboTipoMov.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    
    rsBD.Close
    Set rsBD = Nothing
        
End Sub


Private Sub cboProductos_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & UCase(Trim(cboProductos.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

