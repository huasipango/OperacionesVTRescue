VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmControlEnvios 
   Caption         =   "Control de Envíos"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   14640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Alignment       =   1  'Right Justify
      Caption         =   "Marcar Todos"
      Height          =   375
      Left            =   12600
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   16
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
         ItemData        =   "frmControlEnvios.frx":0000
         Left            =   1800
         List            =   "frmControlEnvios.frx":000A
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
         TabIndex        =   17
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   14415
      Begin FPSpread.vaSpread spdDetalle 
         Height          =   4455
         Left            =   120
         OleObjectBlob   =   "frmControlEnvios.frx":001C
         TabIndex        =   8
         Top             =   240
         Width           =   14175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12015
      Begin VB.CommandButton ImprimeRuta 
         Height          =   765
         Left            =   8520
         Picture         =   "frmControlEnvios.frx":138F
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   360
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   525
         Left            =   10440
         Picture         =   "frmControlEnvios.frx":49C19
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   600
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.CommandButton cmdIr 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   435
         Left            =   6120
         Picture         =   "frmControlEnvios.frx":49D1B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar"
         Top             =   720
         Width           =   525
      End
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "frmControlEnvios.frx":4A15D
         Left            =   1800
         List            =   "frmControlEnvios.frx":4A16D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cbobodegas 
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
         ItemData        =   "frmControlEnvios.frx":4A199
         Left            =   1800
         List            =   "frmControlEnvios.frx":4A1A0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   11280
         Picture         =   "frmControlEnvios.frx":4A1C2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   600
         Width           =   570
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   9600
         Picture         =   "frmControlEnvios.frx":4A2C4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Grabar"
         Top             =   600
         Width           =   570
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   4680
         TabIndex        =   12
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
         Left            =   4680
         TabIndex        =   13
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
         Left            =   4200
         TabIndex        =   15
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "A:"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   915
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Estatus:"
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
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmControlEnvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte, dame As Byte

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptStatusGuias.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = 0
    mdiMain.cryReport.StoredProcParam(2) = cbostatus.ItemData(cbostatus.ListIndex)
    mdiMain.cryReport.StoredProcParam(3) = 1
    mdiMain.cryReport.StoredProcParam(4) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(5) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(6) = CStr(Product)
  
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores presentados"
    End If
End Sub

Sub Imprime2(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEnvios_Rutas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 1 'cbobodegas.ItemData(cbobodegas.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = 0
    mdiMain.cryReport.StoredProcParam(2) = cbostatus.ItemData(cbostatus.ListIndex)
    mdiMain.cryReport.StoredProcParam(3) = 2
    mdiMain.cryReport.StoredProcParam(4) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(5) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(6) = CStr(Product)
  
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores presentados"
    End If
End Sub

Private Sub cboBodegas_Click()
  'Call cmdIr_Click
End Sub

Private Sub cboStatus_Click()
    If cbostatus.ListIndex >= 0 And cboBodegas.ListIndex >= 0 Then
        cmdIr_Click
    End If
End Sub

Private Sub chkTodos_Click()
Dim i As Integer
If chkTodos.value = 0 Then
   For i = 1 To spddetalle.MaxRows
    spddetalle.Col = 15
    spddetalle.Row = i
    spddetalle.value = 0
   Next
End If
If chkTodos.value = 1 Then
   For i = 1 To spddetalle.MaxRows
     spddetalle.Col = 15
     spddetalle.Row = i
     spddetalle.value = 1
   Next
End If

End Sub

Private Sub cmdGrabar_Click()
Dim Selecc As Integer
On Error GoTo ERR:
With spddetalle
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        CveMens = Left(.Text, 2)
        .Col = 2
        Guia = .Text
        .Col = 3
        Refer = .Text
        .Col = 15
        Selecc = .value
        FechaProc = Date
        .Col = 12
        Distribuidor = .Text
        
        If Selecc = 1 Then
        
            sqls = " exec sp_updControlEnvios @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                   " , @Guia = " & Guia & _
                   " , @CveMensajeria = " & CveMens & _
                   " , @Refer = " & Refer & _
                   " , @Status = " & cbostatus.ItemData(cbostatus.ListIndex) & _
                   " , @FechaProc = '" & Format(FechaProc, "mm/dd/yyyy") & "'" & _
                   " , @Distribuidor = '" & Distribuidor & "'"
                   
            cnxbdMty.Execute sqls, intRegistros
        End If
    Next i
End With
MsgBox "Información Actualizada!!!", vbInformation
InicializaForma
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
End Sub

Private Sub cmdIr_Click()
On Error GoTo ERR:
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = " exec spr_controlenvios @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
           " ,@Guia = 0, @StatusI = " & cbostatus.ItemData(cbostatus.ListIndex) & ",@FechaIni = '" & Format(mskFechaIni, "MM/DD/YYYY") & "', @FechaFin = '" & Format(mskFechaFin, "MM/DD/YYYY") & "'" & _
           " ,@Producto=" & Product
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If rsBD.EOF Then
        MsgBox "No se encontraron envíos en ruta del producto seleccionado."
        Exit Sub
    End If
    
    With spddetalle
        .Col = -1
        .Row = -1
        .Action = 12
        .MaxRows = 0
        i = 0
    
        Do While Not rsBD.EOF
            i = i + 1
            .MaxRows = i
            .Row = i
            .Col = 1
            .Text = Format(rsBD!cvemensajeria, "!@@") & "-" & rsBD!descmens
            .Col = 2
            .Text = Val(rsBD!Guia)
            .Col = 3
            .Text = rsBD!Refer
            .Col = 4
            .TypeComboBoxCurSel = rsBD!tipodoc - 1
            .Col = 5
            .Text = rsBD!cliente
            .Col = 6
            .Text = rsBD!Nombre
            .Col = 7
            .Text = rsBD!Factura
            .Col = 8
            .Text = rsBD!TipoTar
            .Col = 9
            .Text = CDbl(rsBD!total)
            .Col = 10
            .Text = Format(rsBD!FechaEnvio, "mm/dd/yy")
            .Col = 12
            .Text = rsBD!Distribuidor
            .Col = 13
            .Text = Format(rsBD!Fecharesp - 1, "mm/dd/yyyy")
            .Col = 14
            .Text = Format(rsBD!Fecharesp, "mm/dd/yyyy")
            rsBD.MoveNext
        Loop
    End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
End Sub

Private Sub cmdPresentar_Click()
    Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub
Sub InicializaForma()
    mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
    mskFechaFin = Format(Date, "MM/DD/YyYY")
    Call CboPosiciona(cbostatus, 1)
    CargaBodegasServ cboBodegas
    spddetalle.MaxRows = 0
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
     Set mclsAniform = New clsAnimated
     dame = 0
     CboProducto.Clear
     Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
     Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
     CboProducto.Text = UCase("Winko Mart")
        
     mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
     mskFechaFin = Format(Date, "MM/DD/YYYY")
     Call CboPosiciona(cbostatus, 1)
     CargaBodegasServ cboBodegas
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub ImprimeRuta_Click()
  Imprime2 crptToWindow
End Sub
