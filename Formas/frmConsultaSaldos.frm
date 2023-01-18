VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsultaSaldos 
   Caption         =   "Consulta de Saldos Finales"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos requeridos"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7575
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
         ItemData        =   "frmConsultaSaldos.frx":0000
         Left            =   1320
         List            =   "frmConsultaSaldos.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4095
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
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2520
         Picture         =   "frmConsultaSaldos.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1320
         TabIndex        =   3
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
      Begin VB.Label Label3 
         Caption         =   "Producto:"
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
         TabIndex        =   8
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
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
         TabIndex        =   7
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
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
         TabIndex        =   5
         Top             =   1480
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Generar"
            Object.ToolTipText     =   "Generar Consulta Saldos Finales"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   240
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
            Picture         =   "frmConsultaSaldos.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaSaldos.frx":0748
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
   CboProducto.Clear
   Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
   CboProducto.Text = UCase("Winko Mart")
   mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
       Case "Salir": Unload Me
       Case "Generar"
            Muestra_Saldos
End Select
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Sub InicializaForma()
    txtCliente.Text = ""
    Label1.Caption = ""
    mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date) - 1), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "ClienteBE"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
      txtCliente = frmConsulta.cliente
    Else
      'spPlazas.Enabled = False
    End If
    Label1.Caption = cliente_busca
    
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCliente <> "" Then
   sqls = " EXEC sp_BuscaCliente_Datos '" & txtCliente.Text & "','ClienteBE'"
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
      Label1.Caption = rsBD!b
      mskFechaIni.SetFocus
   End If
End If
End Sub

Sub BUSCA_Cliente()
   sqls = " EXEC sp_BuscaCliente_Datos '" & txtCliente.Text & "','ClienteBE'"
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
      Label1.Caption = rsBD!b
      mskFechaIni.SetFocus
   End If
End Sub

Sub Muestra_Saldos()
    BUSCA_Cliente
    Ejecuta_consulta
End Sub

Sub Ejecuta_consulta()
Dim j As Integer, band As Boolean
On Error GoTo ERR:
band = False

If txtCliente.Text <> "" And txtCliente.Text <> " " And txtCliente.Text <> "0" Then
   Imprime (crptToWindow)
Else
  MsgBox "El numero de Cliente no es valido", vbCritical, "Error en numero de cliente"
  Label1.Caption = ""
  Exit Sub
End If
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Error..."
   Exit Sub
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptSaldosFinales.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Product
    mdiMain.cryReport.StoredProcParam(1) = txtCliente.Text
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub
