VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGuiaxempleado 
   Caption         =   "Captura de Guias x Empleado"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmGuiaxEmpleado.frx":0000
   ScaleHeight     =   8445
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spddetalle 
      Height          =   5415
      Left            =   240
      OleObjectBlob   =   "frmGuiaxEmpleado.frx":1446
      TabIndex        =   6
      Top             =   2640
      Width           =   14655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de envio"
      Height          =   735
      Left            =   7560
      TabIndex        =   15
      Top             =   1560
      Width           =   7335
      Begin VB.ComboBox cboMens 
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
         ItemData        =   "frmGuiaxEmpleado.frx":2796
         Left            =   1200
         List            =   "frmGuiaxEmpleado.frx":27A0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2760
      End
      Begin VB.TextBox txtGuia 
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
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Mensarjeria:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   4200
         TabIndex        =   18
         Top             =   315
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pegar datos desde excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12360
      Picture         =   "frmGuiaxEmpleado.frx":27B2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Archivo"
            Object.ToolTipText     =   "Importa archivo csv"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Envia"
            Object.ToolTipText     =   "Enviar a ruta"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Otro"
            Object.ToolTipText     =   "Otro Envio"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cargar"
            Object.ToolTipText     =   "Cargar datos del sistema"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Marcar Todos"
      Height          =   375
      Left            =   11040
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos de entrada"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   14535
      Begin VB.ComboBox cboPlazas 
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
         ItemData        =   "frmGuiaxEmpleado.frx":4B03C
         Left            =   7440
         List            =   "frmGuiaxEmpleado.frx":4B046
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   3480
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   345
         Left            =   6000
         Picture         =   "frmGuiaxEmpleado.frx":4B058
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
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
         Left            =   4920
         TabIndex        =   1
         Top             =   240
         Width           =   975
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
         ItemData        =   "frmGuiaxEmpleado.frx":4B15A
         Left            =   1200
         List            =   "frmGuiaxEmpleado.frx":4B164
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2760
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   11400
         TabIndex        =   3
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
         Left            =   13200
         TabIndex        =   4
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
         Height          =   195
         Left            =   6720
         TabIndex        =   21
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "A:"
         Height          =   255
         Left            =   12840
         TabIndex        =   12
         Top             =   315
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   11040
         TabIndex        =   11
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   4200
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   320
         Width           =   690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuiaxEmpleado.frx":4B176
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuiaxEmpleado.frx":4B7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuiaxEmpleado.frx":4BABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuiaxEmpleado.frx":4BF0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGuiaxEmpleado.frx":4C35E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmnAbrir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   7215
      Begin VB.Label lblCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Label lblcte 
      AutoSize        =   -1  'True
      Caption         =   "ojoa"
      ForeColor       =   &H00404080&
      Height          =   555
      Left            =   600
      TabIndex        =   14
      Top             =   1920
      Width           =   6780
   End
End
Attribute VB_Name = "frmGuiaxempleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte, sistk As Byte
'---variables para poner en el grid
Dim emplead As String, nombr As String, plaz As String, Guia As String, fech As Date, tipin As String
Dim fsol As Date, fresp As Date, ctte As Integer
Dim enbusca As String, Empleadora As String
Dim feche As Date, bandera As Boolean
Dim Plaza As Integer
Dim cliente As Integer

Private Sub Check1_Click()
Dim i As Integer
If Check1.value = 0 Then
   For i = 1 To spddetalle.MaxRows
    spddetalle.Col = 9
    spddetalle.Row = i
    spddetalle.value = 0
   Next
End If
If Check1.value = 1 Then
   For i = 1 To spddetalle.MaxRows
     spddetalle.Col = 9
     spddetalle.Row = i
     spddetalle.value = 1
   Next
End If
End Sub

Private Sub Command1_Click()
Dim a As Variant, b As Variant
Dim aux As String
Dim Cte As Long, emp As String, FOL As String, Fenv As String, GUIA2 As String
On Error GoTo ERR:
bandera = True
Me.Caption = "Espere mientras se sube la informacion..."
With spddetalle
   a = Split(Clipboard.GetText, Chr(13))
   .MaxRows = UBound(a)
   .Action = 24
   For i = 1 To .MaxRows
       b = Split(a(i - 1), vbTab)
       If UBound(b) = 2 Then
          Cte = Val(b(0))
          FOL = Trim(b(1))
          emp = Trim(b(2))
          GUIA2 = "SIN GUIA"
          Fenv = Format(Date, "dd/mm/yyyy")
       ElseIf UBound(b) = 3 Then
          Cte = Val(b(0))
          FOL = Trim(b(1))
          emp = Trim(b(2))
          GUIA2 = Trim(b(3))
          Fenv = Format(Date, "dd/mm/yyyy")
       ElseIf UBound(b) = 4 Then
          Cte = Val(b(0))
          FOL = Trim(b(1))
          emp = Trim(b(2))
          GUIA2 = Trim(b(3))
          Fenv = Trim(b(4))
       ElseIf UBound(b) < 2 Or UBound(b) > 4 Then
           MsgBox "Con los parametros proporcionados no puedo realizar busquedas", vbCritical, "Parametros incompletos"
           .MaxRows = 0
           Exit Sub
       End If
       Call BuscaEmpleadox(Cte, emp)
       If enbusca = "" Then
               .Row = i
               '.MaxRows = i
               .Col = 1
               .ForeColor = vbBlack
               .Text = ctte
               .Col = 2
               .ForeColor = vbBlack
               .Text = Empleadora
               .Col = 3
               .ForeColor = vbBlack
               .Text = Val(FOL)
               .Col = 4
               .ForeColor = vbBlack
               .Text = emplead
               .Col = 5
               .ForeColor = vbBlack
               .Text = nombr
               .Col = 6
               .ForeColor = vbBlack
               .Text = plaz
               .Col = 7
               .ForeColor = vbBlack
               .Text = GUIA2
               .Col = 8
               .ForeColor = vbBlack
               If Fenv = "" Or Len(Fenv) = 0 Then
                  .Text = Format(Date, "dd/mm/yyyy")
               Else
                  .Text = Format(Fenv, "dd/mm/yyyy")
               End If
               '.Text = Trim(ArrVal(4))
               .Col = 10
               .ForeColor = vbBlack
               .Text = tipin
               .Col = 11
               .ForeColor = vbBlack
               .Text = fsol
               .Col = 12
               .ForeColor = vbBlack
               .Text = fresp
               .Col = 13
               .ForeColor = vbBlack
               .Text = IIf(sistk = 1, "S", "N")
       Else
               .Row = i
               '.MaxRows = i
               .Col = 1
               .ForeColor = vbRed
               .Text = Cte 'enbusca
               .Col = 2
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 3
               .ForeColor = vbRed
               .Text = Val(FOL)
               .Col = 4
               .ForeColor = vbRed
               .Text = Trim(enbusca)
               .Col = 5
               .ForeColor = vbRed
               .Text = Trim(emp)
               .Col = 6
               '.Text = Val(plaz)
               .Col = 7
               '.Text = Val(Guia)
               .Col = 8
               '.Text = fech
               .Col = 10
               .ForeColor = vbRed
               .Text = Trim(enbusca)
               .Col = 11
               .ForeColor = vbRed
               .Text = Trim(enbusca)
               .Col = 12
               .ForeColor = vbRed
               .Text = Trim(enbusca)
               .Col = 7
               .Action = 0
       End If
   Next
End With
Me.Caption = "Captura de Guias x Empleado"
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  'spddetalle.MaxRows = 0
  Exit Sub
End Sub

Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
  CboProducto.Clear
  Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
  Call CargaMensajerias(cboMens)
  cboMens.ListIndex = 0
  CboProducto.Text = UCase("Winko Mart")
  mskFechaIni = Format(Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Format(IIf(Month(Date) = 1, Year(Date) - 1, Year(Date)), "0000"), "MM/DD/YYYY")
  mskFechaFin = Format(FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), IIf(Month(Date) = 1, Year(Date) - 1, Year(Date))), "MM/DD/YyYY")
  bandera = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
'  If aqui <> Product Then
'     InicializaForma
'  End If
End Sub

Sub InicializaForma()
    Dim i As Integer
    spddetalle.MaxRows = 1
    For i = 1 To 13
        spddetalle.Row = 1
        spddetalle.Col = i
        spddetalle.Text = ""
        spddetalle.ForeColor = vbBlack
    Next
    spddetalle.Col = 9
    spddetalle.value = 0
    spddetalle.Col = 1
    spddetalle.Action = 0
    txtCliente.Text = ""
    lblCliente.Caption = ""
End Sub
Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
 bandera = False
End Sub


Private Sub spdDetalle_Change(ByVal Col As Long, ByVal Row As Long)
Dim cliente As Long, empleado As String, tipo As String
Dim Empleadora As String
If bandera = False Then
With spddetalle
    If Col = 1 Then
         .Col = Col
         .Row = Row
         cliente = Val(.Text)
         Nombre = BuscaCliente(cliente, "")
         .Col = 2
         .Text = Trim(Nombre)
         If Trim(Nombre) = "No existe cliente" Then
            .Col = 0
         Else
            .Col = 4
         End If
         .Action = 0
    End If
    .Col = 1
    cliente = Val(.Text)
    .Col = 2
    Empleadora = Trim(.Text)
    If (Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Or Col = 7 Or Col = 8) And (cliente = 0 Or Empleadora = "" Or Empleadora = "No existe cliente") Then
       .Col = Col
       .Row = Row
       .Text = ""
        MsgBox "Capture primero el numero de cliente", vbCritical, "Cliente pendiente"
        Col = 1
       .Col = 1
       .Row = Row
       .Action = 0
       Exit Sub
    End If
    If (Col = 4) And (cliente <> 0 And Empleadora <> "" And Empleadora <> "No existe cliente") Then
       .Col = 4
       empleado = .Text
       Call BuscaEmpleadoaqui(cliente, empleado)
       If enbusca = "" Then
         .Row = Row
         .ForeColor = vbBlack
         .Col = 4
         .Text = emplead
         .ForeColor = vbBlack
         .Col = 5
         .Text = Trim(nombr)
         .Col = 6
         .Text = Val(plaz)
         .Col = 7
         .Text = Guia
         .Col = 8
         .Text = fech
         .Col = 10
         .ForeColor = vbBlack
         .Text = tipin
         .Col = 11
         .ForeColor = vbBlack
         .Text = fsol
         .Col = 12
         .ForeColor = vbBlack
         .Text = fresp
         .Col = 13
         .ForeColor = vbBlack
         .Text = sistk
         .Col = 7
         .Action = 0
         Exit Sub
       Else
         .ForeColor = vbRed
         .Row = Row
         .Col = 4
         .Text = enbusca
         .Col = 4
         .ForeColor = vbRed
         .Text = Trim(enbusca)
         .Col = 6
         .Text = Val(plaz)
         .Col = 7
         .Text = Val(Guia)
         .Col = 8
         .Text = fech
         .Col = 10
         .ForeColor = vbRed
         .Text = Trim(enbusca)
         .Col = 11
         .ForeColor = vbRed
         .Text = Trim(enbusca)
         .Col = 12
         .ForeColor = vbRed
         .Text = Trim(enbusca)
         .Col = 7
         .Action = 0
         Exit Sub
       End If
    End If
End With
End If
End Sub

Public Sub BuscaEmpleadoaqui(cliente As Long, empleado As String)
        
   sqls = "sp_EnvioTarjetasBE " & Product & "," & cliente & ",'" & Trim(empleado) & "','"
   sqls = sqls & Format(mskFechaIni, "mm/dd/yyyy") & "','" & Format(mskFechaFin, "mm/dd/yyyy") & "',Null,'BuscaEmpleado'"
   Set rsNombre = New ADODB.Recordset
   rsNombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
   
'   sqls = " select s.Cliente,c.Nombre NombreCliente,s.Tipo,s.Empleado,Isnull(s.Plaza,1)Plaza,Isnull(s.NoGuia,'SIN GUIA')NoGuia,Isnull(s.Stock,0)Stock,"
'   sqls = sqls & " Isnull(Substring(Rtrim(Ltrim(s.Nombre))+ ' ' + Rtrim(Ltrim(s.Apat))+ ' ' +Rtrim(Ltrim(s.AMat)),1,26),'SIN NOMBRE') Nombre,"
'   sqls = sqls & " s.Tipo,convert(varchar,s.FechaSol,101) FechaSol,convert(varchar,s.FechaResp,101)FechaResp"
'   sqls = sqls & " from Solicitudesbe s with (Nolock)"
'   sqls = sqls & " Inner Join Clientes c with (Nolock) On c.Cliente=s.Cliente"
'   sqls = sqls & " where s.cliente = " & cliente
'   sqls = sqls & " and s.Empleado='" & empleado & "'"
'   sqls = sqls & " and s.Producto=" & Product
'    sqls = sqls & " AND s.Fecharesp>='" & Format(mskFechaIni, "mm/dd/yyyy") & "' AND s.Fecharesp<='" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'"
'   sqls = sqls & " and s.Status=2" 'status de Aceptadas, listas para enviar
'   Set rsnombre = New ADODB.Recordset
'   rsnombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
   If Not rsNombre.EOF Then
      nombr = rsNombre!Nombre
      emplead = rsNombre!empleado
      plaz = rsNombre!Plaza
      Guia = rsNombre!NoGuia
      fech = Format(Date, "dd/mm/yyyy")
      tipin = rsNombre!tipo
      fsol = Format(rsNombre!FechaSol, "mm/dd/yyyy")
      fresp = Format(rsNombre!Fecharesp, "mm/dd/yyyy")
      sistk = rsNombre!Stock
      enbusca = ""
   Else
      enbusca = "No existe"
   End If
     
'   rsnombre.Close
'   Set rsnombre = Nothing
   Exit Sub
End Sub

Private Sub spddetalle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And spddetalle.ActiveCol = 7 Then
     spddetalle.Col = 8
     spddetalle.Action = 0
     Exit Sub
  End If
  If KeyAscii = 13 And spddetalle.ActiveCol = 8 Then
     spddetalle.Col = 9
     spddetalle.Action = 0
     Exit Sub
  End If
  If KeyAscii = 13 And spddetalle.ActiveCol = 9 Then
       spddetalle.MaxRows = spddetalle.MaxRows + 1
       spddetalle.Row = spddetalle.MaxRows
       spddetalle.Action = 12
       spddetalle.Col = 1
       spddetalle.Action = 0
       spddetalle.SetFocus
  End If
End Sub

Private Sub spddetalle_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 And spddetalle.MaxRows > 1 Then
     spddetalle.Row = spddetalle.ActiveRow
     spddetalle.Action = 5
     spddetalle.MaxRows = spddetalle.MaxRows - 1
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Envia"
         ENVIAR
   Case "Archivo"
         CSV
   Case "Otro"
         denuevo
   Case "Cargar"
         Carga
   End Select
End Sub

Sub Carga()
Dim i As Long, emp As String
Dim hubo As Long
On Error GoTo ERR:
    spddetalle.MaxRows = 0
If txtCliente.Text <> "" Then
   sqls = "sp_EnvioTarjetasBE " & Product & "," & Val(txtCliente.Text) & "," & cboPlazas.ItemData(cboPlazas.ListIndex) & ",'"
   sqls = sqls & Format(mskFechaIni, "mm/dd/yyyy") & "','" & Format(mskFechaFin, "mm/dd/yyyy") & "',Null,'Carga'"
   Set rsNombre = New ADODB.Recordset
   rsNombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
'   sqls = "SELECT s.Cliente,s.Empleado,Isnull(s.Plaza,1)Plaza,'SIN GUIA' guiax,'" & Format(Date, "mm/dd/yyyy") & "' Fecha_enviox,c.Nombre NomCliente,Isnull(S.Folio,0)Folio,Isnull(S.Stock,0)Stock,"
'   sqls = sqls & " Substring(Rtrim(Ltrim(Isnull(s.Nombre,'')))+ ' ' + Rtrim(Ltrim(Isnull(s.Apat,'')))+ ' ' +Rtrim(Ltrim(Isnull(s.AMat,''))),1,26) NombreCom,"
'   sqls = sqls & " s.Tipo,convert(varchar,s.FechaSol,101) FechaSol,convert(varchar,s.FechaResp,101)FechaResp"
'   sqls = sqls & " FROM Solicitudesbe s with (Nolock)"
'   sqls = sqls & " Inner Join Clientes c with (Nolock) On c.Cliente=s.Cliente"
'   sqls = sqls & " WHERE s.Cliente=" & Val(txtCliente.Text)
'   sqls = sqls & " AND Producto=" & Product
'   sqls = sqls & " AND s.Fecharesp>='" & Format(mskFechaIni, "mm/dd/yyyy") & "' AND s.Fecharesp<='" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'"
'   sqls = sqls & " AND s.STATUS=2 ORDER BY s.Folio,s.Empleado"
'   Set rsnombre = New ADODB.Recordset
'   rsnombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
      
   If Not rsNombre.EOF Then
      i = 1
      Do While Not rsNombre.EOF
      With spddetalle
               .MaxRows = i
               .Row = i
               .Col = 1
               .Text = rsNombre!cliente
               .Col = 2
               .Text = rsNombre!NomCliente
               .Col = 3
               .Text = rsNombre!Folio
               .Col = 4
               .Text = rsNombre!empleado
               .Col = 5
               .Text = rsNombre!Nombrecom
               .Col = 6
               .Text = rsNombre!Plaza
               .Col = 7
               .Text = rsNombre!guiax
               .Col = 8
               .Text = rsNombre!Fecha_enviox
               .Col = 10
               .Text = rsNombre!tipo
               .Col = 11
               .Text = rsNombre!FechaSol
               .Col = 12
               .Text = rsNombre!Fecharesp
               .Col = 13
               .Text = IIf(rsNombre!Stock = 1, "S", "N")
               i = i + 1
               rsNombre.MoveNext
      End With
      Loop
   Else
     MsgBox "No hay registros por enviar de este cliente", vbInformation, "Sin informacion"
   End If
'   rsnombre.Close
'   Set rsnombre = Nothing
   Exit Sub
Else
   MsgBox "No ha seleccionado ningun cliente", vbExclamation, "Error en cliente"
   txtCliente.SetFocus
End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
End Sub

Sub ENVIAR()
Dim i As Long, emp As String
Dim hubo As Long, maxim As Integer
Dim gui As String, tip As String, foc As Integer, dst As String
On Error GoTo ERR:
    If Val(txtGuia) <= 0 Then
        MsgBox "Favor de capturar el número de guia"
        txtGuia.SetFocus
        Exit Sub
    End If
   If MsgBox("Esta a punto de mandar estas tarjetas a mensajeria." & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "A punto de enviar tarjetas") = vbYes Then
      With spddetalle
        hubo = 0
        maxim = .MaxRows
        For i = 1 To .MaxRows
            .Row = i
            .Col = 4
            emp = Trim(.Text)
            .Col = 9
            If .value = 1 And emp <> "No existe" Then
               hubo = hubo + 1
               .Col = 1
               cliente = Val(.Text)
               .Col = 3
               foc = Val(.Text)
               .Col = 6
               Plaza = Val(.Text)
               .Col = 7
               gui = UCase(Trim(.Text))
               .Col = 8
               feche = Format(.Text, "mm/dd/yyyy")
               .Col = 10
               tip = UCase(Trim(.Text))
               .Col = 13
               dst = UCase(Trim(.Text))
 '              If dst = "N" Then
                  sqls = "sp_EnvioTarjetasBE " & Product & "," & cliente & ",'" & Trim(emp) & "','"
                  sqls = sqls & Format(feche, "mm/dd/yyyy") & "','" & Format(feche, "mm/dd/yyyy") & "','" & tip & "','ActualizaN',"
                  sqls = sqls & " '" & txtGuia & "'," & foc
                  cnxbdMty.Execute sqls
                  
'                  sqls = " Update Solicitudesbe"
'                  sqls = sqls & " SET Status=3,FechaEnvio='" & Format(feche, "mm/dd/yyyy") & "',NoGuia='" & gui & "',Folio=" & foc
'                  sqls = sqls & " from Solicitudesbe"
'                  sqls = sqls & " where cliente = " & cliente
'                  sqls = sqls & " and Empleado='" & emp & "'"
'                  sqls = sqls & " and Tipo='" & tip & "'"
'                  sqls = sqls & " and Producto=" & Product
'                  sqls = sqls & " and Status=2"
'                  cnxbdMty.Execute sqls
  '             ElseIf dst = "S" And foc > 0 Then
  '               sqls = "sp_EnvioTarjetasBE " & Product & "," & cliente & ",'" & Trim(emp) & "','"
  '              sqls = sqls & Format(feche, "mm/dd/yyyy") & "','" & Format(feche, "mm/dd/yyyy") & "','" & tip & "','ActualizaN',"
  '                sqls = sqls & " '" & Trim(gui) & "'," & foc
   '               cnxbdMty.Execute sqls

'                  sqls = " Update Solicitudesbe"
'                  sqls = sqls & " SET Status=3,FechaEnvio='" & Format(feche, "mm/dd/yyyy") & "',NoGuia='" & gui & "',Folio=" & foc
'                  sqls = sqls & " from Solicitudesbe"
'                  sqls = sqls & " where cliente = " & cliente
'                  sqls = sqls & " and Empleado='" & emp & "'"
'                  sqls = sqls & " and Tipo='" & tip & "'"
'                  sqls = sqls & " and Producto=" & Product
'                  sqls = sqls & " and Status=2"
'                  cnxbdMty.Execute sqls
      '         ElseIf dst = "S" And foc <= 0 Then
      '            .Col = 3
      '            .Text = "¿Folio?"
      '         End If
            End If
        Next
        If hubo > 0 Then
           crea_guia
           MsgBox "Se han enviado a ruta " & hubo & " de las solicitudes proporcionadas", vbInformation, "En ruta..."
           If txtCliente.Text <> "" Then
              Imprime crptToWindow
              Imprime2 crptToWindow
           End If
           InicializaForma
           Check1.value = 0
        Else
           MsgBox "No hay nada seleccionado que enviar o presenta errores la informacion presentada", vbExclamation, "Informacion incompleta"
        End If
      End With
   End If
Exit Sub
ERR:
 MsgBox ERR.Description, vbCritical, "Error"
End Sub

Sub CSV()
Dim nArchivo, clinea As String, i As Long
Dim valor As Double
Dim TiempoIni As Date, TiempoFin As Date
Dim Cuenta As Long, empleado As String, client As Long
Dim ImpVentas, impdev, ImpTotal As Double
Dim TotInter, TotDeb, TotPos As Double
Dim NoInter, NoDeb, NoPos, tipocom As Integer
Dim Nombre As String
Dim ArrVal() As String
Dim ArrImp() As Double
Dim ArrCant() As Integer
Screen.MousePointer = 11
NumEmp = 0
With spddetalle
    .Col = -1
    .Row = -1
    .Action = 12
    cmnAbrir.ShowOpen
    Screen.MousePointer = 1
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        i = 0
        
        On Error GoTo err_file
        Do While Not EOF(nArchivo)
            Line Input #nArchivo, clinea
            ArrVal = Split(clinea, ",")
            
            client = CLng(ArrVal(0))
            empleado = ArrVal(1)
            If client = 0 Or empleado = "" Then
               MsgBox "Error en algun registro...revise su archivo y vuelva a intentarlo", vbCritical, "Error en archivo"
               InicializaForma
               Exit Sub
            End If
            Call BuscaEmpleadox(client, empleado)
            i = i + 1
            If enbusca = "" Then
               .Row = i
               .MaxRows = i
               .Col = 1
               .ForeColor = vbBlack
               .Text = client
               .Col = 2
               .ForeColor = vbBlack
               .Text = Empleadora
               .Col = 4
               .ForeColor = vbBlack
               .Text = emplead
               .Col = 5
               .ForeColor = vbBlack
               .Text = nombr
               .Col = 6
               .ForeColor = vbBlack
               .Text = plaz
               .Col = 7
               .ForeColor = vbBlack
               .Text = Trim(ArrVal(3))
               .Col = 8
               .ForeColor = vbBlack
               If Trim(ArrVal(4)) = "" Or Len(ArrVal(4)) = 0 Then
                   .Text = Format(Date, "mm/dd/yyyy")
               Else
                  .Text = Trim(ArrVal(4))
               End If
               .Col = 10
               .ForeColor = vbBlack
               .Text = tipin
               .Col = 11
               .ForeColor = vbBlack
               .Text = fsol
               .Col = 12
               .ForeColor = vbBlack
               .Text = fresp
               .Col = 13
               .ForeColor = vbBlack
               .Text = IIf(sistk = 1, "S", "N")
             Else
               .Row = i
               .MaxRows = i
               .Col = 1
               .ForeColor = vbRed
               .Text = client
               .Col = 2
               .ForeColor = vbRed
               .Text = Empleadora
               .Col = 4
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 5
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 6
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 7
               .ForeColor = vbRed
               .Text = Trim(ArrVal(3))
               .Col = 8
               .ForeColor = vbRed
               .Text = Trim(ArrVal(4))
               .Col = 10
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 11
               .ForeColor = vbRed
               .Text = enbusca
               .Col = 12
               .ForeColor = vbRed
               .Text = enbusca
             End If
        Loop
        Close #nArchivo
        Screen.MousePointer = 1
        Exit Sub
    Else
        Exit Sub
    End If
End With
Screen.MousePointer = 1
Exit Sub
err_file:
    MsgBox "Error en el formato del archivo", vbCritical
    Screen.MousePointer = 1
End Sub

Public Sub BuscaEmpleadox(cliente As Long, empleado As String)
   sqls = "sp_EnvioTarjetasBE " & Product & "," & cliente & ",'" & Trim(empleado) & "','"
   sqls = sqls & Format(mskFechaIni, "mm/dd/yyyy") & "','" & Format(mskFechaFin, "mm/dd/yyyy") & "',Null,'BuscaEmpleadox'"
   Set rsNombre = New ADODB.Recordset
   rsNombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
'   sqls = " select s.Cliente,c.Nombre NombreCliente,s.Tipo,s.Empleado,Isnull(s.Plaza,1)Plaza,Isnull(s.NoGuia,'SIN GUIA')NoGuia,Isnull(Stock,0) Stock,"
'   sqls = sqls & " Substring(Rtrim(Ltrim(Isnull(s.Nombre,'')))+ ' ' + Rtrim(Ltrim(Isnull(s.Apat,'')))+ ' ' +Rtrim(Ltrim(Isnull(s.AMat,''))),1,26) Nombre,"
'   sqls = sqls & " s.Tipo,convert(varchar,s.FechaSol,101) FechaSol,convert(varchar,s.FechaResp,101)FechaResp"
'   sqls = sqls & " from Solicitudesbe s with (Nolock)"
'   sqls = sqls & " Inner Join Clientes c with (Nolock) On c.Cliente=s.Cliente"
'   sqls = sqls & " where s.cliente = " & cliente
'   sqls = sqls & " and s.Empleado='" & empleado & "'"
'   sqls = sqls & " and s.Producto=" & Product
'   sqls = sqls & " AND s.Fecharesp>='" & Format(mskFechaIni, "mm/dd/yyyy") & "' AND s.Fecharesp<='" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'"
'   sqls = sqls & " and s.Status=2" 'status de Aceptadas, listas para enviar
'   Set rsnombre = New ADODB.Recordset
'   rsnombre.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   
   If Not rsNombre.EOF Then
      client = rsNombre!cliente
      ctte = client
      nombr = rsNombre!Nombre
      emplead = rsNombre!empleado
      Empleadora = rsNombre!NombreCliente
      plaz = rsNombre!Plaza
      Guia = rsNombre!NoGuia
      fech = Format(Date, "dd/mm/yyyy")
      tipin = rsNombre!tipo
      fsol = Format(rsNombre!FechaSol, "mm/dd/yyyy")
      fresp = Format(rsNombre!Fecharesp, "mm/dd/yyyy")
      sistk = rsNombre!Stock
      enbusca = ""
   Else
      enbusca = "No existe"
   End If
     
'   rsnombre.Close
'   Set rsnombre = Nothing
   Exit Sub
End Sub

Sub denuevo()
    If MsgBox("¿Desea capturar otro envio?", vbQuestion + vbYesNo + vbDefaultButton2, "Comenzando otro envio") = vbYes Then
       InicializaForma
    End If
End Sub
Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Cliente"
    frmConsulta.Show vbModal
    
    
    If frmConsulta.cliente >= 0 Then
      txtCliente = frmConsulta.cliente
    End If
    lblcte.Caption = cliente_busca
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Sub crea_guia()
Dim i As Integer
  sqls = "sp_EnvioTarjetasBE " & Product & "," & Val(txtCliente.Text) & ",Null,'"
  sqls = sqls & Format(feche, "mm/dd/yyyy") & "','" & Format(feche, "mm/dd/yyyy") & "',Null,'Crea_guia'"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
'  sqls = "select cliente,tipo,isnull(count(tipo),0) Cantidad,fechaenvio,fecharesp,fechasol from solicitudesbe with (Nolock)"
'  sqls = sqls & " where fechaenvio is not null and status=3 and producto=" & Product
'  sqls = sqls & " and fechaenvio='" & Format(feche, "mm/dd/yyyy") & "'"
'  sqls = sqls & " AND CLIENTE=" & Val(txtCliente.Text)
'  sqls = sqls & " group by cliente,tipo,fechaenvio,fecharesp,fechasol order by fechaenvio, cliente"
'  Set rsBD = New ADODB.Recordset
'  rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
  If Not rsBD.EOF Then
      i = 1
      Do While Not rsBD.EOF
        If txtCliente.Text <> "" Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = " EXEC sp_ControlEnvios_temp "
            sqls = sqls & vbCr & "  @Bodega       = 1"
            sqls = sqls & vbCr & ", @Guia         = " & txtGuia
            sqls = sqls & vbCr & ", @CveMensajeria= " & cboMens.ItemData(cboMens.ListIndex)
            sqls = sqls & vbCr & ", @Refer        = " & i
            sqls = sqls & vbCr & ", @FechaEnvio   = '" & Format(feche, "mm/dd/yyyy") & "'"
            sqls = sqls & vbCr & ", @TipoDoc      = 2"
            sqls = sqls & vbCr & ", @Cliente      = " & Val(txtCliente.Text)
            sqls = sqls & vbCr & ", @Factura      = 0"
            sqls = sqls & vbCr & ", @TipoTar      = '" & rsBD!tipo & "'"
            sqls = sqls & vbCr & ", @FechaRespSB  = '" & Format(rsBD!Fecharesp, "mm/dd/yyyy") & "'"
            sqls = sqls & vbCr & ", @FechaRecSB   = '" & Format(rsBD!Fecharesp, "mm/dd/yyyy") & "'"
            sqls = sqls & vbCr & ", @Total        = " & rsBD!Cantidad
            sqls = sqls & vbCr & ", @Contacto     = ''"
            sqls = sqls & vbCr & ", @status        =1"
            sqls = sqls & vbCr & ", @TipoEnvio    =1"
            sqls = sqls & vbCr & ", @Plaza        =" & Plaza
            sqls = sqls & vbCr & ", @Producto     = " & Product
            cnxbdMty.Execute sqls, intRegistros
        End If
            i = i + 1
            rsBD.MoveNext
      Loop
  End If
'  rsBD.Close
'  Set rsBD = Nothing
  Exit Sub
End Sub


Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtCliente <> "" Then
   sqls = " EXEC sp_BuscaCliente_Datos '" & txtCliente.Text & "','CodCliente'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
   If rsBD.EOF Then
       MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
       txtCliente.SetFocus
'       rsBD.Close
'       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
   Else
      lblCliente.Caption = rsBD!b
      Call CargaComboBE2(cboPlazas, "sp_BuscaCliente_Datos '" & txtCliente.Text & "','PlazasBE'")
      cboPlazas.ListIndex = 0
   End If
End If
End Sub

Sub Imprime(Destino)
Dim stado As String
Dim Result As Integer
    
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEntregas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = 1
    mdiMain.cryReport.StoredProcParam(1) = Val(txtCliente)
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(4) = 3
    mdiMain.cryReport.StoredProcParam(5) = CStr(Product)
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Sub Imprime2(Destino)
  Dim Result As Integer, sql As String
     
    sql = "UPDATE FOLIOS SET Consecutivo=Consecutivo+1"
    sql = sql & " WHERE Tipo='ENV' AND Prefijo='BE'"
    cnxbdMty.Execute sql
     
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCaratulas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(cliente)
    mdiMain.cryReport.StoredProcParam(1) = Val(Usuario)
    mdiMain.cryReport.StoredProcParam(2) = Val(Plaza)
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

