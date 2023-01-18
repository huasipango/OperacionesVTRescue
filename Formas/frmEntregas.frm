VERSION 5.00
Begin VB.Form frmEntrega 
   Caption         =   "Datos de entrega del Cliente"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContac 
      Height          =   375
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   23
      Top             =   5880
      Width           =   3735
   End
   Begin VB.TextBox txtProducto 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   21
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox txtcd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   20
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtestado 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox cboEstados 
      Height          =   315
      ItemData        =   "frmEntregas.frx":0000
      Left            =   2400
      List            =   "frmEntregas.frx":0002
      TabIndex        =   18
      Text            =   "cboEstados"
      Top             =   3555
      Width           =   2610
   End
   Begin VB.ComboBox cboPoblaciones 
      Height          =   315
      Left            =   2400
      TabIndex        =   17
      Text            =   "cboPoblaciones"
      Top             =   4080
      Width           =   2610
   End
   Begin VB.TextBox txtTel 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   15
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox txtCP 
      Height          =   375
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtcolonia 
      Height          =   375
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   2
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox txtcalle 
      Height          =   375
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   1320
      MaxLength       =   34
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtcliente 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   450
      Left            =   5040
      Picture         =   "frmEntregas.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   6840
      Width           =   450
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   450
      Left            =   5880
      Picture         =   "frmEntregas.frx":0106
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   6840
      Width           =   450
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Contacto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   24
      Top             =   6000
      Width           =   780
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   22
      Top             =   6600
      Width           =   795
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   16
      Top             =   5400
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "C. P.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   14
      Top             =   4800
      Width           =   405
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   11
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Colonia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Capture y Valide los datos donde el Cliente recibirá los estados de cuenta."
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
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   6150
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Calle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   450
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   4920
      Picture         =   "frmEntregas.frx":0208
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1425
   End
End
Attribute VB_Name = "frmEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim desEST As String



Private Sub cboEstados_Click()
 CargaPoblaciones cboPoblaciones, cboEstados.ItemData(cboEstados.ListIndex)
End Sub

Private Sub cmdGrabar_Click()
If txtcd = "" Or txtestado = "" Or cboEstados = "" Or cboPoblaciones = "" Then
   MsgBox "Le falta algun...revise bien la informacion", vbExclamation, "Falta informacion"
   txtnombre.SetFocus
   Exit Sub
End If

sqls = "SELECT DescCorta FROM Estados WHERE Descripcion='" & Trim(cboEstados.Text) & "'"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   desEST = rsBD!DescCorta
Else
   MsgBox "Error...no existe el estado", vbCritical, "Error en el estado"
   Exit Sub
End If
  
  sqls = "sp_DatosEnvio @Cliente=" & txtcliente & ","
  sqls = sqls & " @Nombre='" & Trim(UCase(txtnombre)) & "',"
  sqls = sqls & " @Calle='" & Trim(UCase(txtcalle)) & "',"
  sqls = sqls & " @Colonia='" & Trim(UCase(txtcolonia)) & "',"
  sqls = sqls & " @Cd='" & Trim(cboPoblaciones.Text) & "',"
  sqls = sqls & " @Estado='" & Trim(desEST) & "',"
  sqls = sqls & " @CP='" & Trim(txtCP) & "',"
  sqls = sqls & " @Tel='" & Trim(txtTel) & "',"
  sqls = sqls & " @Contacto='" & Trim(UCase(txtContac)) & "',"
  sqls = sqls & " @Producto=" & Gprod
  cnxBD.Execute sqls, intRegistros
  
  MsgBox "Los datos se han actualizado", vbInformation, "Actualizacion correcta"
  Exit Sub
 
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim descor As String, pob As String
Dim contac As String
   CargaEstados cboEstados
   CboPosiciona cboEstados, 1
   CargaPoblaciones cboPoblaciones, cboEstados.ItemData(cboEstados.ListIndex)
   sqls = "SELECT * FROM clientes_envios WHERE Cliente=" & GCliente
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
   If rsBD.EOF Then
      txtcliente = GCliente
      txtnombre = GNombre
      txtcalle = GCalle
      txtcolonia = GColonia
      txtTel = GTel
      txtCP = GCP
      txtestado = GEstado
      txtcd = GCd
      txtproducto = GprodText
      
      
      sqls = "SELECT Descripcion FROM Estados where Estado=" & GEstado
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
        cboEstados.Text = Trim(rsBD!Descripcion)
      Else
        MsgBox "No existe el estado", vbCritical, "Error en el estado"
        Exit Sub
      End If
   
      sqls = "SELECT Descripcion FROM Poblaciones where POBLACION=" & GCd
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
        cboPoblaciones.Text = Trim(rsBD!Descripcion)
      Else
        MsgBox "No existe la ciudad", vbCritical, "Error en la ciudad"
        Exit Sub
      End If
   Else
      txtcliente = rsBD!cliente
      txtnombre = rsBD!nombre
      txtcalle = rsBD!Calle
      txtcolonia = rsBD!Colonia
      txtTel = rsBD!Telefono
      txtCP = rsBD!CP
      txtproducto = GprodText
      contac = IIf(IsNull(rsBD!contacto), "", Trim(rsBD!contacto))
      txtContac = Trim(contac)
      descor = rsBD!estado
      pob = rsBD!Cd
      sqls = "SELECT * FROM ESTADOS WHERE DescCorta='" & descor & "'"
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      
      If Not rsBD.EOF Then
         txtestado = rsBD!estado
      Else
        MsgBox "Error en el estado", vbCritical, "Error en el estado"
        Exit Sub
      End If
      
      sqls = "SELECT Descripcion FROM Estados where Estado=" & txtestado
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
        cboEstados.Text = Trim(rsBD!Descripcion)
      Else
        MsgBox "No existe el estado", vbCritical, "Error en el estado"
        Exit Sub
      End If

      CARGA_POBLAC
      
      sqls = "SELECT * FROM POBLACIONES WHERE ESTADO=" & txtestado & " AND DESCRIPCION='" & Trim(pob) & "'"
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      
      If Not rsBD.EOF Then
         txtcd = rsBD!poblacion
      Else
        MsgBox "Error en la poblacion", vbCritical, "Error en la poblacion"
        Exit Sub
      End If
      
      sqls = "SELECT Descripcion FROM Poblaciones where POBLACION=" & txtcd
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
        cboPoblaciones.Text = Trim(rsBD!Descripcion)
      Else
        MsgBox "No existe la ciudad", vbCritical, "Error en la ciudad"
        Exit Sub
      End If
       
   End If
End Sub


Sub CARGA_POBLAC()
cboPoblaciones.Clear
sqls = "SELECT DESCRIPCION FROM POBLACIONES WHERE ESTADO=" & txtestado
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   Do While Not rsBD.EOF
      cboPoblaciones.AddItem rsBD!Descripcion
      rsBD.MoveNext
   Loop
Else
  MsgBox "No existe la ciudad", vbCritical, "Error en la ciudad"
  Exit Sub
End If
End Sub



