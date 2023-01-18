VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAclara 
   Caption         =   "Aclaraciones"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   47
      Top             =   240
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
         ItemData        =   "frmAclara.frx":0000
         Left            =   1800
         List            =   "frmAclara.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4095
      End
      Begin VB.Label Label18 
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
         TabIndex        =   48
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   7200
      Width           =   9135
      Begin VB.CommandButton cmdNuevo 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   500
         Left            =   7200
         Picture         =   "frmAclara.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Nuevo"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   500
         Left            =   6600
         Picture         =   "frmAclara.frx":05B2
         Style           =   1  'Graphical
         TabIndex        =   44
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   500
         Left            =   8400
         Picture         =   "frmAclara.frx":06B4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   500
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   500
         Left            =   7800
         Picture         =   "frmAclara.frx":07B6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   255
         Width           =   500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txtStatusTar 
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
         Left            =   7620
         TabIndex        =   42
         Top             =   1320
         Width           =   1000
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   4680
         Width           =   3615
      End
      Begin VB.ComboBox cboResp 
         Height          =   315
         Left            =   1440
         TabIndex        =   39
         Text            =   "cboResp"
         Top             =   5640
         Width           =   3615
      End
      Begin VB.ComboBox cboProblemas 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2760
         Width           =   7215
      End
      Begin VB.TextBox txtTarjeta1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1440
         TabIndex        =   35
         Text            =   "58877265"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtTarjeta2 
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
         Left            =   2640
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFolio 
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
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtComentarios 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   5160
         Width           =   7335
      End
      Begin VB.TextBox txtMail 
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
         Left            =   1440
         TabIndex        =   10
         Top             =   4200
         Width           =   4215
      End
      Begin VB.TextBox txtSaldo 
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
         Left            =   7080
         TabIndex        =   8
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtTel 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtNoEmpleado 
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
         Left            =   7620
         TabIndex        =   5
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox txtMonto 
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
         Left            =   1440
         TabIndex        =   7
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscarE 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2520
         Picture         =   "frmAclara.frx":08B8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtComercio 
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
         Left            =   1440
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtEmpleado 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2520
         Picture         =   "frmAclara.frx":09BA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFechaProb 
         Height          =   345
         Left            =   7440
         TabIndex        =   18
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaSol 
         Height          =   345
         Left            =   7440
         TabIndex        =   11
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTipoTar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Estatus Tarjeta:"
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
         Left            =   6120
         TabIndex        =   43
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Status:"
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
         Left            =   120
         TabIndex        =   41
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Responsable:"
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
         Left            =   120
         TabIndex        =   38
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta:"
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
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo despues de Movimiento:"
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
         Left            =   3960
         TabIndex        =   33
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label lblComercio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   2280
         Width           =   6015
      End
      Begin VB.Label lblNombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Folio:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha Solucion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   28
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Comentarios:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Mail:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Telefono:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "No. Empleado:"
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
         Left            =   6170
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Mov:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Problema:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Tienda:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Problema:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Empleado:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAclara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim Nuevo As Boolean
Dim prod As Byte

Sub BuscaDatos(valor As Variant, tipoBusc As String)
Dim rsDatos As ADODB.Recordset
Dim j As Integer

Screen.MousePointer = 11
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = " select a.nocuenta Cuenta,  a.noempleado NumEmp, a.nombre NombreEmp,"
sqls = sqls & " convert(varchar(16),dbo.DesEncriptar(b.notarjeta)) Tarjeta, c.nombre NombreCte"
sqls = sqls & " from cuentasbe a , tarjetasbe b, clientes c"
sqls = sqls & " where a.nocuenta = b.nocuenta"
sqls = sqls & " and a.empleadora = c.cliente and b.tipo = 'T' "
sqls = sqls & " and b.Producto=" & Product

If gnMultiBodega = "N" Then
    If gnMultiUEN = "N" Then
        If gnMultiVend = "N" Then
            sqls = sqls & " and a.empleadora = " & gstrUsuario
        Else
            sqls = sqls & " and C.BODEGA = " & gnBodega
        End If
    End If
End If


Select Case tipoBusc
    Case "Nombre"
        sqls = sqls & " and a.nombre like '%" & valor & "%'"
        sqls = sqls & " order by a.nombre"
    Case "Cuenta"
        sqls = sqls & " and a.nocuenta =" & valor & ""
        sqls = sqls & " order by a.nocuenta"
    Case "Tarjeta"
        sqls = sqls & " and convert(varchar(16),dbo.DesEncriptar(b.tarjeta))  like '%" & valor & "%'"
        sqls = sqls & " order by b.tarjeta"
    Case "Empleado"
        sqls = sqls & " and a.noempleado  = '" & valor & "'"
        sqls = sqls & " order by a.noempleado"
End Select

Set rsDatos = New ADODB.Recordset
rsDatos.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

With spdBusca
.Col = -1
.Row = -1
.Action = 12
.MaxRows = 0
j = 0
Do While Not rsDatos.EOF
    j = j + 1
    .MaxRows = j
    .Row = j
    .Col = 1
    .Text = rsDatos!NombreCte
    .Col = 2
    .Text = rsDatos!NumEmp
    .Col = 3
    .Text = rsDatos!NombreEmp
    .Col = 4
    .Text = rsDatos!Tarjeta
    .Col = 5
    .Text = rsDatos!Cuenta
    
    rsDatos.MoveNext

Loop
spdBusca.Visible = True
End With
Screen.MousePointer = 1


End Sub

Function BuscaFolio()
      
sqls = " exec Sp_Folio_Sel_Upd 'SEL', 0, 'ACL'"
       
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    BuscaFolio = rsBD!Folio
Else
    BuscaFolio = 1
End If

End Function

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


Private Sub cmdBuscarE_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Comercios"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtComercio = frmConsulta.cliente
       lblComercio = frmConsulta.Nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False

End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ERR:
'ValidaDatos

Screen.MousePointer = 11
    If Nuevo = True Then txtFolio.Text = BuscaFolio
    

    sqls = " EXEC sp_Aclaraciones "
    sqls = sqls & vbCr & "  @Folio       = " & txtFolio
    sqls = sqls & vbCr & " ,@Fecha         = '" & Format(mskFechaProb.Text, "mm/dd/yyyy") & "'"
    sqls = sqls & vbCr & " ,@Tarjeta      = '" & txtTarjeta1.Text & txtTarjeta2.Text & "'"
    sqls = sqls & vbCr & " ,@Comercio     ='" & Format(txtComercio.Text, "0000000") & "'"
    sqls = sqls & vbCr & ", @Problema    = '" & cboProblemas.Text & "'"
    sqls = sqls & vbCr & ", @Monto      = " & Val(txtMonto.Text)
    sqls = sqls & vbCr & ", @Saldo      = " & Val(txtSaldo.Text)
    sqls = sqls & vbCr & ", @Telefono      = '" & TxtTel.Text & "'"
    sqls = sqls & vbCr & ", @Mail     = '" & txtMail.Text & "'"
    sqls = sqls & vbCr & ", @Status  = " & cbostatus.ItemData(cbostatus.ListIndex)
    If mskFechaSol.Text <> "__/__/____" Then
        sqls = sqls & vbCr & ", @Fechasol   = '" & Format(mskFechaSol.Text, "mm/dd/yyyy") & "'"
    Else
        If cbostatus.ItemData(cbostatus.ListIndex) = 2 Then
            sqls = sqls & vbCr & ", @Fechasol   = '" & Format(Date, "mm/dd/yyyy") & "'"
        End If
        
    End If
    sqls = sqls & vbCr & ", @Comentarios        = '" & txtComentarios.Text & "'"
    sqls = sqls & vbCr & ", @Responsable   = '" & Left(cboResp.Text, InStr(1, cboResp.Text, " -") - 1) & "'"
    
    
    cnxbdMty.Execute sqls, intRegistros
    
    MsgBox "Informacion Actualizada!!!", vbInformation
    
    InicializaForma
    
Screen.MousePointer = 1
Exit Sub
ERR:
MsgBox "Se han detectado inconsistencias en el rellenado del formulario", vbCritical, ERR.Description
Screen.MousePointer = 1
Exit Sub
End Sub

Private Sub cmdNuevo_Click()
 InicializaForma
End Sub

Private Sub cmdPresentar_Click()
    frmRepAclara.Show vbModal
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

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    InicializaForma
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtFolio.Text) <> 0 Then
            sqls = "select * from aclaraciones " & _
                   " where folio = " & Val(txtFolio.Text)
                   
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                'prod = IIf(Product = 8, 6, Product)
                producto_cual
                sqls = "exec spr_repaclaraciones " & Val(txtFolio) & ", 0, '" & Format(rsBD!Fecha, "mm/dd/yyyy") & "', '" & Format(rsBD!Fecha, "mm/dd/yyyy") & "'," & Product
                Set rsBD = New ADODB.Recordset
                rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                
                    If Not rsBD.EOF Then
                        txtTarjeta2.Text = Right(rsBD!Tarjeta, 8)
                        txtCliente.Text = rsBD!Empleadora
                        lblNombre.Caption = Trim(rsBD!Nombre)
                        txtStatusTar.Text = IIf(rsBD!StatusTar = 2, "CANCELADA", "ACTIVA")
                        txtEmpleado.Text = rsBD!NombreEmpleado
                        txtNoEmpleado.Text = rsBD!noempleado
                        txtComercio.Text = rsBD!comercio
                        lblComercio.Caption = IIf(IsNull(rsBD!desccomercio), "", rsBD!desccomercio)
                        cboProblemas.Text = rsBD!problema
                        txtMonto.Text = rsBD!Monto
                        txtSaldo.Text = rsBD!saldo
                        TxtTel.Text = rsBD!Telefono
                        txtMail.Text = rsBD!Mail
                        Call CboPosiciona(cbostatus, rsBD!Status)
                        mskFechaProb.Text = IIf(IsNull(rsBD!Fecha), "__/__/____", rsBD!Fecha)
                        mskFechaSol.Text = IIf(IsNull(rsBD!FechaSol), "__/__/____", rsBD!FechaSol)
                        txtComentarios.Text = rsBD!comentarios
                        
                        cboResp.Text = Trim(rsBD!responsable) & " - " & UCase(Trim(rsBD!DescResponsable))
                        Nuevo = False
                    Else
                        MsgBox "Folio no existe"
                        
                                  
                    End If
               
            Else
                MsgBox "El folio no existe.", vbInformation
                txtFolio.Text = ""
                txtFolio.SetFocus
                
                
            End If
                 
       End If
        End If
End Sub

Private Sub txtTarjeta2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
Sub InicializaForma()

    LimpiarControles Me
    mskFechaProb.Text = Date
    txtFolio = BuscaFolio
    Nuevo = True
        
    If Product = 1 Then
       txtTarjeta1.Text = "50640601"
    ElseIf Product = 2 Then
      txtTarjeta1.Text = "50640501"
    End If
    
    CargaUsuarios cboResp
    CargaProblemas cboProblemas
    CargaStatus cbostatus
    lblNombre.Caption = ""
    lblComercio.Caption = ""
    lblTipoTar.Caption = ""
End Sub
Private Sub txtTarjeta2_LostFocus()
    If Trim(txtTarjeta2.Text) <> "" Then
    
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        sqls = "sp_Consultas_BE Null,Null," & Product & ",NULL,'" & txtTarjeta1.Text & txtTarjeta2.Text & "',Null,Null,'Aclaraciones'"
        
'        sqls = "select a.empleadora cliente, a.noempleado numempleado, b.nombre ,c.nombre Empleadora, b.status statustarjeta, b.tipo" & _
'               " from cuentasbe a, tarjetasbe b, clientes c" & _
'               " Where a.nocuenta = b.nocuenta and a.empleadora = c.cliente" & _
'               " and b.notarjeta = '" & txtTarjeta1.Text & txttarjeta2.Text & "'" & _
'               " and b.Producto=" & prod
               
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            
        If Not rsBD.EOF Then
                txtCliente.Text = rsBD!cliente
                txtEmpleado.Text = Trim(rsBD!Nombre)
                txtNoEmpleado.Text = rsBD!numempleado
                txtStatusTar.Text = IIf(rsBD!statustarjeta = 1, "ACTIVA", "CANCELADA")
                lblNombre.Caption = rsBD!Empleadora
                lblTipoTar.Caption = IIf(Right(Trim(rsBD!tipo), 1) = "T", "TITULAR", "ADICIONAL")
        Else
            MsgBox "La tarjeta no existe", vbInformation
            InicializaForma
            
        End If
    End If
End Sub
