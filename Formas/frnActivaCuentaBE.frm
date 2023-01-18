VERSION 5.00
Begin VB.Form frmCambiarEstatusTJ 
   Caption         =   "Activar Cuentas"
   ClientHeight    =   4095
   ClientLeft      =   4260
   ClientTop       =   4725
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5500
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
         ItemData        =   "frnActivaCuentaBE.frx":0000
         Left            =   1100
         List            =   "frnActivaCuentaBE.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4095
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
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   5535
      Begin VB.CommandButton btnAceptar 
         Caption         =   "&Activar"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      Begin VB.TextBox txtNumTarjeta 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Pulsar la tecla Enter"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label label1 
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblNombreCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblClaveCte 
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCambiarEstatusTJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub btnAceptar_Click()
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "sp_CuentasBE_Varios " & Me.txtNumTarjeta.Text & "," & Product & ",'Activa'"

'sqls = "Update CuentasBE set Status = 1 WHERE NoCuenta= '" & Me.txtNumTarjeta.Text & "'"
'sqls = sqls & " AND Producto=" & Product

cnxBD.Execute sqls, Status
MsgBox "La cuenta se activó correctamente", vbOKOnly, "Cuenta Activa"
    
Me.txtNumTarjeta.Text = ""
Me.lblClaveCte.Caption = ""
Me.lblNombreCliente.Caption = ""

End Sub

Private Sub btnSalir_Click()
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub


Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Sub InicializaForma()
    txtNumTarjeta.Text = ""
    lblClaveCte.Caption = ""
    lblNombreCliente.Caption = ""
End Sub
Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
Set mclsAniform = New clsAnimated
Me.txtNumTarjeta.Text = ""
cboProducto.Clear
Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
cboProducto.Text = UCase("Despensa Total")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtNumTarjeta_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "sp_CuentasBE_Varios " & Me.txtNumTarjeta.Text & "," & Product & ",'Cuenta'"
    
'    sqls = "select status,nocuenta, empleadora, nombre from CuentasBE" & _
'    " WHERE NoCuenta= '" & Me.txtNumTarjeta.Text & "' and Producto=" & Product
        
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
        lblNombreCliente.Caption = Trim(rsBD!nombre)
        Me.lblClaveCte.Caption = Trim(rsBD!Empleadora)
        If rsBD!Status = 1 Then
            MsgBox "Esta cuenta ya esta activa!! ", vbCritical, "Cuenta activa"
        End If
    Else
        MsgBox "La cuenta no existe o no esta dada de alta en el sistema", vbCritical, "Error en cuenta"
        Me.txtNumTarjeta.Text = ""
    End If

End If
End Sub


