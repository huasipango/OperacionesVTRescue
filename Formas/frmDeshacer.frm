VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeshacer 
   Caption         =   "Modulo reservado para deshacer activaciones de cuentas de Stock"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3435
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
      Begin VB.TextBox txtcliente 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox txtnombre 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   26
         TabIndex        =   8
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtempleado 
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
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtcuenta 
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
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDeshacer.frx":0000
         Left            =   1320
         List            =   "frmDeshacer.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2760
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   2880
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Empleado:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   690
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   1852
      ButtonWidth     =   1455
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Deshacer"
            Key             =   "Deshacer"
            Object.ToolTipText     =   "Deshacer activacion"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpiar"
            Key             =   "Limpia"
            Object.ToolTipText     =   "Limpiar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeshacer.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeshacer.frx":488B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeshacer.frx":91150
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDeshacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim sqls As String

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
   Set mclsAniform = New clsAnimated
   cboProducto.Clear
   Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
   cboProducto.Text = UCase("Despensa Total")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
   Case "Salir"
         Unload Me
   Case "Limpia"
         limpia
   Case "Deshacer"
         Deshacer
   End Select
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Private Sub txtcuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(txtcuenta) <> 0 Then
   sqls = "sp_Activacion_Tarjeta " & Product & ",Null,Null," & Val(txtcuenta) & ",Null,Null,Null,'BuscaCta'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If rsBD.EOF Then
      MsgBox "No existe o no esta activada aun la tarjeta proporcionada", vbCritical, "Cuenta no valida"
   Else
      txtempleado.Text = UCase(Trim(rsBD!empleado))
      txtNombre.Text = UCase(Trim(rsBD!nombre))
      txtcliente.Text = UCase(Trim(rsBD!NomCliente))
   End If
End If
End Sub

Sub limpia()
  txtcuenta.Text = ""
  txtNombre.Text = ""
  txtcliente.Text = ""
  txtempleado.Text = ""
End Sub

Sub Deshacer()
On Error GoTo ERR:
   If MsgBox("¿Esta seguro de desahacer la activacion de esta cuenta [ " & Val(txtcuenta) & " ] de Stock?", vbQuestion + vbYesNo + vbDefaultButton2, "A punto de desactivar cuenta") = vbYes Then
      sqls = "sp_Activacion_Tarjeta " & Product & ",Null,Null," & Val(txtcuenta) & ",Null,Null,Null,'Desactiva'"
      cnxbdMty.Execute sqls
      MsgBox "La cuenta se ha desactivado", vbExclamation, "Cuenta inactiva"
      limpia
   End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub
