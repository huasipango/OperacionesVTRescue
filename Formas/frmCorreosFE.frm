VERSION 5.00
Begin VB.Form frmCorreosFE 
   Caption         =   "Correo Electronico del Cliente"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   6975
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   5640
         Picture         =   "frmCorreosFE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   6240
         Picture         =   "frmCorreosFE.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblinfo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   2400
         Picture         =   "frmCorreosFE.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   285
      End
      Begin VB.CheckBox chkMailFE 
         Caption         =   "Enviar Factura Electronica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtMailFEto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1320
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
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
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboBodegas 
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
         Height          =   315
         ItemData        =   "frmCorreosFE.frx":0306
         Left            =   1320
         List            =   "frmCorreosFE.frx":030D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Las cuentas deben ir separadas por (;) cuenta1@mail.com; cuenta2@mail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblNombre 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCorreosFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim labodega As Byte


Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "ClienteBE"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtCliente = frmConsulta.cliente
       lblNombre = frmConsulta.nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
    txtCliente.SetFocus
    txtCliente_KeyPress (13)
End Sub

Private Sub cmdGrabar_Click()
Dim pslq As String
    If Trim(txtCliente) <> "" Then
        sqls = "exec Sp_ClientesConfigStock_Upd " & cboBodegas.ItemData(cboBodegas.ListIndex) & ", " & txtCliente & ", " & IIf(chkMailFE.value, 1, 0) & ", '" & txtMailFeTo.Text & "'"
        cnxBD.Execute sqls, intRegistros
        MsgBox "Informacion actualizada!!", vbInformation, "Informacion actualizada"
        lblNombre.Caption = ""
        txtCliente.Text = ""
        chkMailFE.value = 0
        txtMailFeTo.Text = ""
        lblinfo.Caption = ""
    End If
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Set mclsAniform = New clsAnimated
      
   CargaBodegas cboBodegas
   Call CboPosiciona(cboBodegas, gnBodega)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Dim bodega1 As Byte
    If KeyAscii = 13 And Trim(txtCliente.Text) <> "" Then
        
        sqls = "Select * from clientes Where Cliente=" & txtCliente.Text
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
        If Not rsBD.EOF Then
           bodega1 = rsBD!Bodega
        Else
           MsgBox "El cliente no existe", vbCritical, "Verifique..."
           txtCliente.Text = ""
           Exit Sub
        End If
        
        If bodega1 = gnBodega Then
           sqls = " exec sp_Clientes_sel " & cboBodegas.ItemData(cboBodegas.ListIndex) & ", " & txtCliente.Text & ""
        
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
           If Not rsBD.EOF Then
              labodega = rsBD!Bodega
              lblNombre = rsBD!nombre
              SendKeys "{TAB}"
           Else
              MsgBox " Cliente no existe", vbCritical, "Cliente no existe"
              Exit Sub
           End If
         Else
           MsgBox "Cliente no existe o no pertenece a la sucursal", vbCritical, "Error con cliente"
           txtCliente.Text = ""
           txtMailFeTo.Text = ""
           lblNombre.Caption = ""
           lblinfo.Caption = ""
           txtCliente.SetFocus
         End If
    End If
    Exit Sub
End Sub

Private Sub txtCliente_LostFocus()
Dim maile, mailt
On Error GoTo ERR:
    If txtCliente.Text <> "" Then
        sqls = "SELECT isnull(mailfeto,'') mailfeto,isnull(mailfe,0) mailfe FROM CLIENTESCONFIG WHERE CLIENTE=" & Val(txtCliente.Text) & _
        " and bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
        If Not rsBD.EOF Then
           txtMailFeTo.Text = rsBD!MailFETo
           chkMailFE.value = rsBD!MailFe
        End If
    End If
    txtMailFeTo.SetFocus
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
Exit Sub
End Sub


