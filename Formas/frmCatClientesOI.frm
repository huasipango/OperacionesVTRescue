VERSION 5.00
Begin VB.Form frmCatClientesOI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de Clientes OI"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPrincipal 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   9015
      Begin VB.CommandButton cmdNuevo 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   1200
         Picture         =   "frmCatClientesOI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Nuevo"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   3000
         Picture         =   "frmCatClientesOI.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         Picture         =   "frmCatClientesOI.frx":0698
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   2400
         Picture         =   "frmCatClientesOI.frx":079A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Tag             =   "Key"
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Baja"
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
         Left            =   7515
         TabIndex        =   28
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha Modificacion"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblFecIngresoT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Fecha deAlta"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblFecBaja 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   7290
         TabIndex        =   25
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblFecCambio 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   5685
         TabIndex        =   24
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblFecIngresoV 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   4080
         TabIndex        =   23
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label lblCliente 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11655
      Begin VB.Frame Frame1 
         Caption         =   "Factura Electronica"
         Height          =   735
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   8655
         Begin VB.TextBox txtMailFeTo 
            Height          =   315
            Left            =   3345
            MaxLength       =   200
            TabIndex        =   39
            Tag             =   "Det"
            ToolTipText     =   "Direcciones de correos separadas por comas (,)"
            Top             =   240
            Width           =   5130
         End
         Begin VB.CheckBox chkMailFE 
            Caption         =   "Enviar por correo"
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
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2700
            TabIndex        =   40
            Top             =   320
            Width           =   540
         End
      End
      Begin VB.ComboBox cbostatus 
         Height          =   315
         Left            =   8745
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2280
         Width           =   2610
      End
      Begin VB.ComboBox cboSucursal 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "Key"
         Top             =   360
         Width           =   2280
      End
      Begin VB.ComboBox cboPoblaciones 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   2610
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Det"
         Top             =   1320
         Width           =   3660
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   1
         Tag             =   "Det"
         Top             =   840
         Width           =   6300
      End
      Begin VB.TextBox txtColonia 
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Det"
         Top             =   1800
         Width           =   3690
      End
      Begin VB.TextBox txtEntreCalles 
         Height          =   315
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Det"
         Top             =   2280
         Width           =   3690
      End
      Begin VB.ComboBox cboEstados 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2610
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   8760
         MaxLength       =   14
         TabIndex        =   7
         Tag             =   "Det"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   315
         Left            =   10440
         MaxLength       =   14
         TabIndex        =   8
         Tag             =   "Det"
         Top             =   1320
         Width           =   870
      End
      Begin VB.TextBox txtRFC 
         Height          =   285
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "Det"
         Top             =   1800
         Width           =   2595
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Estatus:"
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
         Height          =   240
         Left            =   7680
         TabIndex        =   36
         Top             =   2280
         Width           =   1020
      End
      Begin VB.Label lblBodega 
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
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblNombre 
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
         Left            =   135
         TabIndex        =   18
         Top             =   900
         Width           =   825
      End
      Begin VB.Label lblDomicilio 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblPobl 
         Appearance      =   0  'Flat
         Caption         =   "Población:"
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
         Height          =   240
         Left            =   7695
         TabIndex        =   16
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblColonia 
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
         Left            =   135
         TabIndex        =   15
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblEntreCalles 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Entre calles:"
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
         Left            =   135
         TabIndex        =   14
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         Caption         =   "Estado:"
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
         Height          =   240
         Left            =   7695
         TabIndex        =   13
         Top             =   345
         Width           =   1200
      End
      Begin VB.Label lblCP 
         Alignment       =   1  'Right Justify
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
         Left            =   9960
         TabIndex        =   12
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lblTelefono 
         Appearance      =   0  'Flat
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
         Height          =   240
         Left            =   7695
         TabIndex        =   11
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblRFC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
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
         Left            =   7695
         TabIndex        =   10
         Top             =   1800
         Width           =   360
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Clientes Otros Ingresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmCatClientesOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim TipoCambio As String

Private Sub cboEstados_Click()
CargaPoblaciones cboPoblaciones, cboEstados.ItemData(cboEstados.ListIndex)

End Sub

Private Sub cmdActualizar_Click()
On Error GoTo err_grab

TipoCancel = Null

If Valida_RFC(Trim(txtRFC.Text), True) Then

    sqls = "exec sp_ClientesOI_upd" & _
           " @Bodega     = " & cboSucursal.ItemData(cboSucursal.ListIndex) & _
           " ,@Cliente    = " & Val(txtCliente.Text) & _
           " ,@Nombre     = '" & txtnombre.Text & "'" & _
           " ,@Domicilio  =  '" & txtDomicilio.Text & "'" & _
           " ,@Colonia    = '" & txtcolonia.Text & "'" & _
           " ,@EntreCalles =   '" & txtEntreCalles.Text & "'" & _
           " ,@Poblacion   =" & cboPoblaciones.ItemData(cboPoblaciones.ListIndex) & _
           " ,@Estado  =" & cboEstados.ItemData(cboEstados.ListIndex) & _
           " ,@Telefono    =  '" & txtTelefono.Text & "'" & _
           " ,@CodigoPostal =  '" & txtCodigoPostal.Text & "'" & _
           " ,@RFC         =  '" & Trim(txtRFC.Text) & "'" & _
           " ,@Status    =" & cbostatus.ItemData(cbostatus.ListIndex) & _
           " ,@usuario  = '" & gstrUsuario & "'" & _
           " ,@MailFe = " & chkMailFE.Value & _
           " ,@MailFeTo = '" & txtMailFeTo.Text & "'"
           
                      
           If cbostatus.ItemData(cbostatus.ListIndex) = 2 Then
               sqls = sqls & " ,@TipoCancel  = " & cbostatus.ItemData(cbostatus.ListIndex)
           End If
           
           Set rsBD = New ADODB.Recordset
           rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
           
           If Not rsBD.EOF Then
                Cliente = rsBD!Cliente
           End If
            
            
           If Trim(txtCliente.Text) = "" Then
                MsgBox "El numero de cliente es :" & Cliente & "", vbInformation, "Clientes"
                txtCliente.Text = Cliente
            Else
                MsgBox "Cliente " & Cliente & " actualizado!", vbInformation, "Clientes"
            End If
                             
    End If
           
    Exit Sub
err_grab:
    MsgBox "Error al grabar cliente: " & ERR.Description, vbCritical, "ClienteOI"
      
End Sub

Private Sub cmdBuscar_Click()
Dim frmConsulta As New frmBusca_Cliente
    
    frmConsulta.bodega = cboSucursal.ItemData(cboSucursal.ListIndex)
    frmConsulta.Show vbModal
   
    If frmConsulta.Cliente > 0 Then
       txtCliente = frmConsulta.Cliente
       txtnombre = Trim(Mid(frmConsulta.nombre, 1, 90))
       If Trim(txtCliente.Text) <> "" Then
            If Not ValidaLlave() Then Exit Sub
                If Not CargaDatos() Then
                    txtCliente.SetFocus
                    Exit Sub
                End If
            Else
            blnNew = True
            End If
            Call Switch(Me, "Det", True)
            Call Switch(Me, "Key", False)
            cmdActualizar.Enabled = True
    End If
    Set frmConsulta = Nothing
End Sub
Function ValidaLlave()
  
    ValidaLlave = False
    
    If Not IsNumeric(txtCliente) Then Exit Function
    
    ValidaLlave = True

End Function

Function CargaDatos() As Boolean
   CargaDatos = False
   CanalVta = 0

'   Datos del Cliente
   sqls = " "
   sqls = "EXEC   sp_clientesOI_sel "
   sqls = sqls & "   @Bodega  =   " & cboSucursal.ItemData(cboSucursal.ListIndex)
   sqls = sqls & " , @Cliente = " & Val(txtCliente)
    
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
   If rsBD.EOF = False Then
      lblFecIngresoV.Caption = UCase(Format(rsBD!Fecha_Alta, "DD/MMM/YYYY"))
      lblFecCambio.Caption = UCase(Format(rsBD!Fecha_Cambio, "DD/MMM/YYYY"))
      
      If rsBD!nombre = Null Or IsNull(rsBD!nombre) Then
      Else
         txtnombre = rsBD!nombre
      End If
      
      If rsBD!Domicilio <> "" Then txtDomicilio = rsBD!Domicilio
      If rsBD!EntreCalles <> "" Then txtEntreCalles = rsBD!EntreCalles
      If rsBD!Colonia <> "" Then txtcolonia = rsBD!Colonia
      If rsBD!Telefono <> "" Then txtTelefono = rsBD!Telefono
      If rsBD!Rfc <> "" Then txtRFC = rsBD!Rfc
      If rsBD!codigopostal <> "" Then txtCodigoPostal = rsBD!codigopostal
      If IsNull(rsBD!estado) = False Then Call CboPosiciona(cboEstados, rsBD!estado)
      If rsBD!poblacion <> 0 Then Call CboPosiciona(cboPoblaciones, rsBD!poblacion)
      If rsBD!NomPoblacion <> "" Then lblPoblacion = rsBD!NomPoblacion
      If rsBD!NomPoblacion <> "" Then lblPoblacion = rsBD!NomPoblacion
      chkMailFE.Value = rsBD!MailFe
       txtMailFeTo.Text = rsBD!MailFETo
      
      
      If IsNull(rsBD!bodega) = False Then Call CboPosiciona(cboSucursal, rsBD!bodega)
      Call CboPosiciona(cbostatus, rsBD!Status)
             
            
   Else
      If sUbicacion = "01" Then
         If mstrTipo = "A" Then
            If MsgBox("El Cliente # " & Trim(txtCliente) & " No Existe En: " + vbCr + vbCr + "                 << Desea darlo de Alta >>", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
               CargaDatos = True
               blnNew = True
            End If
         End If
      Else
         MsgBox "El Cliente no existe,  favor de solicitar la alta al departamento de creditos!!!", vbCritical, Me.Caption
      End If
      Exit Function
   End If
   CargaDatos = True
   
End Function

Private Sub cmdNuevo_Click()
    LimpiarControles Me
    CargaBodegas cboSucursal
    CargaEstados cboEstados
    txtCliente = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    LimpiarControles Me
    CargaBodegas cboSucursal
    CargaEstados cboEstados
    CboPosiciona cboEstados, 1
    CargaPoblaciones cboPoblaciones, cboEstados.ItemData(cboEstados.ListIndex)
    CargaStatusCliente cbostatus
    TipoCambio = "A"
    TipoBusqueda = "ClientesOI"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
