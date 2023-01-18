VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmProveedores 
   Caption         =   "Catalogo de Comercios"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   13170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEstados 
      Height          =   2175
      Left            =   4478
      TabIndex        =   40
      Top             =   5880
      Width           =   4215
      Begin VB.CommandButton cmdCancelarC 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdAceptarC 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox cboPoblacionesC 
         Height          =   315
         Left            =   1320
         TabIndex        =   44
         Text            =   "Combo3"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cboEstadosC 
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         Text            =   "Combo3"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Población:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBorrar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borrar"
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
      Height          =   550
      Left            =   11640
      Picture         =   "frmProvedores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Borrar Grupo"
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
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
      Height          =   550
      Left            =   10920
      Picture         =   "frmProvedores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancelar alta"
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nuevo"
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
      Height          =   550
      Left            =   10200
      Picture         =   "frmProvedores.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo Grupo"
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Por Comercio"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   4320
      Picture         =   "frmProvedores.frx":0ADA
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CommandButton cmdAbrir 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   450
      Left            =   3720
      Picture         =   "frmProvedores.frx":0BDC
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
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
      Height          =   550
      Left            =   12360
      Picture         =   "frmProvedores.frx":0CDE
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grabar"
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
      Height          =   550
      Left            =   9480
      Picture         =   "frmProvedores.frx":0DE0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Grabar"
      Top             =   120
      Width           =   550
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comercio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   12855
      Begin FPSpread.vaSpread spdComercios 
         Height          =   3375
         Left            =   240
         OleObjectBlob   =   "frmProvedores.frx":0EE2
         TabIndex        =   15
         Top             =   360
         Width           =   12375
      End
      Begin VB.CommandButton cmdBuscarComercio 
         Caption         =   "Buscar Comercio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   51
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdBorrarC 
         Caption         =   "Eliminar Comercio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   38
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdGuardarC 
         Caption         =   "Guardar Comercio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   37
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdNuevoC 
         Caption         =   "Agregar Comercio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   36
         Top             =   3960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12855
      Begin VB.CheckBox chk1 
         Caption         =   "Relacionar con otros productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9960
         TabIndex        =   54
         Top             =   3360
         Width           =   2655
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
         ItemData        =   "frmProvedores.frx":1348
         Left            =   5040
         List            =   "frmProvedores.frx":1352
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   240
         Width           =   3120
      End
      Begin VB.TextBox txtCveEstab 
         Height          =   300
         Left            =   8160
         TabIndex        =   50
         Top             =   2880
         Width           =   885
      End
      Begin VB.ComboBox cboGiros 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   8775
      End
      Begin VB.TextBox txtTel 
         Height          =   300
         Left            =   8160
         TabIndex        =   12
         Top             =   2450
         Width           =   3400
      End
      Begin VB.TextBox txtGrupo 
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
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   350
         Left            =   1320
         Picture         =   "frmProvedores.frx":1364
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   720
         Width           =   350
      End
      Begin VB.TextBox txtCtaCont 
         Height          =   300
         Left            =   8160
         TabIndex        =   14
         Top             =   3360
         Width           =   1485
      End
      Begin VB.TextBox txtContacto 
         Height          =   300
         Left            =   8160
         TabIndex        =   10
         Top             =   2000
         Width           =   3400
      End
      Begin VB.TextBox txtComision 
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtmail 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   2450
         Width           =   2655
      End
      Begin VB.TextBox txtRFC 
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   2000
         Width           =   2655
      End
      Begin VB.ComboBox cboEstados 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "Combo3"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox cboPoblaciones 
         Height          =   315
         Left            =   8160
         TabIndex        =   8
         Text            =   "Combo3"
         Top             =   1560
         Width           =   3400
      End
      Begin VB.TextBox txtCol 
         Height          =   300
         Left            =   8160
         TabIndex        =   6
         Top             =   1150
         Width           =   3400
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   1150
         Width           =   4335
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Giro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2925
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cve Estab:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6600
         TabIndex        =   47
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label13 
         Caption         =   "Telefono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Cuenta Contable:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   28
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Comisión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3405
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2450
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   25
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "RFC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2050
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1605
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Población:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1605
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Colonia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   760
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Catálogo de Comercios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   33
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Public NvoCom As Boolean
Dim prod As Byte
Dim elrfc As String, sqls As String
Private Sub CboEstado_KeyPress(KeyAscii As Integer)
    entertab (KeyAscii)
End Sub

Private Sub cboEstados_Click()
    If cboEstados.Text <> "" Then
         CargaPoblaciones cboPoblaciones, cboEstados.ItemData(cboEstados.ListIndex)
    End If
End Sub

Private Sub cboEstados_KeyPress(KeyAscii As Integer)
    entertab (KeyAscii)
End Sub

Private Sub cboGrupos_KeyPress(KeyAscii As Integer)
 entertab (KeyAscii)
End Sub

Private Sub cboImprimir_KeyPress(KeyAscii As Integer)
    entertab (KeyAscii)
End Sub

Private Sub CboMunicipio_KeyPress(KeyAscii As Integer)
    entertab (KeyAscii)
End Sub

Private Sub cboEstadosC_Click()
If cboEstadosC.ListIndex >= 0 Then
   CargaPoblaciones cboPoblacionesC, cboEstadosC.ItemData(cboEstadosC.ListIndex)
End If
End Sub

Private Sub chk1_Click()
  elgrupo = Val(txtGrupo.Text)
  prod_anterior = Product
  nombre_com = Trim(txtDescripcion.Text)
  If Val(txtGrupo.Text) = 0 Then
     MsgBox "Debe de iniciar con un numero de grupo para algun producto", vbExclamation, "No hay Grupo para iniciar"
     chk1.value = 0
     Exit Sub
  End If
 
  If chk1.value = 1 Then
     frmComisionGpo.Show vbModal
  End If
End Sub

Private Sub cmdAceptarC_Click()
   If cboEstadosC.Text = "" Then
      MsgBox "Seleccione un estado", vbInformation, "Estados..."
      Exit Sub
   End If
   If cboPoblacionesC.Text = "" Then
      MsgBox "Seleccione una población", vbInformation, "Poblaciones..."
      Exit Sub
   End If
      
   frmEstados.Visible = False
   spdComercios.Row = spdComercios.ActiveRow
   spdComercios.Col = 5
   spdComercios.Text = cboEstadosC.Text & " (" & cboEstadosC.ItemData(cboEstadosC.ListIndex) & ")"
   spdComercios.Col = 6
   spdComercios.Text = cboPoblacionesC.Text & " (" & cboPoblacionesC.ItemData(cboPoblacionesC.ListIndex) & ")"
   
End Sub



Private Sub cmdActualizar_Click()
   Call verifica_RFC
   If cboPoblaciones.Text = "" Then
      MsgBox "Primero debe seleccionar la Poblacion", vbCritical, "Poblaciones..."
      Exit Sub
   End If
   
   If cboEstados.Text = "" Then
      MsgBox "Primero debe seleccionar el Estado", vbCritical, "Estados..."
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   On Error GoTo err_gral
   
   If NvoCom = True Then
      BuscaFolioGrupo
   End If
    
    'If txtCveEstab.Text = "" Then  Lo quite porque no estamos grabando en establecimientos vsp 25/Abr/2017
    '    MsgBox "Debe capturar el numero de establecimiento", vbCritical, "Establecimiento"
    '    Exit Sub
    'Else
    '    sqls = "select cveestablecimiento from establecimientos" & _
    '           " where cveestablecimiento = " & IIf(txtCveEstab.Text <> "", txtCveEstab, 0)
    '    Set rsBD = New ADODB.Recordset
    '    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    '
    '    If rsBD.EOF Then
    '        MsgBox "Clave de establecimiento no existe en el catalogo de establecimientos", vbCritical, "No existe establecimiento"
    '        Exit Sub
    '    End If
    'End If
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
'    If Product <> 2 Then
       Call Actualiza_gpos
'    Else
'       sqls = " EXEC sp_InsUpd_Grupos "
'       sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
'       sqls = sqls & vbCr & ", @Grupo      = " & txtGrupo
'       sqls = sqls & vbCr & ", @Descripcion  =    '" & IIf(Trim(txtDescripcion) = "", " ", Trim(txtDescripcion)) & "'"
'       sqls = sqls & vbCr & ", @Domicilio    =    '" & IIf(Trim(txtDomicilio) = "", " ", Trim(txtDomicilio)) & "'"
'       sqls = sqls & vbCr & ", @Colonia      =    '" & IIf(Trim(txtCol) = "", " ", Trim(txtCol)) & "'"
'       sqls = sqls & vbCr & ", @Poblacion    =    " & IIf(cboPoblaciones.Text <> "", cboPoblaciones.ItemData(cboPoblaciones.ListIndex), 0)
'       sqls = sqls & vbCr & ", @Estado    =    " & IIf(cboEstados.Text <> "", cboEstados.ItemData(cboEstados.ListIndex), 0)
'       sqls = sqls & vbCr & ", @rfc    =  '" & IIf(Trim(damerfc(txtRFC)) = "", " ", Trim(damerfc(txtRFC))) & "'"
'       sqls = sqls & vbCr & ", @Contacto       = '" & IIf(Trim(txtContacto) = "", " ", Trim(txtContacto)) & "'"
'       sqls = sqls & vbCr & ", @Telefono = '" & IIf(Trim(txttel) = "", " ", Trim(txttel)) & "'"
'       sqls = sqls & vbCr & ", @email     ='" & IIf(Trim(txtmail) = "", " ", Trim(txtmail)) & "'"
'       sqls = sqls & vbCr & ", @Comision     =" & Val(txtComision)
'       sqls = sqls & vbCr & ", @CtaCont     = '" & IIf(Trim(txtCtaCont) = "", " ", Trim(txtCtaCont)) & "'"
'       sqls = sqls & vbCr & ", @Cveestablecimiento    =    " & IIf(txtCveEstab.Text <> "", txtCveEstab, 0)
'       sqls = sqls & vbCr & ", @Giro    =    " & IIf(cboGiros.Text <> "", cboGiros.ItemData(cboGiros.ListIndex), 0)
'       sqls = sqls & vbCr & ", @Producto    = " & Product
'       cnxbdMty.Execute sqls, intRegistros
'    End If
    
    If NvoCom = True Then
      If Product = 2 Then
         sqls = "update folios set consecutivo = " & Val(txtGrupo) & _
             " where bodega = 0 and tipo = 'CAT' and prefijo = 'GRG'"
      ElseIf Product = 1 Then
         sqls = "update folios set consecutivo = " & Val(txtGrupo) & _
             " where bodega = 0 and tipo = 'CAT' and prefijo = 'GRU'"
      End If
      cnxBD.Execute sqls, intRegistros
    End If
     
    Screen.MousePointer = 1
    
    If txtGrupo.Enabled = True Then
       txtGrupo.Enabled = False
       MsgBox "El nuevo grupo se ha guardado", vbInformation, "Grupo guardado"
    Else
       MsgBox "Los datos del grupo se han actualizado!", vbInformation, "Datos actualizados"
    End If
    
    cmdBuscarC.Enabled = True
    NvoCom = False
    
    Exit Sub
      
err_gral:
      MsgBox "Error al querer guardar el grupo", vbCritical
      Screen.MousePointer = 1
      Exit Sub
End Sub

Private Sub cmdBorrar_Click()
   
If Val(txtGrupo) = 0 Then
   MsgBox "Primero debe seleccionar el grupo que desea borrar", vbInformation, "Grupo..."
   Exit Sub
End If

If MsgBox("Esta seguro de que desea borrar el grupo " & Val(txtGrupo) & " y sus comercios", vbYesNo) = vbYes Then
   
   'Borra el grupo
   'prod = IIf(Product = 8, 6, Product)
   producto_cual
   sqls = "DELETE GRUPOS WHERE GRUPO = " & Val(txtGrupo) & " and Producto=" & Product
   cnxbdMty.Execute sqls, intRegistros
   
   'Borra los comercios ligados a este grupo
   
   sqls = "DELETE comercios WHERE GRUPO = " & Val(txtGrupo) & " and Producto=" & Product
   cnxBD.Execute sqls, intRegistros
   
   txtGrupo = 1
   
   MsgBox " Grupo borrado !", vbInformation, "Grupo borrado"
   
End If
End Sub

Private Sub cmdBorrarC_Click()
Dim comercio As Long
On Error GoTo ERR:
With spdComercios
   .Col = 1
   .Row = .ActiveRow
   comercio = .Text
   If (MsgBox("Desea eliminar el comercio # " & .Text, vbYesNo + vbDefaultButton2 + vbQuestion)) = vbYes Then
      sqls = " delete comercios where grupo = " & txtGrupo & ""
      sqls = sqls & " and comercio  = '" & Format(Trim(comercio), "0000000") & "'"
      sqls = sqls & " and Producto=" & Product
      
      cnxBD.Execute sqls, intRegistros
      CargaComercios
      MsgBox "Comercio borrado...", vbInformation
   End If
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores encontrados"
  Exit Sub
End Sub


Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "Grupo"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente > 0 Then
       txtGrupo = frmConsulta.cliente
       txtDescripcion = frmConsulta.Nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Private Sub cmdBuscarComercio_Click()
Dim resp As String

resp = InputBox("Numero de Comercio que desea buscar", "Busqueda por comercio")

If resp <> "" Then
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "select grupo from comercios where comercio = '" & Format(resp, "0000000") & "'"
    sqls = sqls & " and Producto=" & Product
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        txtGrupo.Text = rsBD!Grupo
    Else
        MsgBox "El numero de comercio no existe", vbCritical, "Comercio incorrecto"
    End If
End If
End Sub


Private Sub cmdCancelar_Click()
   txtGrupo = ""
   NvoCom = False
   cmdBuscarC.Enabled = True
   MsgBox "Alta de Grupo Cancelada!!", vbInformation
   
End Sub

Private Sub cmdGuardarC_Click()
On Error GoTo ERR:
With spdComercios

   .Row = .ActiveRow
   .Col = 3
   tipocom = .Text
   If Trim(tipocom) = "" Then
      MsgBox " No se puede grabar, ya que no ha seleccionado que tipo de Comercio es", vbCritical, "Seleccione tipo de comercio"
      Exit Sub
   End If
      
   .Col = 5
   estado = .Text
   
   If Trim(estado) = "" Then
      MsgBox " No se puede grabar, ya que no ha capturado el Estado", vbCritical, "Falta estado..."
      Exit Sub
   End If
   
   .Col = 6
   poblacion = .Text
   If Trim(poblacion) = "" Then
      MsgBox " No se puede grabar, ya que no ha capturado la Poblacion", vbCritical, "Falta poblacion..."
      Exit Sub
   End If
   
   estado = Mid(estado, InStr(Trim(estado), "(") + 1, Len(Trim(estado)) - InStr(Trim(estado), "(") - 1)
   poblacion = Mid(poblacion, InStr(Trim(poblacion), "(") + 1, Len(Trim(poblacion)) - InStr(Trim(poblacion), "(") - 1)
   
   .Col = 1
   comercio = .Text
   .Col = 2
   Nombre = Trim(.Text)
   .Col = 3
   tipocomercio = .TypeComboBoxCurSel
   .Col = 4
   Direccion = IIf(.Text = "", " ", .Text)
   .Col = 7
   Telefono = IIf(.Text = "", " ", .Text)
   .Col = 8
   contacto = IIf(.Text = "", " ", .Text)
   
   'prod = IIf(Product = 8, 6, Product)
   producto_cual
    sqls = " EXEC sp_Comercios_ins_upd "
    sqls = sqls & vbCr & "  @Grupo       = " & txtGrupo
    sqls = sqls & vbCr & ", @comercio      = '" & Format(Trim(comercio), "0000000") & "'"
    sqls = sqls & vbCr & " , @Descripcion  = '" & Nombre & "'"
    sqls = sqls & vbCr & ", @TipoComercio  = " & tipocomercio
    sqls = sqls & vbCr & ", @Telefono    = '" & Telefono & "'"
    sqls = sqls & vbCr & ", @Direccion   = '" & Direccion & "'"
    sqls = sqls & vbCr & ", @Contacto   = '" & contacto & "'"
    sqls = sqls & vbCr & ", @Estado   = " & estado & ""
    sqls = sqls & vbCr & ", @Poblacion   = " & poblacion & ""
    sqls = sqls & vbCr & ", @Producto   = " & Product & ""
    cnxbdMty.Execute sqls, intRegistros
    MsgBox "Datos Actualizados", vbInformation, "Comercios actualizados"
   
  ' NvoCom = False
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores presentados"
  Exit Sub
End Sub



Private Sub cmdNuevo_Click()
    LimpiarControles Me
    lblGrupo = ""
    CargaBodegas cboBodegas
    CargaPoblaciones cboPoblaciones, 1
    CargaEstados cboEstados
    With spdComercios
      .Col = -1
      .Row = -1
      .Action = 12
      .MaxRows = 0
    End With
    
    cmdBuscarC.Enabled = False
    NvoCom = True
    BuscaFolioGrupo
    
End Sub
Sub BuscaFolioGrupo()

sqls = " SELECT MAX(GRUPO)consecutivo FROM GRUPOS " & _
      " WHERE PRODUCTO=" & Product & " AND GRUPO<4000 "
 
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
   txtGrupo = rsBD!consecutivo + 1
Else
   MsgBox "Error al asignar el numero de Grupo", vbCritical, "Error en numero de Grupo"
End If
Exit Sub
End Sub
Private Sub cmdNuevoC_Click()
With spdComercios
   
   .MaxRows = spdComercios.MaxRows + 1
   NvoCom = True
   .Col = 1
   .Row = spdComercios.MaxRows
   .Lock = False
     
End With
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelarC_Click()
   frmEstados.Visible = False
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    LimpiarControles Me
    CargaBodegas cboBodegas
    CargaPoblaciones cboPoblaciones, 1
    CargaEstados cboEstados
    CargaEstados cboEstadosC
    CargaGiros cboGiros
     
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    With spdComercios
      .Col = -1
      .Row = -1
      .Action = 12
      .MaxRows = 0
    End With
    txtGrupo.Enabled = False
    NvoCom = False
    frmEstados.Visible = False
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
  'If Product <> 7 Then
     chk1.Enabled = True
  'Else
  '   chk1.Enabled = False
  'End If
End Sub

Sub InicializaForma()
     LimpiarControles Me
     With spdComercios
      .Col = -1
      .Row = -1
      .Action = 12
      .MaxRows = 0
    End With
    lblGrupo = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub spdComercios_DblClick(ByVal Col As Long, ByVal Row As Long)
   If Col = 5 Or Col = 6 Then
      cboEstadosC.ListIndex = -1
      cboPoblacionesC.ListIndex = -1
      frmEstados.Visible = True
   End If
End Sub

Private Sub spdComercios_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim strUpd As String
'
'
'With spdComercios
' If NvoCom = False Then
'   .Col = Col
'   .Row = Row
'   If .Col >= 2 Then
'      strUpd = ""
'      SQLS = "update comercios set "
'
'      Select Case Col
'         Case 2
'            SQLS = SQLS & " Descripcion = '" & UCase(.Text) & "'"
'         Case 3
'            SQLS = SQLS & " TipoComercio = " & .TypeComboBoxCurSel
'         Case 4
'            SQLS = SQLS & " Telefono = '" & .Text & "'"
'         Case 5
'            SQLS = SQLS & " direccion = '" & UCase(.Text) & "'"
'         Case 6
'            SQLS = SQLS & " Contacto = '" & UCase(.Text) & "'"
'      End Select
'
'      .Col = 1
'
'      SQLS = SQLS & " where grupo = " & txtGrupo & ""
'      SQLS = SQLS & " and  comercio = " & .Text
'      cnxBD.Execute SQLS, intRegistros
'   End If
'
' End If
'
'End With
End Sub

Private Sub TxtComision_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 45 And KeyAscii < 58) And KeyAscii > 31 Then
      KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then
        entertab KeyAscii
    End If

End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
 entertab (KeyAscii)
End Sub

Private Sub txtNumCliente_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58) And KeyAscii > 31 Then
      KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then
        entertab KeyAscii
    End If
End Sub
Sub CargaDatosGrupo()
 'prod = IIf(Product = 8, 6, Product)
 producto_cual
 sqls = " select * " & _
        " from Grupos" & _
        " where grupo = " & Val(txtGrupo) & " AND Producto=" & Product
 Set rsGpo = New ADODB.Recordset
 rsGpo.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

If Not rsGpo.EOF Then
   txtDescripcion.Text = rsGpo!descripcion
   TxtDomicilio = IIf(IsNull(rsGpo!Domicilio), "", rsGpo!Domicilio)
   txtCol = IIf(IsNull(rsGpo!Colonia), "", rsGpo!Colonia)
   TxtRFC = IIf(IsNull(rsGpo!Rfc), "", rsGpo!Rfc)
   txtContacto = IIf(IsNull(rsGpo!contacto), "", rsGpo!contacto)
   txtMail = IIf(IsNull(rsGpo!Email), "", rsGpo!Email)
   TxtTel = IIf(IsNull(rsGpo!Telefono), "", rsGpo!Telefono)
   TxtComision = Format(IIf(IsNull(rsGpo!Comision), 0, rsGpo!Comision), "0.00")
   txtCtaCont = IIf(IsNull(rsGpo!ctacont), "", rsGpo!ctacont)
   txtCveEstab = IIf(IsNull(rsGpo!CveEstablecimiento), "", rsGpo!CveEstablecimiento)
   
   CargaPoblaciones cboPoblaciones, IIf(IsNull(rsGpo!estado), 0, rsGpo!estado)
   Call CboPosiciona(cboEstados, IIf(IsNull(rsGpo!estado), -1, rsGpo!estado))
   Call CboPosiciona(cboPoblaciones, IIf(IsNull(rsGpo!poblacion), -1, rsGpo!poblacion))
   Call CboPosiciona(cboBodegas, IIf(IsNull(rsGpo!Bodega), -1, rsGpo!Bodega))
   Call CboPosiciona(cboGiros, IIf(IsNull(rsGpo!Giro), -1, rsGpo!Giro))
   
Else
   MsgBox "Error al cargar los datos del Grupo.", vbCritical, "Error en grupo"
End If

rsGpo.Close
Set rsGpo = Nothing
   
End Sub

Sub CargaComercios()
On Error GoTo ERR:
With spdComercios
 'prod = IIf(Product = 8, 6, Product)
 producto_cual
   sqls = "sp_Consultas_BE Null,Null," & Product & "," & Val(txtGrupo) & ",Null,Null,Null,'CargaComercio'"
'   sqls = " select a.* , b.descripcion descPoblacion , c.descripcion descEstado" & _
'          " from comercios a, poblaciones  b, estados c" & _
'          " Where grupo = " & Val(txtGrupo) & _
'          " and a.poblacion *= b.poblacion" & _
'          " and a.estado  *=c.estado" & _
'          " and a.producto=" & CStr(prod)
          
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
   i = 0
   .Col = -1
   .Row = -1
   .Action = 12
   .MaxRows = 0
   
   Do While Not rsBD.EOF
      i = i + 1
      .MaxRows = i
      .Row = i
      .Col = 1
      .Text = rsBD!comercio
      .Col = 2
      .Text = rsBD!descripcion
      .Col = 3
      .TypeComboBoxCurSel = rsBD!tipocomercio
      .Col = 4
      .Text = IIf(IsNull(rsBD!Direccion), " ", rsBD!Direccion)
      .Col = 5
      If IsNull(rsBD!descestado) Then
         .Text = " "
      Else
         .Text = IIf(IsNull(rsBD!descestado), " ", Trim(rsBD!descestado)) & " (" & IIf(IsNull(rsBD!estado), " ", Trim(rsBD!estado)) & ")"
      End If
      .Col = 6
      If IsNull(rsBD!descpoblacion) Then
         .Text = " "
      Else
         .Text = IIf(IsNull(rsBD!descpoblacion), " ", Trim(rsBD!descpoblacion)) & " (" & IIf(IsNull(rsBD!poblacion), " ", Trim(rsBD!poblacion)) & ")"
      End If
      
      .Col = 7
      .Text = IIf(IsNull(rsBD!Telefono), " ", rsBD!Telefono)
      .Col = 8
      .Text = IIf(IsNull(rsBD!contacto), " ", rsBD!contacto)
     
      rsBD.MoveNext
   Loop
   
End With
rsBD.Close
Set rsBD = Nothing
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Error"
  Exit Sub
End Sub
 
Private Sub txtDescripcion_LostFocus()
   txtDescripcion = UCase(txtDescripcion)
End Sub

Private Sub txtGrupo_Change()
If txtGrupo.Enabled = False Then
   If txtGrupo <> "" And NvoCom = False Then
      CargaDatosGrupo
      CargaComercios
   End If
End If
End Sub

Sub verifica_RFC()
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "SELECT * FROM GRUPOS WHERE PRODUCTO=" & Product
    sqls = sqls & " AND RFC ='" & damerfc(Trim(TxtRFC)) & "'"
    sqls = sqls & " AND GRUPO<>" & Val(txtGrupo.Text)
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       MsgBox "El Grupo No. " & rsBD!Grupo & " tiene el RFC que acaba de teclear", vbCritical, "Ya existe el RFC"
       'txtRFC.SetFocus
       Exit Sub
    End If
End Sub

Function damerfc(nrfc As String) As String
Dim r As Byte, mientras As String
damerfc = ""
For r = 1 To Len(nrfc)
    If Mid(nrfc, r, 1) <> " " And Mid(nrfc, r, 1) <> "," And Mid(nrfc, r, 1) <> "-" And Mid(nrfc, r, 1) <> "." And Mid(nrfc, r, 1) <> "_" Then
       damerfc = damerfc & Mid(nrfc, r, 1)
    End If
Next
mientras = damerfc
If Mid(mientras, 4, 1) = "0" Or Mid(mientras, 4, 1) = "1" Or Mid(mientras, 4, 1) = "2" Or Mid(mientras, 4, 1) = "3" _
   Or Mid(mientras, 4, 1) = "4" Or Mid(mientras, 4, 1) = "5" Or Mid(mientras, 4, 1) = "6" Or Mid(mientras, 4, 1) = "7" Or Mid(mientras, 4, 1) = "8" Or Mid(mientras, 4, 1) = "9" Then
   mientras = Mid(mientras, 1, 3) & " " & Mid(mientras, 4)
   damerfc = mientras
End If
End Function

Private Sub TxtRFC_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And TxtRFC <> "" Then
     verifica_RFC
  End If
End Sub

Private Sub txtRFC_LostFocus()
 verifica_RFC
End Sub

Sub Actualiza_gpos()
Dim sqls As String, cta_con As String
On Error GoTo ERR:
sqls = "SELECT * FROM GRUPOS"
sqls = sqls & " WHERE Descripcion='" & Trim(txtDescripcion.Text) & "'"
sqls = sqls & " AND PRODUCTO<>2"
sqls = sqls & " ORDER BY Producto"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   Do While Not rsBD.EOF
      cta_con = IIf(IsNull(rsBD!ctacont), "", Trim(rsBD!ctacont))
      sqls = " EXEC sp_InsUpd_Grupos "
      sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
      sqls = sqls & vbCr & ", @Grupo      = " & rsBD!Grupo
      sqls = sqls & vbCr & ", @Descripcion  =    '" & IIf(Trim(txtDescripcion) = "", " ", Trim(txtDescripcion)) & "'"
      sqls = sqls & vbCr & ", @Domicilio    =    '" & IIf(Trim(TxtDomicilio) = "", " ", Trim(TxtDomicilio)) & "'"
      sqls = sqls & vbCr & ", @Colonia      =    '" & IIf(Trim(txtCol) = "", " ", Trim(txtCol)) & "'"
      sqls = sqls & vbCr & ", @Poblacion    =    " & IIf(cboPoblaciones.Text <> "", cboPoblaciones.ItemData(cboPoblaciones.ListIndex), 0)
      sqls = sqls & vbCr & ", @Estado    =    " & IIf(cboEstados.Text <> "", cboEstados.ItemData(cboEstados.ListIndex), 0)
      sqls = sqls & vbCr & ", @rfc    =  '" & IIf(Trim(damerfc(TxtRFC)) = "", " ", Trim(damerfc(TxtRFC))) & "'"
      sqls = sqls & vbCr & ", @Contacto       = '" & IIf(Trim(txtContacto) = "", " ", Trim(txtContacto)) & "'"
      sqls = sqls & vbCr & ", @Telefono = '" & IIf(Trim(TxtTel) = "", " ", Trim(TxtTel)) & "'"
      sqls = sqls & vbCr & ", @email     ='" & IIf(Trim(txtMail) = "", " ", Trim(txtMail)) & "'"
      sqls = sqls & vbCr & ", @Comision     =" & Val(rsBD!Comision)
      sqls = sqls & vbCr & ", @CtaCont     = '" & IIf(Trim(txtCtaCont) = "", " ", Trim(txtCtaCont)) & "'"
      'Trim(cta_con)
      sqls = sqls & vbCr & ", @Cveestablecimiento    =    " & IIf(txtCveEstab.Text <> "", txtCveEstab, 0)
      sqls = sqls & vbCr & ", @Giro    =    " & IIf(cboGiros.Text <> "", cboGiros.ItemData(cboGiros.ListIndex), 0)
      sqls = sqls & vbCr & ", @Producto    = " & rsBD!Producto
      cnxbdMty.Execute sqls, intRegistros
      rsBD.MoveNext
   Loop
Else
    sqls = " EXEC sp_InsUpd_Grupos "
    sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    sqls = sqls & vbCr & ", @Grupo      = " & txtGrupo
    sqls = sqls & vbCr & ", @Descripcion  =    '" & IIf(Trim(txtDescripcion) = "", " ", Trim(txtDescripcion)) & "'"
    sqls = sqls & vbCr & ", @Domicilio    =    '" & IIf(Trim(TxtDomicilio) = "", " ", Trim(TxtDomicilio)) & "'"
    sqls = sqls & vbCr & ", @Colonia      =    '" & IIf(Trim(txtCol) = "", " ", Trim(txtCol)) & "'"
    sqls = sqls & vbCr & ", @Poblacion    =    " & IIf(cboPoblaciones.Text <> "", cboPoblaciones.ItemData(cboPoblaciones.ListIndex), 0)
    sqls = sqls & vbCr & ", @Estado    =    " & IIf(cboEstados.Text <> "", cboEstados.ItemData(cboEstados.ListIndex), 0)
    sqls = sqls & vbCr & ", @rfc    =  '" & IIf(Trim(damerfc(TxtRFC)) = "", " ", Trim(damerfc(TxtRFC))) & "'"
    sqls = sqls & vbCr & ", @Contacto       = '" & IIf(Trim(txtContacto) = "", " ", Trim(txtContacto)) & "'"
    sqls = sqls & vbCr & ", @Telefono = '" & IIf(Trim(TxtTel) = "", " ", Trim(TxtTel)) & "'"
    sqls = sqls & vbCr & ", @email     ='" & IIf(Trim(txtMail) = "", " ", Trim(txtMail)) & "'"
    sqls = sqls & vbCr & ", @Comision     =" & Val(TxtComision)
    sqls = sqls & vbCr & ", @CtaCont     = '" & IIf(Trim(txtCtaCont) = "", " ", Trim(txtCtaCont)) & "'"
    sqls = sqls & vbCr & ", @Cveestablecimiento    =    " & IIf(txtCveEstab.Text <> "", txtCveEstab, 0)
    sqls = sqls & vbCr & ", @Giro    =    " & IIf(cboGiros.Text <> "", cboGiros.ItemData(cboGiros.ListIndex), 0)
    sqls = sqls & vbCr & ", @Producto    = " & Product
    cnxbdMty.Execute sqls, intRegistros
End If
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub


