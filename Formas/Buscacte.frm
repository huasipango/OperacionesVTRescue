VERSION 5.00
Begin VB.Form frmBusca_Cliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Cliente"
   ClientHeight    =   4455
   ClientLeft      =   1575
   ClientTop       =   1875
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Picture         =   "Buscacte.frx":0000
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Picture         =   "Buscacte.frx":02E2
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   450
      Left            =   4560
      Picture         =   "Buscacte.frx":05C4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.ListBox lstClientes 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   6315
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   4410
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   6300
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   6300
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Resultados:"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   825
      Width           =   1455
   End
   Begin VB.Label lblTipo 
      Caption         =   "Nombre del Cliente:"
      Height          =   210
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmBusca_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mbytBodega  As Long
Dim mintCliente As Long
Dim mstrNombre  As String
Dim xMas        As Boolean
Dim i As Long


Property Let Bodega(vdata As Long)
    mbytBodega = vdata
End Property
Property Let Mas(vdata As Boolean)
    xMas = vdata
End Property

Property Let cliente(vdata As Long)
    mintCliente = vdata
End Property

Sub BuscaDatos()
On Error GoTo ERRO:
    Dim i As Long
    Dim Blancos As String
    Blancos = "                                        "
    MsgBar "Leyendo Datos", True
    If TipoBusqueda = "Cliente" Then
       sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda
    End If
    If TipoBusqueda = "ClienteBE" Then
       sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "','" & TipoBusqueda & "'"
    End If
    If TipoBusqueda = "Grupo" Then
       sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda & "," & Product
    End If
    If TipoBusqueda = "Emisores" Then
       sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda
    End If
    If TipoBusqueda = "Comercios" Then
      sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda
    End If
    If TipoBusqueda = "ClientesOI" Then
      sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda
    End If
    If TipoBusqueda = "Establecimientos" Then
      sqls = " EXEC sp_BuscaCliente_Datos '" & txtNombre.Text & "'," & TipoBusqueda
    End If
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
    If rsBD.EOF Then
       MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
       txtNombre.SetFocus
       rsBD.Close
       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
    End If
    lstClientes.Clear
    i = 0
    Do While rsBD.EOF = False
      If TipoBusqueda = "Grupo" Then
         Bodegp = Val(rsBD!c)
      End If
      lstClientes.AddItem Mid(Trim("" & rsBD!b) + Blancos, 1, 40) + Mid(Trim(rsBD!a) + Blancos, 1, 9)
      lstClientes.ItemData(i) = Val("" & rsBD!a)
      i = i + 1
      rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    MsgBar "", False
    Exit Sub
ERRO:
   MsgBox " Los datos capturados NO SON VALIDOS O NO EXISTE INFORMACION AL RESPECTO!", vbCritical, "Verifique y corrija"
   txtNombre.Text = ""
'   rsBD.Close
   Set rsBD = Nothing
   MsgBar "", False
   Exit Sub
End Sub

Property Let nombre(vdata2 As String)
    mstrNombre = vdata2
End Property

Property Get cliente() As Long
    cliente = mintCliente
End Property

Property Get nombre() As String
    nombre = mstrNombre
End Property

Sub Seleccion()
    For i = 0 To lstClientes.ListCount - 1
        If lstClientes.Selected(i) Then
           mintCliente = lstClientes.ItemData(i)
           mstrNombre = Mid(lstClientes.List(i), 1, 40)
        End If
    Next i
    cliente_busca = mstrNombre
End Sub

Private Sub cmdAceptar_Click()
    Seleccion
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
'If Len(txtNombre.Text) > 0 Then
    BuscaDatos
'Else
    'MsgBox "Debe introducir algun dato, para obtener resultados", vbExclamation, "Agregue datos"
    'txtNombre.SetFocus
'End If
Exit Sub
End Sub


Private Sub cmdregresar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForma(Me)
    Select Case TipoBusqueda
      Case "Cliente"
            Me.Caption = "Buscar Cliente"
            lblTipo.Caption = "Nombre del Cliente"
      Case "Grupo"
            Me.Caption = "Buscar Grupo"
            lblTipo.Caption = "Nombre del Grupo"
      Case "Establecimiento"
            Me.Caption = "Buscar Establecimiento"
            lblTipo.Caption = "Nombre del Establecimiento"
      End Select
     cmdAceptar.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBar "", False
End Sub

Private Sub lstClientes_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub lstclientes_DblClick()
    Seleccion
    Unload Me
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtNombre.Text <> "" Then
       BuscaDatos
    Else
       txtNombre.SetFocus
    End If
Exit Sub
End Sub
