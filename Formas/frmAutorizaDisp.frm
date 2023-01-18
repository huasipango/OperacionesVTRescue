VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmAutorizaDisp 
   Caption         =   " "
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Tarjetas Pendientes de Autorizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11175
      Begin FPSpread.vaSpread spdTarjetas 
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "frmAutorizaDisp.frx":0000
         TabIndex        =   10
         Top             =   360
         Width           =   10815
      End
   End
   Begin VB.Frame Frame7 
      Height          =   675
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   315
         Left            =   3360
         Picture         =   "frmAutorizaDisp.frx":11DC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1800
         Picture         =   "frmAutorizaDisp.frx":12DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         Picture         =   "frmAutorizaDisp.frx":13E0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   120
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
         ItemData        =   "frmAutorizaDisp.frx":14E2
         Left            =   1800
         List            =   "frmAutorizaDisp.frx":14EC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label11 
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
         TabIndex        =   4
         Top             =   310
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedidos de Dispersiones Pendientes de Autorizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   11175
      Begin FPSpread.vaSpread sprDatos 
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "frmAutorizaDisp.frx":14FE
         TabIndex        =   2
         Top             =   360
         Width           =   10815
      End
   End
End
Attribute VB_Name = "frmAutorizaDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim inicia As Boolean
Dim prod As Byte, SI As Boolean

Private Sub cboProducto_Click()
Dim aqui As Byte, yano As Boolean
  yano = False
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
      CargaDatos
      yano = True
     'InicializaForma
  End If
  If inicia = False And yano = False Then
     cmdBuscarC_Click
  End If
End Sub


Private Sub cmdBuscarC_Click()
    CargaDatos
End Sub

Private Sub cmdGrabar_Click()
 Dim Sucursal As Integer, Pedido As Long
 Dim foli As Long
 Screen.MousePointer = 11
 SI = False
 Call checa_interruptor
 If SI = False Then
 With spdtarjetas
    For i = 1 To .MaxRows
        .Row = i
        .Col = 9
        If .value = True Then
            .Col = 1
            Sucursal = .Text
            .Col = 2
            Pedido = Mid(.Text, 3, Len(.Text) - 3)
            sqls = "sp_Consultas_BE Null,Null," & Product & "," & Sucursal & ",Null," & Pedido & ",Null,'Autoriza_Tarjet'"
            cnxBD.Execute sqls
        End If
    Next i
 End With
 With sprDatos
    For i = 1 To .MaxRows
        .Row = i
        .Col = 9
        If .value = True Then
            .Col = 1
            Sucursal = .Text
            .Col = 2
            Pedido = Mid(.Text, 3, Len(.Text) - 3)
            sqls = "sp_Consultas_BE Null,Null," & Product & "," & Sucursal & ",Null," & Pedido & ",Null,'Autoriza_pedido'"
            cnxBD.Execute sqls
        End If
    Next i
 End With
 Screen.MousePointer = 1
 MsgBox "Datos Actualizados!", vbInformation, "Informacion actualizada"
 CargaDatos
 End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
 'Call checa_interruptor
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Sub checa_interruptor()
    Dim sqls As String
    sqls = "SELECT * FROM Claves Where Tabla='SistemaFBE'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If rsBD.EOF Then
       MsgBox "El interruptor ha sido cambiado de su posicion original", vbCritical, "Interruptor off-line"
       Exit Sub
    Else
       If rsBD!Status = 1 And gnBodega <> 1 Then
          MsgBox "Lo siento ... la opcion ha sido bloqueada temporalmente debido a que se esta realizando un cierre en este momento", vbCritical, "Opcion bloqueada"
          SI = True
          Screen.MousePointer = 1
          Unload Me
       End If
    End If
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    inicia = True
    CboProducto.Clear
    inicia = False
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart") 'aqui lo omiti
    CargaDatos
End Sub

Sub CargaDatos()
Dim Status As Integer

Status = 0
 

' Busca pedidos de tarjetas pendientes de autorizar

sqls = "exec spr_FactAutorizaTarj " & Product

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
With spdtarjetas
.Col = -1
.Row = -1
.Action = 12
.MaxRows = 0
i = 0
Do While Not rsBD.EOF
    i = i + 1
    .MaxRows = i
    .Row = i
    .Col = 1
    .Text = Val(rsBD!Bodega)
    .Col = 2
    .Text = rsBD!Pedido
    .Col = 3
    .Text = Format(rsBD!fechaped, "mm/dd/yy")
    .Col = 4
    .Text = Val(rsBD!cliente)
    .Col = 5
    .Text = rsBD!Nombre
    .Col = 6
    .Text = CDbl(rsBD!valor)
    .Col = 7
    .Text = CDbl(rsBD!pago)
    .Col = 8
    .Text = Format(rsBD!fechapago, "mm/dd/yy")
    .Col = 9
    If rsBD!Status = 10 Then
        .value = 1
    Else
        .value = 0
    End If
    
    rsBD.MoveNext
Loop
End With
rsBD.Close
Set rsBD = Nothing

' Busca pedidos de dispersion pendientes de autorizar
sqls = "exec spr_FactAutorizaDisp " & Product

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
With sprDatos
.Col = -1
.Row = -1
.Action = 12
.MaxRows = 0
i = 0
Do While Not rsBD.EOF
    i = i + 1
    .MaxRows = i
    .Row = i
    .Col = 1
    .Text = Val(rsBD!Bodega)
    .Col = 2
    .Text = rsBD!Pedido
    .Col = 3
    .Text = Format(rsBD!fechaped, "mm/dd/yy")
    .Col = 4
    .Text = Val(rsBD!cliente)
    .Col = 5
    .Text = rsBD!Nombre
    .Col = 6
    .Text = CDbl(rsBD!valor)
    .Col = 7
    .Text = CDbl(rsBD!pago)
    .Col = 8
    .Text = Format(rsBD!fechapago, "mm/dd/yy")
    .Col = 9
    If rsBD!Status = 10 Then
        .value = 1
    Else
        .value = 0
    End If
    
    rsBD.MoveNext
Loop
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

