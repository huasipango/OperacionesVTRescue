VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmMovEmpleados 
   Caption         =   "Consulta de Saldos"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "frmMovEmpleados"
   ScaleHeight     =   8145
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spdBusca 
      Height          =   1935
      Left            =   1200
      OleObjectBlob   =   "frmMovEmpleados.frx":0000
      TabIndex        =   19
      Top             =   3120
      Width           =   9495
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   36
      Top             =   360
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
         ItemData        =   "frmMovEmpleados.frx":02DE
         Left            =   1800
         List            =   "frmMovEmpleados.frx":02E8
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
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
         TabIndex        =   37
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame pandisp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispersiones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3420
      TabIndex        =   34
      Top             =   2805
      Visible         =   0   'False
      Width           =   5535
      Begin FPSpread.vaSpread spddisp 
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "frmMovEmpleados.frx":02FA
         TabIndex        =   35
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   12015
      Begin VB.TextBox txtfechamin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPresentar 
         Height          =   450
         Left            =   10680
         Picture         =   "frmMovEmpleados.frx":05EA
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "Det"
         ToolTipText     =   "Imprime en la Pantalla"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   11280
         Picture         =   "frmMovEmpleados.frx":06EC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Desde el dia[dd/mm/aaaa]:"
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
         Left            =   6600
         TabIndex        =   32
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   12015
      Begin FPSpread.vaSpread spdMov 
         Height          =   3615
         Left            =   120
         OleObjectBlob   =   "frmMovEmpleados.frx":085E
         TabIndex        =   18
         Top             =   360
         Width           =   11775
      End
      Begin FPSpread.vaSpread spdtarjetas 
         Height          =   1695
         Left            =   1200
         OleObjectBlob   =   "frmMovEmpleados.frx":34F8
         TabIndex        =   33
         Top             =   2160
         Visible         =   0   'False
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12015
      Begin VB.CommandButton cmddisp 
         Caption         =   "Disp X Lib"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7320
         TabIndex        =   30
         Top             =   1200
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdtarj 
         Caption         =   "Tarjetas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8520
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtFechacorte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10920
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "Calcular Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9720
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtTipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   15
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtEstado 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtEmpleado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtTarjeta2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txttarjeta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "58877265"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Corte al día:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Estatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Empleadora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Numero Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Numero Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Numero Tarjeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMovEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim fech1, fech2 As String
Dim prod As Byte

Private Sub cmdBorrar_Click()
LimpiarControles Me
If Product = 1 Then
   txttarjeta.Text = "50640601"
ElseIf Product = 2 Then
  txttarjeta.Text = "50640501"
End If

With spdMov
    .Col = -1
    .Row = -1
    .Action = 12
    .MaxRows = 0
End With
Call BuscaFechaCorte(1)
cmddisp.Enabled = False
cmdtarj.Enabled = False
txtEmpleado.SetFocus
spdtarjetas.Visible = False
spdBusca.Visible = False
pandisp.Visible = False
Exit Sub
End Sub

Private Sub cmdCalcular_Click()

Dim SaldoIni As Double
Dim FechaCorte As Date
    
    Screen.MousePointer = 11

    If Val(TxtCuenta) <> 0 Then
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         
         sqls = "sp_Consultas_BE NULL,Null," & Product & "," & TxtCuenta & ",Null,Null,Null,'SaldoFinal'"
'         sqls = "select importe , fechaCorte  from saldosfinalesbe" & _
'                " where  cuenta = " & txtcuenta & _
'                " and Producto=" & prod & _
'                " order by fechacorte desc"
                
         Set rsBD = New ADODB.Recordset
         rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
         
         If rsBD.EOF Then
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = "sp_Consultas_BE NULL,Null," & Product & ",Null,Null,Null,Null,'FechacorteSFinal'"
'            sqls = "select max(fechaCorte) Fecha  from saldosfinalesbe"
'            sqls = sqls & " where Producto=" & prod
            
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
               
            SaldoIni = 0
            FechaCorte = rsBD!Fecha
         Else
            FechaCorte = rsBD!FechaCorte
            SaldoIni = rsBD!importe
         End If
         rsBD.Close
         Set rsBD = Nothing
         
    Else
        MsgBox "Debe capturar la cuenta que desea consultar", vbInformation, "Consulta de Saldos"
        Screen.MousePointer = 1
        Exit Sub
    End If
    
   '------------------------------------COMPRAS
     'prod = IIf(Product = 8, 6, Product)
     producto_cual
    
     sqls = "sp_Consultas_BE '" & Format(FechaCorte, "MM/DD/YYYY") & "',Null," & Product & "," & Val(TxtCuenta) & ",Null,Null,Null,'Compras'"
'    sqls = " select Movimientos=IsNull(Sum(Case  When (Tipomov='C'and status=1 and comercio<>'DISP' AND comercio<>'AJUS') Then Importe" & _
'                                              " Else 0 " & _
'             "End),0)" & _
'           " From liquidacionesbe" & _
'           " Where cuenta = " & Val(txtcuenta) & _
'           " and fechatran > '" & Format(FechaCorte, "MM/DD/YYYY") & " 23:59:00' " & _
'           " and comercio <> '0099999' and cveresp <= 1" & _
'           " and Producto=" & prod

     Set rsBD = New ADODB.Recordset
     rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
     
     transacciones = rsBD!movimientos
     rsBD.Close
     Set rsBD = Nothing
   '--------------------------------------------DISPERSIONES
     'prod = IIf(Product = 8, 6, Product)
     producto_cual
     sqls = "sp_Consultas_BE '" & Format(FechaCorte, "MM/DD/YYYY") & "',Null," & Product & "," & Val(TxtCuenta) & ",Null,Null,Null,'Dispersiones'"
'     sqls = " select Movimientos=IsNull(Sum(Case  When (Tipomov='A') Then Importe ELSE 0" & _
'           " End),0)" & _
'           " From liquidacionesbe" & _
'           " Where cuenta = " & Val(txtcuenta) & _
'           " and fechatran > '" & Format(FechaCorte, "MM/DD/YYYY") & " 23:59:00' " & _
'           " and comercio = 'DISP' and cveresp <= 1" & _
'           " and Producto=" & prod
           
     Set rsBD = New ADODB.Recordset
     rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
     
     dispersiones = rsBD!movimientos
     rsBD.Close
     Set rsBD = Nothing
   '--------------------------------------------AJUSTES
     'prod = IIf(Product = 8, 6, Product)
     producto_cual
     sqls = "sp_Consultas_BE '" & Format(FechaCorte, "MM/DD/YYYY") & "',Null," & Product & "," & Val(TxtCuenta) & ",Null,Null,Null,'Ajustes'"
'     sqls = " select Movimientos=IsNull(Sum(Case  When (Tipomov='C') Then Importe ELSE 0" & _
'           " End),0)" & _
'           " From liquidacionesbe" & _
'           " Where cuenta = " & Val(txtcuenta) & _
'           " and fechatran > '" & Format(FechaCorte, "MM/DD/YYYY") & " 23:59:00' " & _
'           " and comercio = 'AJUS' and cveresp <= 1" & _
'           " and Producto=" & prod
           
     Set rsBD = New ADODB.Recordset
     rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
     
     ajustes = rsBD!movimientos
     rsBD.Close
     Set rsBD = Nothing
   '------------------------------------------VENCIMIENTOS
    'prod = IIf(Product = 8, 6, Product)
    producto_cual
    sqls = "sp_Consultas_BE '" & Format(FechaCorte, "MM/DD/YYYY") & "',Null," & Product & "," & Val(TxtCuenta) & ",Null,Null,Null,'Vencimientos'"
'    sqls = " select Movimientos=IsNull(Sum(Case  When (Tipomov='A') Then Importe ELSE 0" & _
'                                                      " End),0)" & _
'           " From liquidacionesbe" & _
'           " Where cuenta = " & Val(txtcuenta) & _
'           " and fechatran > '" & Format(FechaCorte, "MM/DD/YYYY") & " 23:59:00' " & _
'           " and comercio <> '0099999' and cveresp <= 1" & _
'           " AND STATUS=2" & _
'           " and Producto=" & prod

     Set rsbd1 = New ADODB.Recordset
     rsbd1.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
     
     VENCIMIENTOS = rsbd1!movimientos
     rsbd1.Close
     Set rsbd1 = Nothing
   '------------------------------------------SUMA TOTAL
     txtSaldo = SaldoIni + dispersiones - transacciones - ajustes '+ VENCIMIENTOS
     
     Screen.MousePointer = 1
     'rsFecha.Close
     'Set rsFecha = Nothing
                
End Sub



Private Sub cmddisp_Click()
  cmdtarj.Enabled = False
  pandisp.Visible = True
  spddisp.SetFocus
  Call buscadispxlib(Val(TxtCuenta.Text))
End Sub

Private Sub cmdPresentar_Click()
   Imprime crptToWindow
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdtarj_Click()
  cmddisp.Enabled = False
  spdtarjetas.Top = 360
  spdtarjetas.Visible = True
  spdtarjetas.SetFocus
  Call buscatarjetas(Val(TxtCuenta.Text))
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
    cmdCalcular.Caption = "Calcular" & Chr(13) & " Saldo"
    cmdBorrar.Caption = "Borrar" & Chr(13) & " Datos"
    rango_periodo = ""
    spdBusca.Visible = False
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    Call BuscaFechaCorte(0)
End Sub
Sub BuscaFechaCorte(Opcion As Integer)
Dim lafmin, mm As String
Dim sqlmin As String
Dim elmes As Integer, cercano As String
Dim rsFecha As New ADODB.Recordset

If rsFecha.State = 1 Then
    rsFecha.Close
    Set rsFecha = Nothing
End If
'para obtener fecha maxima o limite superior de fecha
'prod = IIf(Product = 8, 6, Product)
producto_cual
'sqls = "sp_Consultas_BE Null,Null," & prod & ",Null,Null,Null,Null,'FechaCorte'"
sqls = "select max(fechaENVIO) fechacorte from liquidacionesbe where tipomov  = 'C' and status < 2 and cveresp in (0,1)"
sqls = sqls & " and Producto=" & Product
If Product >= 8 Then
   cercano = Format((Date - 10), "mm/dd/yyyy")
   sqls = sqls & " and Fechaenvio>='" & cercano & "'"
End If

Set rsFecha = New ADODB.Recordset
If Opcion = 0 Then
    rsFecha.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly, adCmdText
Else
    rsFecha.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly, adCmdText
    'rsFecha.Open sqls, cnxbdMty2, adOpenForwardOnly, adLockReadOnly, adCmdText
End If

If Not rsFecha.EOF Then
    txtFechacorte = Format(rsFecha!FechaCorte, "DD/MM/YYYY")
    If txtFechacorte <= Date Then
       txtFechacorte = Format(Date, "DD/MM/YYYY")
    End If
    If txtFechacorte <> "" Then
       lafmin = txtFechacorte.Text
    Else
       txtFechacorte = Format(Date, "DD/MM/YYYY")
       lafmin = txtFechacorte.Text
    End If
Else
  MsgBox "No hay transacciones aun para este producto", vbCritical, "Producto sin movimientos aun"
  InicializaForma
  Exit Sub
End If
rsFecha.Close
Set rsFecha = Nothing

mm = Mid(lafmin, 4)
sqlmin = Mid(mm, 1, 2)
elmes = Val(sqlmin)
If (elmes - 1) <= 0 Then
   elmes = 12
   lafmin = "01/" & elmes & "/" & Val(Mid(mm, 4)) - 1
Else
   elmes = Val(sqlmin) - 1
   lafmin = "01/" & elmes & "/" & Mid(mm, 4)
End If
txtfechamin = Format(lafmin, "DD/MM/YYYY")
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEdoCuenta.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & serverc & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & based
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(TxtCuenta.Text)
    mdiMain.cryReport.StoredProcParam(1) = Format(txtfechamin.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(2) = Format(txtFechacorte.Text, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = CStr(Product)
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub spdBusca_DblClick(ByVal Col As Long, ByVal Row As Long)

With spdBusca

 If Row = 0 And Col = 0 Then
    spdBusca.Visible = False
    Exit Sub
 End If

 If Row > 0 Then
    .Row = Row
    .Col = 4
    txtTarjeta2.Text = Right(Trim(.Text), 8)
    txtTarjeta2_LostFocus
    .Visible = False
 ElseIf Row = 0 Then
             spdBusca.Visible = False
             InicializaForma
             Exit Sub
            .SortBy = SS_SORT_BY_ROW
            .SortKey(1) = Col
            If .SortKeyOrder(1) <> SS_SORT_ORDER_ASCENDING Then
                .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            Else
                .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
            End If
            
            .Col = 1
            .Col2 = 7
            .Row = 1
            .Row2 = .MaxRows
            .Action = SS_ACTION_SORT
    
  End If
 
End With
End Sub

Private Sub spddisp_DblClick(ByVal Col As Long, ByVal Row As Long)
  pandisp.Visible = False
  cmdtarj.Enabled = True
End Sub

Private Sub spdtarjetas_DblClick(ByVal Col As Long, ByVal Row As Long)
  spdtarjetas.Visible = False
  cmddisp.Enabled = True
End Sub

Private Sub txtcuenta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(TxtCuenta.Text) <> "" Then
           Call BuscaDatos(Trim(TxtCuenta.Text), "Cuenta")
        End If
    End If
End Sub

Sub BuscaDatos(valor As Variant, tipoBusc As String)
Dim rsDatos As ADODB.Recordset
Dim j As Integer
On Error GoTo ERR:
Screen.MousePointer = 11
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = " select a.nocuenta Cuenta,  Isnull(a.NoEmpleadoNvo,NoEmpleado) NumEmp, ISNULL(a.nombre,'SIN NOMBRE') NombreEmp,"
sqls = sqls & " convert(varchar(16),dbo.DesEncriptar(b.notarjeta)) Tarjeta, convert(varchar(10), c.cliente)  + ' - ' + c.nombre NombreCte"
sqls = sqls & " from cuentasbe a (nolock) , tarjetasbe b (nolock), clientes c"
sqls = sqls & " where a.nocuenta = b.nocuenta"
sqls = sqls & " and a.empleadora = c.cliente and b.tipo = 'T' "
sqls = sqls & " and a.Producto=" & Product
sqls = sqls & " and b.Producto=" & Product

Select Case tipoBusc
    Case "Nombre"
        sqls = sqls & " and a.nombre like '%" & valor & "%'"
        sqls = sqls & " order by a.nombre,c.cliente"
    Case "Cuenta"
        sqls = sqls & " and a.nocuenta =" & valor & ""
        sqls = sqls & " order by a.nocuenta,c.cliente"
    Case "Tarjeta"
        sqls = sqls & " and convert(varchar(16),dbo.DesEncriptar(b.tarjeta))  like '%" & valor & "%'"
        sqls = sqls & " order by b.tarjeta,c.cliente"
    Case "Empleado"
        sqls = sqls & " and (a.noempleado  = '" & valor & "'"
        sqls = sqls & " OR a.NoEmpleadoNvo  = '" & valor & "')"
        sqls = sqls & " order by c.cliente,a.noempleado,a.NoEmpleadoNvo"
End Select

Set rsDatos = New ADODB.Recordset
rsDatos.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly 'AKI ES CNXBDMTY

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
rsDatos.Close
Set rsDatos = Nothing
Exit Sub
ERR:
   MsgBox "Se presento el siguiente error: " & ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub

Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtEmpleado.Text) <> "" Then
           Call BuscaDatos(Trim(txtEmpleado.Text), "Empleado")
        End If
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        If Trim(TxtNombre.Text) <> "" Then
           Call BuscaDatos(Trim(TxtNombre.Text), "Nombre")
        End If
  End If
End Sub

Private Sub txtTarjeta2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtTarjeta2.Text) <> "" Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTarjeta2_LostFocus()
    Screen.MousePointer = 11
    If Trim(txtTarjeta2.Text) <> "" Then
        txtSaldo.Text = ""
        Tarjeta = Trim(txttarjeta.Text) & Trim(txtTarjeta2.Text)
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        sqls = "select nocuenta from tarjetasbe(nolock) where (notarjeta = dbo.Encriptar('" & Tarjeta & "')"
        sqls = sqls & " )"
        sqls = sqls & " and  Producto=" & Product
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly 'conexion cxnxbdmty
        
        If Not rsBD.EOF Then
            Cuenta = rsBD!noCuenta
            sqls = " select a.nocuenta, a.empleadora, a.noempleado, a.nombre,c.nombre nombreEmpresa, a.status"
           
        Else
            MsgBox "No se encuentra la cuenta de esta tarjeta", vbCritical, "Tarjeta no existe"
            Screen.MousePointer = 1
            Exit Sub
        End If
        
        If Len(txtFechacorte.Text) < 10 Or Len(txtfechamin.Text) < 10 Then
           txtfechamin.Text = Date
           txtFechacorte.Text = Date - 30
        Else
           'txtfechamin.Text = corrige_fecha(txtfechamin)
           'txtFechacorte.Text = corrige_fecha(txtFechacorte)
           fech1 = corrige_fecha(txtfechamin)
           fech2 = corrige_fecha(txtFechacorte)
        End If
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        sqls = "exec spr_liquidacionesbe " & Cuenta & ", '" & fech1 & "', '" & fech2 & "'," & Product
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly ''conexion cxnxbdmty
        If Not rsBD.EOF Then
          cmdBorrar_Click
          
            txtEmpleado.Text = rsBD!numempleado
            TxtCuenta.Text = rsBD!noCuenta
            TxtNombre.Text = Trim(rsBD!Nombre)
            txtCliente.Text = rsBD!NomCliente
            txtEstado.Text = IIf(rsBD!statustarjeta = 1, "ACTIVA", "CANCELADA")
            txtTipo.Text = Trim(rsBD!tipo)
            txtTarjeta2.Text = Right(Trim(rsBD!Tarjeta), 8)
           ' txtStatusTarjeta = IIf(rsBD!statustarjeta = 1, "Activa", "Cancelada")
            
            
            rsBD.MoveFirst
            i = 0
            
         With spdMov
                .Col = -1
                .Row = -1
                .Action = 12
                
                Do While Not rsBD.EOF
                    i = i + 1
                    .MaxRows = i
                    .Row = i
                    .Col = 1
                    .Text = Format(rsBD!fechatran, "mm/dd/yy")
                    .Col = 2
                    If rsBD!CveTran = 1 And rsBD!establecimiento = "AJUSTE A CUENTA" Then
                        .Text = "AJUSTE POR FONDOS INSUFICIENTES"
                    Else
                        .Text = IIf(IsNull(rsBD!establecimiento), 0, rsBD!establecimiento)
                    End If
                    .Col = 4
                    fechaaplic = IIf(IsNull(rsBD!fechaaplic), "", rsBD!fechaaplic)
                    .Text = Format(fechaaplic, "mm/dd/yy")
                    .Col = 3
                        If rsBD!Status = 2 Then
                            .Text = "Vencimiento"
                        ElseIf rsBD!Status = 3 Then
                            .Text = "Devolucion a la cuenta"
                        ElseIf rsBD!Status = 4 Then
                            .Text = "Rechazo"
                        
                        Else
                        
                            .Text = IIf(IsNull(rsBD!descrechazo), 0, rsBD!descrechazo)
                        End If
                    .Col = 5
                    .Text = CDbl(rsBD!importe)
                    .Col = 6
                    .Text = CDbl(rsBD!saldo)
                    .Col = 7
                    .Text = rsBD!Tarjeta
                    Select Case rsBD!Status
                        Case 1
                            .Col = 8
                            .value = 1
                        Case 2
                            .Col = 9
                            .value = 1
                        Case 3
                            .Col = 10
                            .value = 1
                    End Select
                
                    rsBD.MoveNext
                    
                Loop
            
            End With
            cmddisp.Enabled = True
            cmdtarj.Enabled = True
        Else
            MsgBox "No hay movimientos para esta tarjeta !", vbCritical
            'cmdBorrar_Click
            spdBusca.Visible = False
            InicializaForma
        
        End If
rsBD.Close
Set rsBD = Nothing
End If
Screen.MousePointer = 1
End Sub

Private Function corrige_fecha(lafecha As String) As String
Dim dia, mes, ano As String
dia = Mid(lafecha, 1, 2)
mes = Mid(lafecha, 4, 2)
ano = Mid(lafecha, 7, 4)

  If Val(dia) > 31 Or Val(dia) < 0 Then
     dia = "01"
  End If
  If Val(mes) > 12 Or Val(mes) < 0 Then
     mes = Mid(Date, 4, 2)
  End If
  lafecha = mes & "/" & dia & "/" & ano
  corrige_fecha = lafecha

End Function

Sub buscatarjetas(ncuent As Double)
Dim rsDatos As ADODB.Recordset
Dim j As Integer, i As Integer
If TxtCuenta <> "" Then
Screen.MousePointer = 11
'prod = IIf(Product = 8, 6, Product)
producto_cual

sqls = "sp_Consultas_BE Null,Null," & Product & "," & ncuent & ",Null,Null,Null,'BuscaTarjetas'"
'sqls = " select NoTarjeta,ISNULL(Nombre,'SIN NOMBRE') Nombre,FechaModificacion,FechaCancelacion,status,Tipo"
'sqls = sqls & " from tarjetasbe"
'sqls = sqls & " where nocuenta =" & ncuent
'sqls = sqls & " and Producto=" & prod
'sqls = sqls & " ORDER BY NoTarjeta,TIPO,FechaModificacion"

Set rsDatos = New ADODB.Recordset
rsDatos.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

With spdtarjetas
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
    .Text = rsDatos!NoTarjeta
    .Col = 2
    .Text = rsDatos!Nombre
    .Col = 3
    .Text = rsDatos!FechaModificacion
    .Col = 4
    .Text = Format(rsDatos!FechaCancelacion, "mm/dd/yy")
    .Col = 5
    .Text = rsDatos!tipo
    .Col = 6
    .Text = rsDatos!Status
    If .Text = 1 Then
       .Text = "ACTIVA"
    Else
       .Text = "CANCELADA"
    End If
        
    If rsDatos!Status = 2 Then
        For i = 1 To 6
            .Row = j
            .Col = i
            .BackColor = &HC0C0FF
        Next
    End If
    rsDatos.MoveNext

Loop
spdtarjetas.Visible = True
End With
rsDatos.Close
Set rsDatos = Nothing
Screen.MousePointer = 1
Else
  MsgBox "No hay cuenta que mostrar", vbCritical, "Especifique una cuenta"
  TxtCuenta.SetFocus
End If
End Sub

Sub buscadispxlib(ncuent As Double)
Dim rsDatos As ADODB.Recordset
Dim j As Integer, i As Integer
Dim manana As Date
manana = Date + 1
If TxtCuenta.Text <> "" Then
Screen.MousePointer = 11
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "sp_Consultas_BE '" & corrige_fecha(CStr(manana)) & "',Null," & Product & "," & ncuent & ",Null,Null,Null,'Buscadispxlib'"

'sqls = " select top 1 l.Cuenta,l.Importe,b.BON_FAC_FECHAENT as Fecha_Dispersion from liquidacionesbe l(nolock),cuentasbe c(nolock),bon_factura b(nolock)"
'sqls = sqls & " Where c.nocuenta = l.cuenta"
'sqls = sqls & " and l.Cuenta = " & ncuent
'sqls = sqls & " and l.comercio='DISP'"
'sqls = sqls & " And c.Status=1"
'sqls = sqls & " and l.cliente=b.BON_FAC_CLIENTE"
'sqls = sqls & " and b.BON_FAC_TPOBON=" & prod
'sqls = sqls & " and b.BON_FAC_FECHAENT='" & corrige_fecha(CStr(manana)) & "'"
'sqls = sqls & " and b.BON_FAC_STATUS=1"

Set rsDatos = New ADODB.Recordset
rsDatos.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly

With spddisp
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
    .Text = rsDatos!Cuenta
    .Col = 2
    .Text = Val(rsDatos!importe)
    .Col = 3
    .Text = rsDatos!Fecha_Dispersion
     rsDatos.MoveNext
Loop
spddisp.Visible = True
End With
rsDatos.Close
Set rsDatos = Nothing
Screen.MousePointer = 1
Else
  MsgBox "No hay cuenta que mostrar", vbCritical, "Especifique una cuenta"
  TxtCuenta.SetFocus
End If
End Sub
Sub InicializaForma()
    LimpiarControles Me
    Call BuscaFechaCorte(0)
    If Product = 1 Then
       txttarjeta.Text = "50640601"
    ElseIf Product = 2 Then
      txttarjeta.Text = "50640501"
    End If
    txtCliente.Text = ""
    TxtCuenta.Text = ""
    txtEmpleado.Text = ""
    txtEstado.Text = ""
    TxtNombre.Text = ""
    txtSaldo.Text = ""
    txtTarjeta2.Text = ""
    txtTipo.Text = ""
    spdMov.MaxRows = 0
End Sub




