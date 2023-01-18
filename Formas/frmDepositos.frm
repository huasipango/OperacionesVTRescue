VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmDepositos 
   Caption         =   "Depositos por comisiones BE"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spddetalle 
      Height          =   3855
      Left            =   240
      OleObjectBlob   =   "frmDepositos.frx":0000
      TabIndex        =   5
      Top             =   1800
      Width           =   10935
   End
   Begin FPSpread.vaSpread spdTipo 
      Height          =   1935
      Left            =   3120
      OleObjectBlob   =   "frmDepositos.frx":0470
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quitar"
      CausesValidation=   0   'False
      Height          =   525
      Left            =   1680
      Picture         =   "frmDepositos.frx":06FD
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar registro seleccionado"
      Top             =   5760
      Width           =   765
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Agregar"
      CausesValidation=   0   'False
      Height          =   525
      Left            =   600
      Picture         =   "frmDepositos.frx":07FF
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agregar  nuevo regsitro"
      Top             =   5760
      Width           =   765
   End
   Begin VB.TextBox txtdepositos 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9120
      TabIndex        =   10
      Top             =   5800
      Width           =   1575
   End
   Begin VB.TextBox txtingresos 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      Top             =   5800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   10935
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   9720
         Picture         =   "frmDepositos.frx":0901
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar folio"
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   10320
         Picture         =   "frmDepositos.frx":0A03
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   9120
         Picture         =   "frmDepositos.frx":0B05
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar Ficha"
         Top             =   720
         Width           =   400
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   345
         Left            =   9600
         TabIndex        =   1
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Aplicación:"
         Height          =   195
         Left            =   7920
         TabIndex        =   12
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Suma de Depositos:"
      Height          =   195
      Left            =   7680
      TabIndex        =   15
      Top             =   5880
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Suma de Ingresos:"
      Height          =   195
      Left            =   4440
      TabIndex        =   14
      Top             =   5880
      Width           =   1320
   End
End
Attribute VB_Name = "frmDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim continua As Boolean, foliu As Integer
Dim modifica As Boolean

Private Sub cmdCancelar_Click()
With spddetalle
    .Row = .ActiveRow
    RESP = MsgBox("¿Esta seguro de que desea borrar este registro?", vbYesNo + vbQuestion + vbDefaultButton2, "Quitando registro")
    
    If RESP = vbYes And .MaxRows > 0 Then
        sqls = "sp_Depositos " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & ",'" & Format(mskFecha.Text, "mm/dd/yyyy") & "',"
        sqls = sqls & "'Elimina_OI'," & .Row
        cnxBD.Execute sqls
        Call carga_detalle
        Call Suma_Depositos
        Call Suma_Ingresos
        
        '.Action = 5
        '.MaxRows = .MaxRows - 1
    End If
End With
End Sub

Private Sub cmdGrabar_Click()
Dim banco As Integer, valor As Double, fec As Date, ref As Integer, refban As Long, tipo As Integer
Dim cont As Integer
On Error GoTo err:
Call Suma_Depositos
Call Suma_Ingresos
cont = 0
With spddetalle
    For i = 1 To .MaxRows
       .Row = i
       .Col = 1
       banco = Val(.Text)
       .Col = 3
       valor = CDbl(.Text)
       .Col = 4
       fec = Format(.Text, "mm/dd/yyyy")
       .Col = 5
       ref = CInt(.Text)
       .Col = 6
       refban = CLng(.Text)
       .Col = 7
       tipo = Val(.Text)
       If CDbl(valor) <> 0 And Val(tipo) <> 0 And IsDate(Format(fec, "mm/dd/yyyy")) Then
          sqls = "sp_InsertaDepositosOI @Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
          sqls = sqls & " ,@Fecha='" & Format(mskFecha.Text, "mm/dd/yyyy") & "'"
          sqls = sqls & " ,@Registro=" & i
          sqls = sqls & " ,@Banco=" & banco
          sqls = sqls & " ,@Valor=" & CDbl(valor)
          sqls = sqls & " ,@Fecha_dep='" & Format(fec, "mm/dd/yyyy") & "'"
          sqls = sqls & " ,@Referencia=" & ref
          sqls = sqls & " ,@RefBanco=" & refban
          sqls = sqls & " ,@Tipomov=" & tipo
          sqls = sqls & " ,@Usuario='" & gstrUsuario & "'"
          cnxBD.Execute sqls
          cont = cont + 1
       End If
    Next
    If cont > 0 Then
       MsgBox "Datos actualizados", vbInformation, "Informacion actualizada"
    End If
End With
Call carga_detalle
Call Suma_Depositos
Call Suma_Ingresos
Exit Sub
err:
  MsgBox err.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub carga_detalle()
Dim i As Integer
On Error GoTo err:
  sqls = "sp_Depositos " & cboBodegas.ItemData(cboBodegas.ListIndex)
  sqls = sqls & ",'" & Format(mskFecha.Text, "mm/dd/yyyy") & "',"
  sqls = sqls & "'Detalle_OI'"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
  i = 0
  
  With spddetalle
    .Col = -1
    .Row = -1
    .Action = 12
    Do While Not rsBD.EOF
        If rsBD!valor <> 0 Then
           i = i + 1
           .MaxRows = i
           .Row = i
           .Col = 1
           .Text = rsBD!banco
           .Col = 2
           .Text = rsBD!nombre
           .Col = 3
           .Text = CDbl(rsBD!valor)
           .Col = 4
           .Text = Format(rsBD!Fecha_Deposito, "mm/dd/yyyy")
           .Col = 5
           .Text = CLng(rsBD!referencia)
           .Col = 6
           .Text = CLng(rsBD!referbanco)
           .Col = 7
           .Text = rsBD!Tipo_mov
           .Col = 8
           .Text = Trim(rsBD!Descripcion)
         End If
        rsBD.MoveNext
    Loop
  End With
Exit Sub
err:
   MsgBox err.Description, vbCritical, "Errores generados"
End Sub

Private Sub cmdNuevo_Click()
    With spddetalle
    .MaxRows = .MaxRows + 1
    .Row = .ActiveRow + 1
    .Col = 1
    .SetFocus
    .Action = SS_ACTION_ACTIVE_CEL
    End With
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    mskFecha.Text = Format(Date, "mm/dd/yyyy")
    CargaBodegas cboBodegas
    spddetalle.MaxRows = 1
    Call carga_detalle
    Call Suma_Depositos
    Call Suma_Ingresos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And IsDate(mskFecha) Then
    Call carga_detalle
    Call Suma_Ingresos
    Call Suma_Depositos
 End If
End Sub

Private Sub mskFecha_LostFocus()
      Call Suma_Ingresos
End Sub

Private Sub spdDetalle_DblClick(ByVal Col As Long, ByVal Row As Long)
 If Col = 1 Then
    spdTipo.Visible = True
    sqls = " select banco id, nombre descripcion from bancos"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    With spdTipo
    i = 0
    .MaxRows = i
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!id
        .Col = 2
        .Text = rsBD!Descripcion
        rsBD.MoveNext
    Loop
    End With
 ElseIf Col = 7 Then
    spdTipo.Visible = True
    sqls = " select tipo_mov id, descripcion from fm_tipos_mov_cartera" & _
           " Where Tipo_mov >= 50" & _
           " and subtipo = 'P'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    With spdTipo
    i = 0
    .MaxRows = i
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!id
        .Col = 2
        .Text = rsBD!Descripcion
        rsBD.MoveNext
    Loop
    Call Suma_Depositos
    End With
 End If
End Sub

Private Sub spddetalle_KeyPress(KeyAscii As Integer)
Dim valor As Double, act As Integer, dato As Integer
On Error GoTo err:
With spddetalle
     If .ActiveCol = 3 And KeyAscii = 13 Then
        If CDbl(.Text) > 0 Then
           .Col = 4
           .Text = Format(Date, "mm/dd/yyyy")
           .Action = 0
        End If
     ElseIf .ActiveCol = 4 And KeyAscii = 13 Then
        .Col = 4
        If IsDate(.Text) Then
            .Col = 5
           .Action = 0
        End If
     ElseIf .ActiveCol = 5 And KeyAscii = 13 Then
        .Col = 6
        .Action = 0
     ElseIf .ActiveCol = 6 And KeyAscii = 13 Then
        .Col = 7
        .Action = 0
     ElseIf .ActiveCol = 1 And KeyAscii = 13 Then
        act = .ActiveRow
        .Row = act
        .Col = 1
        dato = Abs(CInt(.Text))
        .Text = ""
        .Text = dato
        sqls = " select banco id, nombre descripcion from bancos WHERE Banco=" & Val(dato)
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        If Not rsBD.EOF Then
          'Call Suma_Depositos
          .Row = act
          .Col = 2
          .Text = Trim(UCase(rsBD!Descripcion))
          .Col = 3
          .Action = 0
        Else
          MsgBox "No existe el No. de Banco", vbCritical, "Error en banco"
        End If
     ElseIf .ActiveCol = 7 And KeyAscii = 13 Then
        act = .ActiveRow
        .Row = act
        .Col = 7
        dato = Abs(CInt(.Text))
        .Text = ""
        .Text = dato
        sqls = " select tipo_mov id, descripcion from fm_tipos_mov_cartera" & _
           " Where Tipo_mov=" & dato
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        If Not rsBD.EOF Then
          Call Suma_Depositos
          .Row = act
          .Col = 8
          .Action = 0
          .Text = Trim(UCase(rsBD!Descripcion))
          cmdNuevo.SetFocus
        Else
           MsgBox "Tipo de movimientos no válido", vbCritical, "Tipo no válido"
        End If
     End If
End With
Exit Sub
err:
   MsgBox err.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub

Private Sub spdTipo_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim id As Integer, van As Integer, nombre As String
    
    spdTipo.Row = Row
    spdTipo.Col = 1
    id = spdTipo.Text
    spdTipo.Col = 2
    nombre = Trim(UCase(spdTipo.Text))
    
    spddetalle.Row = spddetalle.ActiveRow
    spddetalle.Col = spddetalle.ActiveCol
    van = spddetalle.ActiveCol
    spddetalle.Text = id
    spddetalle.Col = spddetalle.Col + 1
    spddetalle.Text = nombre
    '
    spddetalle.Col = van + 2
    spddetalle.Action = 0
    
    If van > 1 Then
       Call Suma_Depositos
    End If
    spdTipo.Visible = False
    spddetalle.SetFocus
End Sub

Sub Suma_Depositos()
Dim i As Integer, importe As Double, tipomov As Integer
Dim total As Double, totstr As String
With spddetalle
total = 0
For i = 1 To .MaxRows
    .Row = i
    .Col = 3
    If .Text <> "" Then
       importe = CDbl(.Text)
    End If
    .Col = 7
    tipomov = Val(.Text)
    If Val(importe) <> 0 And Val(tipomov) <> 0 Then
       total = total + importe
    End If
Next
txtdepositos.Text = Format(CCur(total), "$###,###,###,##0.00")
End With
End Sub

Sub Suma_Ingresos()
    sqls = "sp_ConsultaOI @Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
    sqls = sqls & " ,@Accion='SumaIng',@FechaIni='" & Format(mskFecha.Text, "mm/dd/yyyy") & "'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       txtingresos.Text = Format(CCur(rsBD!total), "$###,###,###,##0.00")
    End If
    
End Sub
