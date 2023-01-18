VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPagosCom 
   Caption         =   "Pago de Facturas de Comisiones BE"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spdDetalle 
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "frmPagosCom.frx":0000
      TabIndex        =   0
      Top             =   2040
      Width           =   10095
   End
   Begin FPSpread.vaSpread spdTipo 
      Height          =   1935
      Left            =   6240
      OleObjectBlob   =   "frmPagosCom.frx":04A3
      TabIndex        =   11
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5175
      TabIndex        =   14
      Top             =   5085
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   7560
         Picture         =   "frmPagosCom.frx":0728
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Buscar Ficha"
         Top             =   720
         Width           =   400
      End
      Begin VB.TextBox txtfolio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   200
         Width           =   1215
      End
      Begin VB.ComboBox cboOrigen 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   9360
         Picture         =   "frmPagosCom.frx":082A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   700
         Width           =   400
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   8760
         Picture         =   "frmPagosCom.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Grabar folio"
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   8160
         Picture         =   "frmPagosCom.frx":0A2E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   700
         Width           =   400
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1215
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   2760
         TabIndex        =   4
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1215
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   345
         Left            =   8640
         TabIndex        =   12
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Origen:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   1250
         Width           =   150
      End
      Begin VB.Label lblAño1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1250
         Width           =   255
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total capturado:"
      Height          =   195
      Left            =   3945
      TabIndex        =   15
      Top             =   5160
      Width           =   1170
   End
End
Attribute VB_Name = "frmPagosCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim continua As Boolean, foliu As Integer
Dim modifica As Boolean

Private Sub cboBodegas_Click()
    CargaFacturas
End Sub

Sub valida_datos()
On Error GoTo ERR:
   continua = True
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdBuscar_Click()
Dim resp, i As Integer, valor As Double, cliente As Integer
On Error GoTo ERR:
   resp = InputBox("Proporcione el folio a consultar: ", "Buscando Ficha de ingreso")
   If Val(resp) > 0 Then
      sqls = "sp_ConsultaOI @Accion='Folio',@Factura=" & Val(resp)
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
         modifica = True
         With spddetalle
         txtFolio.Text = rsBD!Ficha
         i = 0
         .MaxRows = 0
              Do While Not rsBD.EOF
                  i = i + 1
                  .MaxRows = i
                  .Row = i
                  .Col = 1
                  .Text = rsBD!refer_apl
                  .Col = 2
                  cliente = rsBD!cliente
                  .Text = rsBD!cliente
                  .Col = 3
                  .Text = UCase(rsBD!Descripcion)
                  .Col = 5
                  valor = CDbl(rsBD!importe)
                  .Text = CDbl(rsBD!importe)
                  .Col = 6
                  .Text = rsBD!TIPO_MOV
                  .Col = 7
                  .Text = rsBD!banco
                  .Col = 8
                  .Text = rsBD!Cheque
                  '--saldo
                  .Col = 4
                  If valor > 0 Then
                     sqls = "sp_ConsultaOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & cliente & "," & Val(rsBD!refer_apl) & ",'Saldo'"
                     Set rsBD2 = New ADODB.Recordset
                     rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
                     If Not rsBD2.EOF Then
                        .Text = CDbl(rsBD2!saldo)
                     Else
                        .Text = 0
                     End If
                  End If
                  '-----
                  rsBD.MoveNext
              Loop
              Call Suma_Todo
         End With
      Else
         MsgBox "No existe el folio proporcionado", vbExclamation, "Folio no encontrado"
      End If
   End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdGrabar_Click()
Dim hubo As Integer
Call valida_datos
hubo = 0
If continua = False Then
   Exit Sub
End If
Call Suma_Todo

If modifica = False Then
   sqls = "sp_ConsultaOI @Accion='Nuevo'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   If rsBD.EOF Then
      foliu = 1
   Else
      If IsNull(rsBD!Nuevo) Then
         foliu = 1
      Else
         foliu = rsBD!Nuevo
      End If
   End If
   txtFolio.Text = Val(foliu)
End If
modifica = False

With spddetalle
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        Factura = .Text
        .Col = 2
        Grupo = .Text
        .Col = 5
        importe = .Text
        .Col = 6
        tipomov = .Text
        .Col = 7
        banco = .Text
        .Col = 8
        Refer = .Text
        
        If Val(importe) <> 0 And Val(tipomov) <> 0 Then
            sqls = "exec sp_UpdFM_Facturas " & cboBodegas.ItemData(cboBodegas.ListIndex)
            sqls = sqls & ", " & Val(Factura) & ", " & Val(importe) & ", " & Val(tipomov) & ", " & Val(banco) & ", " & Val(Refer) & ", '" & Format(mskFecha, "mm/dd/yyyy") & "'"
            'cnxBD.Execute sqls, intRegistros
                        
            sqls = "sp_ConsultaOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & Grupo & "," & Factura & ",'Factura'"
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
               sqls = " EXEC sp_fm_Clientes_mov_ins2 "
               sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
               sqls = sqls & vbCr & ", @Cliente      = " & Grupo
               sqls = sqls & vbCr & ", @Fecha        =  '" & Format(Date, "MM/DD/YYYY") & "'"
               sqls = sqls & vbCr & ", @Tipo_Mov     = " & Val(tipomov)
               sqls = sqls & vbCr & ", @Serie        = '" & Trim(rsBD!serie) & "'"
               sqls = sqls & vbCr & ", @Refer        = " & Val(i + 10)
               sqls = sqls & vbCr & ", @Refer_Apl    = " & Factura
               sqls = sqls & vbCr & ", @CarAbo       = 'A'"
               sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
               sqls = sqls & vbCr & ", @Importe      = " & Val(importe)
               sqls = sqls & vbCr & ", @iva      = Null" '& rsBD!Iva
               sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
               sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
               sqls = sqls & vbCr & ", @Usuario = '" & Usuario & "'"
               sqls = sqls & vbCr & ", @Ficha = " & Val(txtFolio.Text)
               cnxBD.Execute sqls, intRegistros
            
               sqls = " EXEC sp_IngresosOI_updins "
               sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
               sqls = sqls & vbCr & ", @Fecha        ='" & Format(Date, "MM/DD/YYYY") & "'"
               sqls = sqls & vbCr & ", @Ficha        = " & Val(txtFolio)
               sqls = sqls & vbCr & ", @Registro     = " & Val(i + 10)
               sqls = sqls & vbCr & ", @Vendedor     = 0"
               sqls = sqls & vbCr & ", @Recibo       = " & Val(i + 10)
               sqls = sqls & vbCr & ", @Cliente      = " & Grupo
               sqls = sqls & vbCr & ", @Tipo_mov     = " & Val(tipomov)
               sqls = sqls & vbCr & ", @Serie        = '" & Trim(rsBD!serie) & "'"
               sqls = sqls & vbCr & ", @Refer_Apl    = " & Factura
               sqls = sqls & vbCr & ", @Tipo_mov_apl = 10"
               sqls = sqls & vbCr & ", @Importe      = " & importe
               sqls = sqls & vbCr & ", @Fecha_Apl    = '" & Format(Date, "MM/DD/YYYY") & "'"
               sqls = sqls & vbCr & ", @Banco        = " & banco
               sqls = sqls & vbCr & ", @Cheque       = " & Val(Refer)
               sqls = sqls & vbCr & ", @Status       = 'S'"
               sqls = sqls & vbCr & ", @Aplicado     = 'S'"
               sqls = sqls & vbCr & ", @Poliza = " & Val(txtFolio)
               sqls = sqls & vbCr & ", @origen = " & cboOrigen.ItemData(cboOrigen.ListIndex)
               sqls = sqls & vbCr & ", @cifra = " & CDbl(txtTotal)
               cnxBD.Execute sqls, intRegistros
               hubo = hubo + 1
            End If
        End If
        
    Next i
End With
If hubo > 0 Then
   MsgBox "Datos Actualizados!!" & vbCrLf & "Folio: " & Format(txtFolio.Text, "00000"), vbInformation, "Informacion actualizada"
End If
txtFolio.Text = ""
txtTotal.Text = ""
CargaFacturas
End Sub

Private Sub cmdImprimir_Click()
 Imprime crptToWindow
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptLibroComBE.rpt"
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = cboBodegas.ItemData(cboBodegas.ListIndex)
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = 0
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical, "Errores generados en el reporte"
        Exit Sub
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    mskFechaIni.Text = CDate("01/" & Month(Date) & "/" & Year(Date))
    mskFechaFin.Text = Date
    mskFecha.Text = Date
    CargaBodegas cboBodegas
    CargaBodegas cboOrigen
    spdTipo.Visible = False
    CargaFacturas
    Call STAT
End Sub

Sub CargaFacturas()
    'sqls = "exec spr_librocomBe " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & _
    '       "'" & Format(mskFechaIni.Text, "mm/dd/yyyy") & "','" & Format(mskFechaFin.Text, "mm/dd/yyyy") & "', 1"
           
    modifica = False
    sqls = "sp_ConsultaOI @Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex) & ",@FechaIni='" & Format(mskFechaIni.Text, "mm/dd/yyyy") & "',"
    sqls = sqls & " @FechaFin='" & Format(mskFechaFin.Text, "mm/dd/yyyy") & "',@Accion='Carga'"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
    i = 0
    
    With spddetalle
    .Col = -1
    .Row = -1
    .Action = 12
    
    Do While Not rsBD.EOF
        If rsBD!saldo <> 0 Then
           i = i + 1
           .MaxRows = i
           .Row = i
           .Col = 1
           .Text = rsBD!Factura
           .Col = 2
           .Text = rsBD!Grupo
           .Col = 3
           .Text = rsBD!Descripcion
           .Col = 4
           .Text = Val(rsBD!saldo)
         End If
        rsBD.MoveNext
    Loop
    End With
    Call STAT
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsDate(mskFechaIni.Text) Then
    CargaFacturas
End If
End If
End Sub

Private Sub mskFechaFin_LostFocus()
'If IsDate(mskFechaIni.Text) Then
'    CargaFacturas
'End If
End Sub

Private Sub mskFechaIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsDate(mskFechaIni.Text) Then
      txtFolio.Text = ""
      txtTotal.Text = ""
      CargaFacturas
   End If
End If
End Sub

Private Sub mskFechaIni_LostFocus()
'If IsDate(mskFechaIni.Text) Then
'    CargaFacturas
'End If
End Sub

Private Sub spdDetalle_DblClick(ByVal Col As Long, ByVal Row As Long)
If Col = 5 Then
    spddetalle.Row = Row
    spddetalle.Col = 4
    importe = spddetalle.Text
    spddetalle.Col = 5
    spddetalle.Text = importe

ElseIf Col = 6 Then
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
    End With
ElseIf Col = 7 Then
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

End If
    
End Sub

Sub Suma_Todo()
Dim i As Integer, importe As Double, tipomov As Integer
Dim total As Double, totstr As String
With spddetalle
total = 0
For i = 1 To .MaxRows
    .Row = i
    .Col = 5
    importe = Val(.Text)
    .Col = 6
    tipomov = Val(.Text)
    If Val(importe) <> 0 And Val(tipomov) <> 0 Then
       total = total + importe
    End If
Next
txtTotal.Text = Format(CCur(total), "$###,###,###,##0.00")
End With
End Sub


Private Sub spddetalle_KeyPress(KeyAscii As Integer)
Dim valor As Double, act As Integer, dato As Integer
On Error GoTo ERR:
With spddetalle
   If .ActiveCol = 5 And KeyAscii = 8 Then
      act = .ActiveRow
      .Row = act
      Call Suma_Todo
   End If
   If .ActiveCol = 5 And KeyAscii = 13 Then
      act = .ActiveRow
      .Row = act
      .Col = 1
      Factura = .Text
      .Col = 2
      Grupo = .Text
      .Col = 5
      valor = Abs(CDbl(.Text))
      .Text = ""
      .Text = valor
     If valor > 0 Then
        If modifica = False Then
        sqls = "sp_ConsultaOI " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & Grupo & "," & Factura & ",'Saldo'"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        If Not rsBD.EOF Then
           If rsBD!saldo <= 0 Then
              If MsgBox("La factura a la que le quiere capturar este pago ya esta saldada. Saldo=$" & Val(rsBD!saldo) & vbCrLf & "¿Desea capturar otro pago?", vbExclamation + vbYesNo + vbDefaultButton2, "Factura ya esta saldada") = vbYes Then
                 .Col = 6
                 .Action = 0
                 '.ActiveCol = 6
              End If
           Else
              If valor > rsBD!saldo Then
                 If MsgBox("El monto a capturar es mayor que el adeudo pendiente. Saldo=$" & Val(rsBD!saldo) & vbCrLf & "¿Aun asi desea capturarlo?", vbExclamation + vbYesNo + vbDefaultButton2, "Importe sobrepasa el adeudo de la factura") = vbYes Then
                    .Col = 6
                    .Action = 0
                 End If
              Else
                 .Col = 6
                 .Action = 0
              End If
           End If
        Else
           MsgBox "No hay informacion con los datos proporcionados", vbCritical, "Vacío..."
        End If
        Else
           .Col = 6
           .Action = 0
        End If '--end de modifica=true
     End If
   ElseIf .ActiveCol = 6 And KeyAscii = 13 Then
       act = .ActiveRow
      .Row = act
       .Col = 6
       dato = Abs(CInt(.Text))
       .Text = ""
       .Text = dato
       sqls = " select tipo_mov id, descripcion from fm_tipos_mov_cartera" & _
           " Where Tipo_mov >= 50" & _
           " and subtipo = 'P' And Tipo_Mov=" & Val(dato)
       Set rsBD = New ADODB.Recordset
       rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
       If Not rsBD.EOF Then
          .Col = 7
          .Action = 0
       End If
   ElseIf .ActiveCol = 7 And KeyAscii = 13 Then
       act = .ActiveRow
      .Row = act
       .Col = 7
       dato = Abs(CInt(.Text))
       .Text = ""
       .Text = dato
       sqls = " select banco id, nombre descripcion from bancos WHERE Banco=" & Val(dato)
       Set rsBD = New ADODB.Recordset
       rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
       If Not rsBD.EOF Then
          Call Suma_Todo
          .Row = act
          .Col = 8
          .Action = 0
       End If
   End If
End With
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
End Sub

Private Sub spddetalle_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR:
With spddetalle
   If .ActiveCol = 5 And KeyCode = 46 Then
      act = .ActiveRow
      .Row = act
      Call Suma_Todo
   End If
   If .ActiveCol = 6 And KeyCode = 115 Then
      act = .ActiveRow
      .Row = act
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
      End With
      Call Suma_Todo
   End If
   If .ActiveCol = 7 And KeyCode = 115 Then
      act = .ActiveRow
      .Row = act
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
   End If
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub spdTipo_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim id As Integer, van As Integer
    
    spdTipo.Row = Row
    spdTipo.Col = 1
    id = spdTipo.Text
    
    spddetalle.Row = spddetalle.ActiveRow
    spddetalle.Col = spddetalle.ActiveCol
    van = spddetalle.ActiveCol
    spddetalle.Text = id
    '
    spddetalle.Col = van + 1
    spddetalle.Action = 0
    
    Call Suma_Todo
    spdTipo.Visible = False
    spddetalle.SetFocus
End Sub

Sub STAT()
    sql = "SELECT * FROM CONFIGBODEGAS WHERE Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       If Month(Date) <> Month(rsBD!FechaCierre) Then
          cmdGrabar.Enabled = False
       Else
          cmdGrabar.Enabled = True
       End If
    Else
       MsgBox "Error no hay fecha de cierre actual para esa Sucursal", vbCritical, "Sin fecha"
    End If
End Sub

