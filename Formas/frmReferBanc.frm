VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmReferBanc 
   Caption         =   "Integración Bancaria"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   16785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   16575
      Begin FPSpread.vaSpread spdDepositos 
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "frmReferBanc.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   16275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin MSComctlLib.ProgressBar prgGuardar 
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   720
         Picture         =   "frmReferBanc.frx":1163
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   120
         Picture         =   "frmReferBanc.frx":1265
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Subir desde archivo"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   1320
         Picture         =   "frmReferBanc.frx":1367
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Integración Bancaria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9240
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog cmnAbrir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReferBanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Dim TotDepositos As Single
Dim Folio As Long
Dim Serie As String

Private Sub cmdAbrir_Click()
   Call Subearchivo
End Sub

Private Function Subearchivo()
    Dim intFile As Integer
    Dim strBuffer As String
    Dim varArray As Variant
    Dim Renglon As Integer
    Dim Referencia As String
    Dim ReferValida As Integer
    Dim MontoPedido As Long
    Dim MontoPago As Long
    Dim PosicionInicial As Integer
    Dim PosicionFinal As Integer
    
    Dim Fecha As String
    Dim Concepto As String
    Dim Monto As String
    
    Dim SFILE As String
    
    TotDepositos = 0
    intFile = FreeFile
    On Error GoTo ErrorImport
    cmnAbrir.ShowOpen
    If cmnAbrir.Filename <> "" Then
        SFILE = Mid(cmnAbrir.Filename, InStrRev(cmnAbrir.Filename, "\") + 1)
        cmdAbrir.Enabled = False
        prgGuardar.Visible = True
        Open cmnAbrir.Filename For Input As #intFile
        Renglon = 1
        Do Until EOF(intFile)
            Line Input #intFile, strBuffer
            Renglon = Renglon + 1
        Loop
        Close #intFile
        prgGuardar.Max = Renglon - 1
                
        Open cmnAbrir.Filename For Input As #intFile
        With spdDepositos
        Renglon = 1
        Do Until EOF(intFile)
            DoEvents
            Referencia = ""
            prgGuardar.value = Renglon
            Line Input #intFile, strBuffer
            Fecha = ""
            Concepto = ""
            Monto = ""
            If Left(SFILE, 2) = "MD" Then   ' Es archivo tipo NetCash Diario
                If Left(strBuffer, 2) = 22 Then
                    Fecha = Mid(strBuffer, 129, 8)
                    Concepto = Mid(strBuffer, 13, 89)
                    Monto = Val(Mid(strBuffer, 90, 17))
                End If
            ElseIf Left(SFILE, 2) = "MH" Then   ' Es archivo tipo NetCash Historico
                If Left(strBuffer, 2) = 22 Then
                    Fecha = Mid(strBuffer, 15, 2) & "/" & Mid(strBuffer, 13, 2) & "/" & Mid(strBuffer, 11, 2)
                    Concepto = Mid(strBuffer, 53)
                    Monto = Val(Mid(strBuffer, 29, 14)) / 100
                    Line Input #intFile, strBuffer
                    Concepto = Concepto & Mid(strBuffer, 5, 38)
                End If
            Else    ' Es archivo tipo BeCom
                varArray = Split(strBuffer, vbTab)
                Fecha = varArray(0)
                Concepto = varArray(1)
                Monto = varArray(3)
            End If
            If Val(Monto) > 0 Then
                .Row = Renglon
                .MaxRows = Renglon
                .Col = 1
                .Text = Fecha
                .Col = 2
                .Text = Concepto
                .Col = 7
                MontoPago = Monto
                .Text = Monto
                .Col = 6
                ' Busco referencia
                PosicionInicial = 0
                PosicionFinal = 0
                PosicionInicial = InStr(Concepto, "PT")
                If PosicionInicial = 0 Then
                    PosicionInicial = InStr(Concepto, "PD")
                End If
                PosicionFinal = InStr(Concepto, "/")
                If PosicionFinal = 0 And PosicionInicial > 0 Then
                    PosicionFinal = InStr(Mid(Concepto, PosicionInicial), " ") + PosicionInicial - 1
                End If
                If PosicionFinal < PosicionInicial Then
                    PosicionFinal = Len(Concepto) + 1
                End If
                If PosicionInicial > 0 And PosicionFinal - PosicionInicial > 0 Then
                    Referencia = Trim(Mid(Concepto, PosicionInicial, PosicionFinal - PosicionInicial))
                    If Len(Referencia) <> Len(Concepto) Then
                        .Text = Referencia
                    End If
                End If
                
                sqls = "sp_BuscaDepositos '" & Format(CDate(Fecha), "mm/dd/yyyy") & "','" & Concepto & "'"
                Set rsBD = New ADODB.Recordset
                rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                If rsBD.EOF Then
                    sqls = "sp_BuscaDepositos '" & Format(CDate(Fecha), "mm/dd/yyyy") & "','" & Referencia & "'"
                    Set rsBD2 = New ADODB.Recordset
                    rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    If rsBD2.EOF Then
                        .Col = 6
                        ReferValida = ValidaReferencia(.Text)
                        Select Case ReferValida
                            Case 0
                                ' Busco que la factura exista
                                sqls = "sp_BuscaReferenciaBanc '" & .Text & "'"
                                Set rsBD3 = New ADODB.Recordset
                                rsBD3.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                                If Not rsBD3.EOF Then
                        If Left(rsBD3!Status, 9) = "Cancelado" Then
                                        .Col = 8
                                        .value = False
                                        .Col = 9
                                        .value = "El pedido está cancelado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "3" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "8" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "1" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "2" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "6" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "7" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "10" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "11" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "16" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "Dispersado" And Left(rsBD3!Producto, 2) = "17" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD3!Status, 10) = "En Proceso" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "Solicitar a Operaciones integrar respuesta de Syc"
                                    Else
                                        .Col = 3
                                        .value = rsBD3!cliente
                                        .Col = 4
                                        .value = Str(rsBD3!Pedido)
                                        .Col = 5
                                        .value = Str(Round(rsBD3!MontoPedido, 2))
                                        MontoPedido = rsBD3!MontoPedido
                                        .Col = 9
                                        If Left(Referencia, 2) = "PD" Then
                                            .Text = "Pago Dispersion Cliente " & rsBD3!Nombre
                                        Else
                                            .Text = "Pago Tarjetas Cliente " & rsBD3!Nombre
                                        End If
                                        TotDepositos = TotDepositos + Monto
                                        .Col = 10
                                        .value = Str(rsBD3!Producto)
                                        If Round(MontoPedido - MontoPago) <> 0 Then
                                            .Col = 8
                                            .value = False
                                            .Col = 9
                                            .value = "El monto del pedido es distinto al monto de pago, por lo tanto no se puede hacer la integración automática"
                                        Else
                                            .Col = 8
                                            .value = True
                                        End If
                                    End If
                                Else
                                    .Col = 8
                                    .value = False
                                    .Col = 9
                                    .Text = "Referencia no existente"
                                End If
                                rsBD3.Close
                                Set rsBD3 = Nothing
    
                            Case 1
                                .Col = 8
                                .value = False
                                .Col = 9
                                .Text = "La longitud de la referencia es muy corta"
                            Case 2
                                .Col = 8
                                .value = False
                                .Col = 9
                                .Text = "La longitud de la referencia es muy larga"
                            Case 3
                                .Col = 8
                                .value = False
                                .Col = 9
                                .Text = "El digito verficador es incorrecto"
                            Case 4
                                .Col = 8
                                .value = False
                                .Col = 9
                                .Text = "No cuenta con una referencia"
                            Case Else
                                .Col = 8
                                .value = False
                                .Col = 9
                                .value = "Error desconocido, favor de avisar a sistemas"
                        End Select
                    Else
                        .Col = 8
                        .value = False
                        .Col = 9
                        .value = "Ya fue integrado el dia: " & Format(rsBD2!FechaIntegracion, "dd/mm/yyyy")
                    End If
                    rsBD2.Close
                    Set rsBD2 = Nothing
                Else
                    .Col = 8
                    .value = False
                    .Col = 9
                    .value = "Ya fue integrado el dia: " & Format(rsBD!FechaIntegracion, "dd/mm/yyyy")
                End If
                rsBD.Close
                Set rsBD = Nothing
                Renglon = Renglon + 1
            End If
        Loop
        End With
        Close #intFile
        prgGuardar.Visible = False
        prgGuardar.value = 0
        Exit Function
    End If
    Exit Function
ErrorImport:
    Beep
    MsgBox "Hubo un error al actualizar! Favor de avisar a sistemas! Error: " & ERR.Number & vbCrLf & ERR.Description, vbCritical + vbOKOnly, Me.Caption
    Screen.MousePointer = 1
    Resume Next
End Function
Private Function ValidaReferencia(Referencia As String)
Dim tipo As String
Dim Refer As String
Dim ReferNum As String
Dim rsBD4 As ADODB.Recordset
    
    ValidaReferencia = 99
    If Len(Referencia) = 0 Then ' No cuenta con referencia
        ValidaReferencia = 4
        Exit Function
    End If
    If Len(Referencia) < 4 Then ' La longitud de la referencia es muy corta
        ValidaReferencia = 1
        Exit Function
    End If
    If Len(Referencia) > 20 Then ' La longitud de la referencia es muy larga
        ValidaReferencia = 2
        Exit Function
    End If
    Refer = Left(Referencia, Len(Referencia) - 1)
    tipo = Left(Referencia, 2)
    ReferNum = Right(Refer, Len(Refer) - 2)
    
    ' Valido el DV sea correcto
    sqls = "Select dbo.fg_Alg36('" & Refer & "') AS ReferBD"
    Set rsBD4 = New ADODB.Recordset
    rsBD4.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD4.EOF Then
        If rsBD4!ReferBD <> Referencia Then ' El digito verificador es incorrecto
            ValidaReferencia = 3
            Exit Function
        End If
    End If
    rsBD4.Close
    Set rsBD4 = Nothing
                       
    ' Todo esta bien
    ValidaReferencia = 0
End Function
Private Sub cmdGrabar_Click()
Dim Fecha As Date
Dim Concepto As String
Dim cliente As String
Dim Pedido As String
Dim MontoPedido As String
Dim Referencia As String
Dim MontoPago As String
Dim Producto As String
Dim Ficha As Integer
Dim Deposito As Integer
Dim Registro As Integer

If TotDepositos > 0 Then
    cmdGrabar.Enabled = False
  If spdDepositos.MaxRows = 0 Then Exit Sub
    prgGuardar.Visible = True
    prgGuardar.Max = spdDepositos.MaxRows
        
    'Activo Cobranza del dia
    sqls = "sp_Bloqueacobranza @Bodega=1,@Fecha='" & Format(Date, "mm/dd/yyyy") & "',@Accion='Activar'"
    cnxbdMty.Execute sqls, intRegistros
    
    'Busco siguiente ficha de ingreso
    sqls = "sp_FichaMax_Cartera @Bodega=1,@Folioman=0,@Registro=N'N'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
        Ficha = rsBD!Ficha
    End If
    rsBD.Close
    Set rsBD = Nothing
    
    ' Busco siguiente consecutivo de deposito
    sqls = "sp_Depositos @Bodega=1,@Fecha='" & Format(Date, "mm/dd/yyyy") & "',@Accion='Consecutivo'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
        Deposito = rsBD!MAX_Registro
    End If
    rsBD.Close
    Set rsBD = Nothing

    Registro = 1
      With spdDepositos
          For i = 1 To .MaxRows
                DoEvents
                prgGuardar.value = i
              .Row = i
              .Col = 8
              If .value Then
                .Col = 1
                Fecha = .Text
                .Col = 2
                Concepto = .Text
                .Col = 3
                cliente = .Text
                .Col = 4
                Pedido = .Text
                .Col = 5
                MontoPedido = .Text
                .Col = 6
                Referencia = .Text
                .Col = 7
                MontoPago = CDbl(.Text)
                .Col = 10
                Producto = .Text
                
                ' Inserto en LogDepositos
                sqls = "sp_DepositosReferenciados @Fecha='" & Format(Fecha, "mm/dd/yyyy") & "', @Concepto = '" & Concepto & "', @Referencia = '" & Referencia & "', @Monto = " & MontoPago
                cnxbdMty.Execute sqls, intRegistros
                
                ' Genero factura
                If Left(Referencia, 2) = "PD" Then
                    Call GeneraFactura(cliente, Pedido, Producto, MontoPago)
                Else
                    Call GeneraFacturaOI(cliente, Pedido, Producto, MontoPago)
                End If
                
               ' Inserto en Clientes_Movimientos
                sqls = "sp_Clientes_mov_ins_New @Bodega = 1, @Cliente = " & cliente & ", @Fecha = '" & Format(Date, "mm/dd/yyyy") & "', @Tipo_Mov = 60, @Serie = '" & Serie & "', @Refer = " & Registro & ", @Refer_Apl = " & Folio & ", @CarAbo = 'A', @Tipo_Mov_Apl = 10, @Importe = " & MontoPago & ", @Vendedor = 0, @Ficha = " & Ficha & ", @Fecha_Mov = '" & Format(Date, "mm/dd/yyyy ") & Format(Time, "hh:mm:ss") & "', @Usuario = " & Usuario
                cnxbdMty.Execute sqls, intRegistros
                        
                ' Inserto en Ingresos
                sqls = "sp_Ingresos_updins_paso @Bodega = 1, @Fecha = '" & Format(Date, "mm/dd/yyyy") & "', @Ficha = " & Ficha & ", @Registro = " & Registro & ", @Vendedor = 0, @Recibo = " & Registro & ", @Cliente = " & cliente & ",  @Tipo_Mov = 60, @Serie = '" & Serie & "', @Refer_Apl = " & Folio & ", @Tipo_Mov_Apl = 10, @Importe = " & MontoPago & ", @Fecha_Apl = '" & Format(Date, "mm/dd/yyyy") & "', @Banco = 1, @Cheque = 0, @Status = 'S', @Aplicado = 'S', @Poliza = " & Ficha & ", @Origen = 1, @Cifra = " & TotDepositos
                cnxbdMty.Execute sqls, intRegistros
                
                ' Reviso Morosidad del cliente
                sqls = "sp_IncrementaMorosidad @Bodega = 1, @Cliente = " & cliente & ",  @Serie = '" & Serie & "', @Factura = " & Folio
                cnxbdMty.Execute sqls, intRegistros
                
                ' Inserto en Depositos
                sqls = "sp_InsertaDepositos @Bodega = 1, @Fecha = '" & Format(Date, "mm/dd/yyyy") & "', @Registro = " & Deposito & ", @Banco = 1, @Valor = " & MontoPago & ", @Fecha_dep = '" & Format(Date, "mm/dd/yyyy") & "', @Referencia = " & Ficha & ", @RefBanco = '" & Referencia & "', @TipoMov = 60, @Usuario = " & Usuario
                cnxbdMty.Execute sqls, intRegistros
                Deposito = Deposito + 1
                Registro = Registro + 1
                
                ' Actualizo en FM_Facturas
                If Left(Referencia, 2) = "PT" Then
                    sqls = "sp_UpdFM_Facturas @Bodega = 1, @Factura = " & Folio & ", @Importe = " & MontoPago & ", @TipoPago = 60, @Banco = 1, @Refer = " & Ficha & ", @FechaMov = '" & Format(Date, "mm/dd/yyyy") & "'"
                    cnxbdMty.Execute sqls, intRegistros
                End If
                        
              End If
          Next i
      .Col = -1
      .Row = -1
      .Action = 12
      .MaxRows = 0
    
      End With

    ' Bloqueo cobranza
    sqls = "sp_Bloqueacobranza @Bodega=1,@Fecha='" & Format(Date, "mm/dd/yyyy") & "',@Accion='BloqueoAuto'"
    cnxbdMty.Execute sqls, intRegistros
    
    MsgBox "Los depósitos han sido aplicados!", vbInformation, "Depositos aplicados"
    prgGuardar.Visible = False
    cmdGrabar.Enabled = False
Else
    MsgBox "No hay ningún depósito por aplicar, validar información!", vbError, "Error"
End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub InicializaForma()
    spddetalle.MaxRows = 1
End Sub
Private Sub Form_Load()
Set mclsAniform = New clsAnimated
 '    InicializaForma
    TotDepositos = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
Private Sub spdDepositos_Change(ByVal Col As Long, ByVal Row As Long)
Dim MontoPedido As Single
Dim MontoPago As Single
Dim Fecha As String

    With spdDepositos
        .Col = 1
        .Row = Row
        Fecha = .Text
        .Col = Col
        Referencia = .Text
        sqls = "sp_BuscaDepositos '" & Format(CDate(Fecha), "mm/dd/yyyy") & "','" & Referencia & "'"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        If rsBD.EOF Then
            ReferValida = ValidaReferencia(.Text)
            Select Case ReferValida
                Case 0
                    ' Busco que la factura exista
                    sqls = "sp_BuscaReferenciaBanc '" & .Text & "'"
                    Set rsBD2 = New ADODB.Recordset
                    rsBD2.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
                    If Not rsBD2.EOF Then
                       If Left(rsBD2!Status, 9) = "Cancelado" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido está cancelado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "3" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "8" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "1" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "2" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "6" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "7" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "10" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "11" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "16" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "Dispersado" And Left(rsBD2!Producto, 2) = "17" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "El pedido ya fue Dispersado"
                        ElseIf Left(rsBD2!Status, 10) = "En Proceso" Then
                            .Col = 8
                            .value = False
                            .Col = 9
                            .value = "Solicitar a Operaciones integrar respuesta de Syc"
                        Else
                            .Col = 3
                            .value = rsBD2!cliente
                            .Col = 4
                            .value = Str(rsBD2!Pedido)
                            .Col = 5
                            .value = Str(rsBD2!MontoPedido)
                            MontoPedido = rsBD2!MontoPedido
                            .Col = 7
                            MontoPago = .value
                            .Col = 9
                        If Left(Referencia, 2) = "PD" Or Left(Referencia, 2) = "pd" Or Left(Referencia, 2) = "Pd" Or Left(Referencia, 2) = "pD" Then
                                .Text = "Pago Dispersion Cliente " & rsBD2!Nombre
                                .Col = 7
                            Else
                                .Text = "Pago Tarjetas Cliente " & rsBD2!Nombre
                                .Col = 7
                            End If
                            TotDepositos = TotDepositos + .value
                            .Col = 10
                            .value = Str(rsBD2!Producto)
                            If MontoPedido - MontoPago <> 0 Then
                                .Col = 8
                                .value = False
                                .Col = 9
                                .value = "El monto del pedido es distinto al monto de pago, por lo tanto no se puede hacer la integración automática"
                            Else
                                .Col = 8
                                .value = True
                            End If
                        End If
                    Else
                        .Col = 8
                        .value = False
                        .Col = 9
                        .Text = "Referencia no existente"
                    End If
                    rsBD2.Close
                    Set rsBD2 = Nothing
    
                Case 1
                    .Col = 8
                    .value = False
                    .Col = 9
                    .Text = "La longitud de la referencia es muy corta"
                Case 2
                    .Col = 8
                    .value = False
                    .Col = 9
                    .Text = "La longitud de la referencia es muy larga"
                Case 3
                    .Col = 8
                    .value = False
                    .Col = 9
                    .Text = "El digito verficador es incorrecto"
                Case Else
                    .Col = 8
                    .value = False
                    .Col = 9
                    .value = "Error desconocido, favor de avisar a sistemas"
            End Select
        Else
            .Col = 8
            .value = False
            .Col = 9
            .value = "Ya fue integrado el dia: " & Format(rsBD!FechaIntegracion, "dd/mm/yyyy")
        End If
    rsBD.Close
    Set rsBD = Nothing
    End With
End Sub

Private Sub GeneraFacturaOI(ByVal cliente As Integer, ByVal Pedido As Integer, ByVal Producto As Integer, ByVal importe As Double)
Dim impuesto As Double
Dim Subtotal As Double
Dim Bodega As Integer
Dim i As Integer

On Error GoTo err_gral
          
          
    If importe > 0 Then
        
        Bodega = 1
        
        sqls = " select Prefijo serie, consecutivo " & _
              " From folios " & _
              " Where Bodega =" & Bodega & _
              " and tipo = 'FCM'"
        Set rsBD = New ADODB.Recordset
        'rsbd.Open SQLS, cnxBD, adOpenDynamic, adLockPessimistic
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockOptimistic
            
        If rsBD.EOF = False Then
           Serie = Trim(rsBD!Serie)
           Folio = rsBD!consecutivo + 1
        Else
           MsgBox "Falta valor de la serie ", vbOKOnly, "Avisar a Sistemas"
           Exit Sub
        End If
                
        If Serie = "" Then
           MsgBox "Falta valor de la serie, Verifique con Sistemas  ", vbOKOnly, "Avisar a Sistemas"
           Exit Sub
        End If
         
        Subtotal = importe / 1.16
        impuesto = Subtotal * 0.16
        
        sqls = " exec sp_FM_facturas @Bodega = " & Bodega & _
                   " ,@AnoFactura = " & Year(Date) & _
                   " ,@Serie = '" & Trim(Serie) & "'" & _
                   " ,@Factura = " & Folio & _
                   " ,@Cliente = " & Val(cliente) & _
                   " ,@Fecha    = '" & Format(Date, "mm/dd/yyyy") & "' " & _
                   " ,@Subtotal = " & Subtotal & _
                   " ,@Iva    = " & impuesto & _
                   " ,@Rubro = 12" & _
                   " ,@BodegaOrigen =  1" & _
                   " ,@Status = 1" & _
                   " ,@StatusImpreso = 0" & _
                   " ,@Pedido = " & Pedido
                   
        cnxBD.Execute sqls, intRegistros
        
        sqls = " EXEC sp_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & Bodega
        sqls = sqls & vbCr & ", @Cliente      = " & Val(cliente)
        sqls = sqls & vbCr & ", @Fecha        = '" & Format(Date, "mm/dd/yyyy") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = 12"
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(Serie) & "'"
        sqls = sqls & vbCr & ", @Refer        = 2"
        sqls = sqls & vbCr & ", @Refer_Apl    = " & Folio
        sqls = sqls & vbCr & ", @CarAbo       = 'C'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & CDbl(Subtotal) + CDbl(impuesto)
        sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Vendedor     = 0"
        sqls = sqls & vbCr & ", @CreditoFac = 'N'"
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = '0'"
        sqls = sqls & vbCr & ", @TipoBon = " & Producto
        
        cnxBD.Execute sqls, intRegistros
  
        i = 0
      sqls = "exec sp_SolicitudesBE_varios @Accion = 'BuscaTipoTarj' , @Pedido = " & Pedido & ", @Producto = " & Producto
        Set rsBD2 = New ADODB.Recordset
        rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly

        Do While Not rsBD2.EOF
            i = i + 1
            sqls = " exec sp_fm_facturas_detalle @Bodega = " & Bodega & _
                   " ,@AnoFactura = " & Year(Date) & _
                   " ,@Serie = '" & Trim(Serie) & "'" & _
                   " ,@Factura = " & Folio & _
                   " ,@Consecutivo   =" & i & _
                   " ,@Concepto = '" & IIf(Left(rsBD2!TipoTarjeta, 1) = "A", "ADICIONALES", IIf(Left(rsBD2!TipoTarjeta, 1) = "R", "REPOSICIONES", IIf(rsBD2!TipoTarjeta = "T", "TITULARES", "X"))) & "'" & _
                   " ,@Cantidad = " & rsBD2!Cantidad & _
                   " ,@PrecioVta = " & rsBD2!costo & _
                   " ,@PorcIva = " & rsBD2!Cantidad * rsBD2!costo * 0.16 & _
                   " ,@BodegaOrigen = 1"
            cnxbdMty.Execute sqls, intRegistros
            rsBD2.MoveNext
        Loop
        
        sqls = " update folios set consecutivo   = " & Folio & _
          " Where Bodega =" & Bodega & _
          " and tipo = 'FCM'"
        
        cnxBD.Execute sqls, intRegistros
        
            
        Call doGenArchFE_OI(Bodega, Trim(Serie), Folio, Folio)
          
        
    End If
    
    sqls = "sp_SolicitudesBE_Varios @Accion = 'Facturado', @Pedido = " & Pedido & ", @Factura = " & Folio
    cnxBD.Execute sqls, intRegistros
    
    Exit Sub
    
err_gral:
       Call doErrorLog(gnBodega, "FACAU", ERR.Number, ERR.Description, Usuario, "frmFacturas.GeneraFacturaOI", 0, sqls)
       MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Genera Factura Tarjetas"
       Resume Next
    
End Sub

Sub GeneraFactura(ByVal cliente As Integer, ByVal Pedido As Integer, ByVal Producto As Integer, ByVal importe As Double)
    
    sqls = " EXEC SP_FACTURABE @BODEGA = 1" & _
       " ,@CLIENTE = " & cliente & _
       " ,@PEDIDO = " & Pedido & _
       " ,@FECHAFAC = '" & Format(Now, "mm/dd/yyyy") & "'" & _
       " ,@Usuario = '" & Usuario & "'" & _
       " ,@Producto =" & Producto
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
        dob_factura = Val(rsBD!Dob)
        Folio = rsBD!folioFac
        Serie = Trim(rsBD!Serie)
        serie2 = Trim(rsBD!serie2)
        Product = Producto
        Producto_factura = Producto
        sqls = "sp_Consultas_BE Null,Null," & Producto & ",1,Null," & Pedido & ",Null,'Autoriza_pedido'"
        cnxBD.Execute sqls
        If dob_factura = 0 Then
           Call doGenArchFE(1, CStr(Serie), Val(Folio), Val(Folio), 7)
        Else
           Call doGenArchFE(1, CStr(Serie), Val(Folio) - 1, Val(Folio), 7)
        End If
    Else
      MsgBox "Error al generar la factura", vbCritical, "Error..."
      Exit Sub
    End If
 
End Sub
